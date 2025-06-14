import zipfile
import xml.etree.ElementTree as ET
import os
import argparse
import shutil
import tempfile
import re

def extract_idml(idml_path, extract_to):
    """Extracts the contents of an IDML file (which is a zip archive)."""
    try:
        with zipfile.ZipFile(idml_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)
        return True
    except Exception as e:
        print(f"Extraction failed: {e}")
        return False

def get_story_text(story_path):
    """Parses a Story XML file and extracts all text content."""
    text_content = []
    try:
        tree = ET.parse(story_path)
        root = tree.getroot()
        # Iterate through elements in document order to try and preserve paragraph/break sequence
        for element in root.findall('.//{*}XMLElement/*'): # A common parent for story content
            if element.tag.endswith('Content'):
                if element.text:
                    text_content.append(element.text) # .strip() could remove leading/trailing spaces of segments
            elif element.tag.endswith('Br'):
                text_content.append('\n')
        
        # Fallback or additional method for broader content grabbing if above is too restrictive
        if not text_content:
            for content_element in root.findall('.//{*}Content'):
                if content_element.text:
                    text_content.append(content_element.text)
            # This Br handling might add too many newlines if also caught by XMLElement iteration
            # Consider if Br should only be handled once or if Story structure guarantees no overlap.
            # For now, let's assume it's okay and will be cleaned up.
            for br_element in root.findall('.//{*}Br'): # Br might not be under XMLElement
                 text_content.append('\n')

        full_text = "".join(text_content) # Don't filter None, join handles it. Avoid stripping individual segments.
        
        # Consolidate multiple newlines and strip leading/trailing whitespace from the whole story
        full_text = re.sub(r'\s*\n\s*', '\n', full_text) # Normalize newlines and surrounding spaces
        full_text = re.sub(r'\n+', '\n', full_text).strip() # Consolidate multiple newlines
        return full_text
    except ET.ParseError as e:
        print(f"Error parsing story XML {story_path}: {e}")
        return ""
    except Exception as e:
        print(f"Unexpected error getting story text from {story_path}: {e}")
        return ""


def get_page_content_from_spread(spread_path, stories_dir, story_cache):
    """Parses a spread XML and extracts image/text grouped by Page ID."""
    try:
        tree = ET.parse(spread_path)
        root = tree.getroot()
    except ET.ParseError as e:
        print(f"Error parsing spread XML {spread_path}: {e}")
        return {}

    pages_content = {} # Key: Page Self ID, Value: {name, images, texts}

    # Map page ID to page Name and initialize content lists
    for page_element in root.findall('.//{*}Page'):
        pid = page_element.attrib.get("Self")
        if not pid:
            continue
        name = page_element.attrib.get("Name", f"UnnamedPage_{pid}")
        pages_content[pid] = {"name": name, "images": [], "texts": []}
    
    # Process TextFrames
    for tf in root.findall('.//{*}TextFrame'):
        # ParentPage is an attribute on TextFrame pointing to the Page's Self ID
        page_id = tf.attrib.get("ParentPage")
        story_id = tf.attrib.get("ParentStory") # Points to Story Self ID (e.g., "u123")
        print(f'page_id = {page_id}   story_id = {story_id}')
        
        if page_id in pages_content and story_id:
            if story_id not in story_cache:
                story_filename = f"Story_{story_id}.xml" # Assuming this naming convention
                story_file_path = os.path.join(stories_dir, story_filename)
                if os.path.exists(story_file_path):
                    story_cache[story_id] = get_story_text(story_file_path)
                else:
                    # print(f"Warning: Story file {story_filename} not found for story ID {story_id}")
                    story_cache[story_id] = "" # Cache empty if not found
            
            # Avoid adding the same story multiple times if multiple text frames on the page use it
            # Check if this story content is already added for this page_id
            is_story_already_added = any(
                t['story_id'] == story_id for t in pages_content[page_id]["texts"]
            )
            if not is_story_already_added and story_cache[story_id]: # Add if not present and has content
                pages_content[page_id]["texts"].append({
                    "story_id": story_id,
                    "content": story_cache[story_id]
                })

    # Process Images - Images are inside graphic frames like Rectangle, Oval, Polygon
    # These frames have ParentPage.
    for frame_type_name in ['Rectangle', 'Oval', 'Polygon', 'Group']: # Add other containers if necessary
        for frame in root.findall(f'.//{{*}}{frame_type_name}'):
            page_id = frame.attrib.get("ParentPage")
            if page_id in pages_content:
                # Find Image elements within this frame (can be nested)
                for img in frame.findall('.//{*}Image'):
                    link = img.find('.//{*}Link')
                    if link is not None: # Check if link element exists
                        uri = link.attrib.get("LinkResourceURI")
                        if uri:
                            # Avoid duplicate URIs for the same page
                            if uri not in pages_content[page_id]["images"]:
                                pages_content[page_id]["images"].append(uri)
    return pages_content

def find_spread_files(spreads_dir):
    """Finds all spread XML files in the Spreads directory."""
    return [os.path.join(spreads_dir, f) for f in os.listdir(spreads_dir) if f.startswith("Spread_") and f.endswith(".xml")]

def main():
    parser = argparse.ArgumentParser(description="Extract content per page from IDML.")
    parser.add_argument("idml_file", help="Path to .idml file")
    args = parser.parse_args()

    temp_dir = tempfile.mkdtemp(prefix="idml_")
    print(f"⏳ Extracting to: {temp_dir}")
    if not extract_idml(args.idml_file, temp_dir):
        shutil.rmtree(temp_dir)
        return

    spreads_dir = os.path.join(temp_dir, "Spreads")
    stories_dir = os.path.join(temp_dir, "Stories")

    if not os.path.isdir(spreads_dir) or not os.path.isdir(stories_dir):
        print("❌ Spreads or Stories folder missing from the IDML package.")
        shutil.rmtree(temp_dir)
        return

    # This will store all page data, keyed by unique Page Self ID
    all_page_data_by_id = {}
    story_cache = {} # Cache stories to avoid re-reading and parsing

    spread_files = find_spread_files(spreads_dir)
    if not spread_files:
        print("❌ No spread files found in the Spreads directory.")
        shutil.rmtree(temp_dir)
        return

    for spread_file in spread_files:
        spread_name = os.path.basename(spread_file)
        print(f"📄 Processing Spread: {spread_name}")
        
        # Get content for pages within this specific spread
        content_from_spread = get_page_content_from_spread(spread_file, stories_dir, story_cache)
        
        for pid, pdata_in_spread in content_from_spread.items():
            if pid not in all_page_data_by_id:
                # Initialize page data if this page ID is encountered for the first time
                all_page_data_by_id[pid] = {
                    "name": pdata_in_spread["name"], 
                    "images": [], 
                    "texts": []
                }
            
            # Aggregate images, ensuring no duplicates
            for img_uri in pdata_in_spread["images"]:
                if img_uri not in all_page_data_by_id[pid]["images"]:
                    all_page_data_by_id[pid]["images"].append(img_uri)
            
            # Aggregate texts, ensuring no duplicate story objects for the same page
            current_story_ids_on_page = {t['story_id'] for t in all_page_data_by_id[pid]["texts"]}
            for text_item in pdata_in_spread["texts"]:
                if text_item['story_id'] not in current_story_ids_on_page:
                    all_page_data_by_id[pid]["texts"].append(text_item)
                    # No need to add to current_story_ids_on_page here as it's rebuilt on next page if needed

    print("\n\n--- ✨ Consolidated Content Per Page ✨ ---")
    if not all_page_data_by_id:
        print("No content found on any pages.")

    # Sort pages for consistent output, e.g., by name.
    # True document order might require parsing designmap.xml
    sorted_page_ids = sorted(all_page_data_by_id.keys(), key=lambda page_id_key: all_page_data_by_id[page_id_key]["name"])

    for pid in sorted_page_ids:
        pdata = all_page_data_by_id[pid]
        print(f"\n--- Page \"{pdata['name']}\" (ID: {pid}) ---")
        
        if pdata["texts"]:
            print("📝 Texts:")
            for t in pdata["texts"]:
                # Limit long text preview for conciseness in terminal
                content_preview = (t['content'][:150] + '...') if len(t['content']) > 150 else t['content']
                print(f"  Story ID: {t['story_id']}\n  Content: {content_preview}\n")
        else:
            print("📝 Texts: None")
            
        if pdata["images"]:
            print("🖼 Images:")
            for img_uri in pdata["images"]:
                print(f"  - {img_uri}")
        else:
            print("🖼 Images: None")

    print("\n✅ Done. Cleaning up...")
    shutil.rmtree(temp_dir)

if __name__ == "__main__":
    main()