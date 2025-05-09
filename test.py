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
        # print(f"Successfully extracted {idml_path} to {extract_to}")
        return True
    except zipfile.BadZipFile:
        print(f"Error: {idml_path} is not a valid zip file or is corrupted.")
        return False
    except FileNotFoundError:
        print(f"Error: IDML file not found at {idml_path}")
        return False
    except Exception as e:
        print(f"An error occurred during extraction: {e}")
        return False

def get_story_text(story_path):
    """Parses a Story XML file and extracts all text content."""
    text_content = []
    try:
        tree = ET.parse(story_path)
        root = tree.getroot()
        # Find all Content tags anywhere in the story file
        # Stories can have complex structures (ParagraphStyleRange, CharacterStyleRange, etc.)
        for content_element in root.findall('.//{*}Content'):
            if content_element.text:
                text_content.append(content_element.text.strip())
        # Also find Br tags for line breaks
        for br_element in root.findall('.//{*}Br'):
             text_content.append('\n') # Represent break as newline

        # Join the extracted text pieces. Handle potential None values just in case.
        full_text = "".join(filter(None, text_content))
        # Clean up multiple consecutive newlines that might result from <Br/> tags
        full_text = re.sub(r'\n+', '\n', full_text).strip()
        return full_text
    except ET.ParseError:
        print(f"Warning: Could not parse Story file: {story_path}")
        return ""
    except FileNotFoundError:
        print(f"Warning: Story file not found: {story_path}")
        return ""
    except Exception as e:
        print(f"An error occurred parsing story {story_path}: {e}")
        return ""

def get_spread_content(spread_path, stories_dir):
    """Parses a Spread XML file to find image links and story references."""
    image_links = set()
    story_texts = {} # Dictionary to store {story_id: text}
    referenced_story_ids = set()

    try:
        tree = ET.parse(spread_path)
        root = tree.getroot()

        # Find all Link elements with a LinkResourceURI (likely images)
        # These can be nested within Image, Rectangle, Group etc.
        for link_element in root.findall('.//{*}Link[@LinkResourceURI]'):
            uri = link_element.get('LinkResourceURI')
            if uri:
                image_links.add(uri)

        # Find all TextFrame elements with a ParentStory attribute
        for text_frame in root.findall('.//{*}TextFrame[@ParentStory]'):
            story_id = text_frame.get('ParentStory')
            if story_id:
                referenced_story_ids.add(story_id)

        # Extract text from referenced stories
        for story_id in referenced_story_ids:
            story_filename = f"Story_{story_id}.xml"
            story_path = os.path.join(stories_dir, story_filename)
            story_texts[story_id] = get_story_text(story_path)

        return image_links, story_texts

    except ET.ParseError:
        print(f"Error: Could not parse Spread file: {spread_path}")
        return set(), {}
    except FileNotFoundError:
        print(f"Error: Spread file not found: {spread_path}")
        return set(), {}
    except Exception as e:
        print(f"An error occurred parsing spread {spread_path}: {e}")
        return set(), {}

def find_spread_files(spreads_dir):
    """Finds all Spread XML files in the Spreads directory."""
    spread_files = []
    if not os.path.isdir(spreads_dir):
        print(f"Error: Spreads directory not found: {spreads_dir}")
        return []
    for filename in os.listdir(spreads_dir):
        if filename.startswith("Spread_") and filename.endswith(".xml"):
            spread_files.append(os.path.join(spreads_dir, filename))
    return spread_files

def main():
    parser = argparse.ArgumentParser(description="Extract text and image links from an Adobe IDML file.")
    parser.add_argument("idml_file", help="Path to the .idml file.")
    parser.add_argument("-s", "--spread", help="Specific spread file name to process (e.g., Spread_u160a.xml). If not provided, process all spreads.", default=None)

    args = parser.parse_args()

    # Create a temporary directory for extraction
    temp_dir = tempfile.mkdtemp(prefix="idml_extract_")
    print(f"Extracting IDML to temporary directory: {temp_dir}")

    if not extract_idml(args.idml_file, temp_dir):
        shutil.rmtree(temp_dir) # Clean up temp dir on extraction failure
        return
    
    spreads_dir = os.path.join(temp_dir, "Spreads")
    stories_dir = os.path.join(temp_dir, "Stories")

    if not os.path.isdir(stories_dir):
         print(f"Error: Stories directory not found in the extracted IDML structure: {stories_dir}")
         shutil.rmtree(temp_dir)
         return

    spreads_to_process = []
    if args.spread:
        # Process only the specified spread
        specific_spread_path = os.path.join(spreads_dir, args.spread)
        if os.path.exists(specific_spread_path):
            spreads_to_process.append(specific_spread_path)
        else:
            print(f"Error: Specified spread file not found: {specific_spread_path}")
    else:
        # Process all spreads found
        spreads_to_process = find_spread_files(spreads_dir)
        if not spreads_to_process:
             print("No spread files found to process.")


    all_images = set()
    all_texts = {} # Using spread filename as key: {spread_filename: {story_id: text}}

    # Process the selected spread files
    for spread_path in spreads_to_process:
        print(f"\n--- Processing Spread: {os.path.basename(spread_path)} ---")
        image_links, story_texts = get_spread_content(spread_path, stories_dir)

        if image_links:
            print("\nImage Links Found:")
            for link in sorted(list(image_links)):
                print(f"- {link}")
            all_images.update(image_links) # Add to overall set

        if story_texts:
            print("\nText Content Found:")
            spread_filename = os.path.basename(spread_path)
            all_texts[spread_filename] = {}
            for story_id, text in story_texts.items():
                 if text: # Only print if text was actually extracted
                    print(f"\n[From Story_{story_id}.xml]:")
                    print(text)
                    all_texts[spread_filename][story_id] = text
                 else:
                     print(f"\n[No text content extracted from Story_{story_id}.xml]")
        else:
             print("\nNo text frames referencing stories found in this spread.")


    # Optional: Print summary at the end
    print("\n--- Extraction Summary ---")
    if all_images:
        print(f"\nTotal Unique Image Links Found Across Processed Spreads ({len(all_images)}):")
        for link in sorted(list(all_images)):
            print(f"- {link}")
    else:
        print("\nNo image links found in the processed spreads.")

    if all_texts:
        print(f"\nTotal Text Stories Found Across Processed Spreads:")
        for spread_file, stories in all_texts.items():
             print(f"  Spread: {spread_file} ({len(stories)} stories with content)")
    else:
        print("\nNo text content found in the processed spreads.")


    # Clean up the temporary directory
    try:
        print(f"\nCleaning up temporary directory: {temp_dir}")
        shutil.rmtree(temp_dir)
    except Exception as e:
        print(f"Warning: Could not remove temporary directory {temp_dir}: {e}")

if __name__ == "__main__":
    main()