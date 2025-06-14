import zipfile
import xml.etree.ElementTree as ET
import os
import argparse
import shutil
import tempfile
import re
import math

# --- Helper Functions for Geometry and Transforms (existing ones are mostly fine) ---
def parse_transform_matrix(transform_str):
    if not transform_str: return (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
    try:
        parts = list(map(float, transform_str.split()));
        if len(parts) == 6: return tuple(parts)
    except ValueError: pass
    return (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)

def parse_geometric_bounds(bounds_str):
    if not bounds_str: return (0.0, 0.0, 0.0, 0.0)
    try:
        parts = list(map(float, bounds_str.split()));
        if len(parts) == 4: return tuple(parts) # y1, x1, y2, x2
    except ValueError: pass
    return (0.0, 0.0, 0.0, 0.0)

def multiply_matrices(m1, m2):
    a1, b1, c1, d1, tx1, ty1 = m1; a2, b2, c2, d2, tx2, ty2 = m2
    return (a1*a2+c1*b2, b1*a2+d1*b2, a1*c2+c1*d2, b1*c2+d1*d2, a1*tx2+c1*ty2+tx1, b1*tx2+d1*ty2+ty1)

def apply_transform_to_point(x, y, matrix):
    a, b, c, d, tx, ty = matrix
    return (a*x+c*y+tx, b*x+d*y+ty)

def get_global_corners(local_bounds, global_item_matrix):
    y1, x1, y2, x2 = local_bounds
    return [
        apply_transform_to_point(x1, y1, global_item_matrix), apply_transform_to_point(x2, y1, global_item_matrix),
        apply_transform_to_point(x1, y2, global_item_matrix), apply_transform_to_point(x2, y2, global_item_matrix)
    ]

def get_axis_aligned_bounding_box(corners):
    if not corners: return {"x1":0.0,"y1":0.0,"x2":0.0,"y2":0.0}
    all_x=[p[0] for p in corners]; all_y=[p[1] for p in corners]
    return {"x1":min(all_x),"y1":min(all_y),"x2":max(all_x),"y2":max(all_y)}

def get_item_center(global_aabb):
    return ((global_aabb["x1"]+global_aabb["x2"])/2.0, (global_aabb["y1"]+global_aabb["y2"])/2.0)

def is_point_in_rect(px, py, rect_y1, rect_x1, rect_y2, rect_x2):
    """Checks if a point (px, py) is within a rectangle [x1,y1,x2,y2]."""
    return (rect_x1 <= px <= rect_x2) and \
           (rect_y1 <= py <= rect_y2)

class PageGeometricInfo:
    def __init__(self, self_id, name, page_matrix_str, page_bounds_str):
        self.id=self_id; self.name=name
        self.matrix=parse_transform_matrix(page_matrix_str)
        self.local_bounds=parse_geometric_bounds(page_bounds_str)
        pgc=get_global_corners(self.local_bounds,self.matrix)
        self.global_aabb=get_axis_aligned_bounding_box(pgc)
        print(f"  Page '{self.name}' (ID: {self.id}): Global AABB=(x: {self.global_aabb['x1']:.2f}-{self.global_aabb['x2']:.2f}, y: {self.global_aabb['y1']:.2f}-{self.global_aabb['y2']:.2f})")

# NEW HELPER FUNCTION for TextFrame bounds
def get_local_bounds_from_path_geometry(tf_element):
    """
    Attempts to parse <PathGeometry> from a TextFrame or other element
    to determine its local bounding box from its anchor points.
    Returns (y1, x1, y2, x2) or None if path geometry can't be parsed or is empty.
    """
    # Common path for PathGeometry within TextFrames is under Properties
    path_point_array = tf_element.find('.//{*}Properties/{*}PathGeometry/{*}GeometryPathType/{*}PathPointArray')
    if path_point_array is None:
        # Fallback: Try finding PathGeometry not nested under Properties (less common for TFs)
        path_point_array = tf_element.find('.//{*}PathGeometry/{*}GeometryPathType/{*}PathPointArray')
        if path_point_array is None:
            return None

    all_x = []
    all_y = []
    # PathPointType elements are direct children of PathPointArray
    for path_point_type in path_point_array.findall('{*}PathPointType'):
        anchor_str = path_point_type.get("Anchor")
        if anchor_str:
            try:
                coords = list(map(float, anchor_str.split()))
                if len(coords) == 2:
                    all_x.append(coords[0])  # First value is X
                    all_y.append(coords[1])  # Second value is Y
            except ValueError:
                # print(f"Warning: Could not parse Anchor coordinates: {anchor_str}")
                continue
    
    if not all_x or not all_y:
        return None # No valid anchor points found
        
    # y1 (min Y), x1 (min X), y2 (max Y), x2 (max X)
    return (min(all_y), min(all_x), max(all_y), max(all_x))

# --- Main Extraction Logic ---
def extract_idml(idml_path, extract_to):
    try:
        with zipfile.ZipFile(idml_path, 'r') as zf: zf.extractall(extract_to)
        print(f"✅ Successfully extracted IDML to {extract_to}"); return True
    except Exception as e: print(f"❌ Extraction failed: {e}"); return False

def get_story_text(story_path):
    text_content_segments = []
    try:
        tree = ET.parse(story_path)
        root = tree.getroot()
        for element in root.iter():
            tag_name = element.tag.split('}')[-1]
            if tag_name == 'Content':
                if element.text: text_content_segments.append(element.text)
            elif tag_name == 'Br':
                text_content_segments.append('\n')
        full_text = "".join(text_content_segments)
        full_text = re.sub(r'\s*\n\s*', '\n', full_text)
        full_text = re.sub(r'\n{2,}', '\n', full_text).strip()
        if not full_text and os.path.exists(story_path): # Added os.path.exists to avoid logging for non-existent stories
            print(f"    ℹ️ Story {os.path.basename(story_path)}: No text content extracted (story might be empty or structure not recognized for text).")
        return full_text
    except ET.ParseError as e:
        print(f"    ❌ ERROR PARSING STORY XML {os.path.basename(story_path)}: {e}")
        return ""
    except Exception as e:
        print(f"    ❌ UNEXPECTED ERROR in get_story_text for {os.path.basename(story_path)}: {e}")
        return ""

def find_page_for_item_center(item_center_x, item_center_y, geometric_pages):
    for page_info in geometric_pages:
        if is_point_in_rect(item_center_x, item_center_y,
                              page_info.global_aabb["y1"], page_info.global_aabb["x1"],
                              page_info.global_aabb["y2"], page_info.global_aabb["x2"]):
            return page_info.id
    return None

def process_spread_element_recursively(element, parent_element, current_accumulated_matrix,
                                       geometric_pages, pages_content_map,
                                       stories_dir, story_cache, depth=0):
    indent = "  " * (depth + 2)
    element_tag_local = element.tag.split('}')[-1]
    element_id = element.get("Self", "UnknownID")

    item_local_matrix_str = element.get("ItemTransform")
    item_local_matrix = parse_transform_matrix(item_local_matrix_str)
    item_global_matrix = multiply_matrices(current_accumulated_matrix, item_local_matrix)

    determined_local_bounds = (0.0, 0.0, 0.0, 0.0) # Default

    if element_tag_local == 'Image':
        graphic_bounds_prop = element.find('.//{*}Properties/{*}GraphicBounds')
        if graphic_bounds_prop is not None:
            try:
                gb_l=float(graphic_bounds_prop.get("Left","0")); gb_t=float(graphic_bounds_prop.get("Top","0"))
                gb_r=float(graphic_bounds_prop.get("Right","0")); gb_b=float(graphic_bounds_prop.get("Bottom","0"))
                determined_local_bounds = (gb_t, gb_l, gb_b, gb_r) # y1, x1, y2, x2
                # print(f"{indent}ℹ️ Image ID: {element_id} using GraphicBounds: {determined_local_bounds}")
            except (ValueError, TypeError) as e: print(f"{indent}⚠️ Error parsing GraphicBounds for Image {element_id}: {e}")
        if determined_local_bounds == (0.0,0.0,0.0,0.0): # Fallback for Image
            item_lbs_img = element.get("GeometricBounds")
            if item_lbs_img: determined_local_bounds = parse_geometric_bounds(item_lbs_img)

    elif element_tag_local == "TextFrame":
        path_geom_bounds = get_local_bounds_from_path_geometry(element)
        if path_geom_bounds:
            determined_local_bounds = path_geom_bounds
            print(f"{indent}ℹ️ TextFrame ID: {element_id} using PathGeometry bounds: {determined_local_bounds}")
        else: # Fallback to GeometricBounds attribute if PathGeometry fails or not present
            item_lbs_tf = element.get("GeometricBounds")
            if item_lbs_tf: determined_local_bounds = parse_geometric_bounds(item_lbs_tf)
            else: print(f"{indent}ℹ️ TextFrame ID: {element_id} no PathGeometry or GeometricBounds. Bounds will be 0.")
                
    else: # For Rectangle, Group, Oval, Polygon, etc.
        item_lbs_other = element.get("GeometricBounds")
        determined_local_bounds = parse_geometric_bounds(item_lbs_other)
    
    item_global_aabb = {"x1":0.0,"y1":0.0,"x2":0.0,"y2":0.0}
    item_center_x, item_center_y = item_global_matrix[4], item_global_matrix[5]

    if determined_local_bounds != (0.0, 0.0, 0.0, 0.0):
        item_global_corners = get_global_corners(determined_local_bounds, item_global_matrix)
        item_global_aabb = get_axis_aligned_bounding_box(item_global_corners)
        item_center_x, item_center_y = get_item_center(item_global_aabb)
    
    assigned_page_id = find_page_for_item_center(item_center_x, item_center_y, geometric_pages)
    
    if element_tag_local == "TextFrame":
        story_id = element.get("ParentStory")
        # print(f"{indent}Processing TextFrame ID: {element_id}, StoryID: {story_id}, Center: ({item_center_x:.1f},{item_center_y:.1f}), PageID: {assigned_page_id}")
        if assigned_page_id and story_id:
            if story_id not in story_cache:
                story_fn = f"Story_{story_id}.xml"; story_fp = os.path.join(stories_dir, story_fn)
                if os.path.exists(story_fp): story_cache[story_id] = get_story_text(story_fp)
                else: story_cache[story_id] = ""; print(f"{indent}  ❌ Story file {story_fn} NOT FOUND for TF {element_id}")
            
            text_content = story_cache.get(story_id, "")
            page_data = pages_content_map.get(assigned_page_id)
            # print(f"{indent}  For TF {element_id} on Page {assigned_page_id}: PageDataOk: {page_data is not None}. TextLen: {len(text_content)}. Story: {story_id}")
            if page_data:
                is_tf_added = any(tf['text_frame_id'] == element_id for tf in page_data["texts"])
                if not is_tf_added:
                    if text_content:
                        page_data["texts"].append({
                            "story_id":story_id, "text_frame_id":element_id, "content":text_content,
                            "global_bounds":item_global_aabb, "item_transform":item_local_matrix_str 
                        })
                        print(f"{indent}  ✅ Added TextFrame ID: {element_id} (Story: {story_id}) to Page ID: {assigned_page_id}.")
                    elif story_id in story_cache:
                        print(f"{indent}  ℹ️ TF {element_id} (Story: {story_id}) has empty text_content from cache, not adding.")
        elif story_id:
             print(f"{indent}⚠️ TF {element_id} (Story: {story_id}) at C:({item_center_x:.1f},{item_center_y:.1f}) NOT assigned to page.")

    elif element_tag_local == "Image":
        if assigned_page_id:
            link_el = element.find('.//{*}Link')
            if link_el is not None:
                uri = link_el.get("LinkResourceURI")
                if uri:
                    page_data = pages_content_map.get(assigned_page_id)
                    if page_data:
                        cont_tag=parent_element.tag.split('}')[-1] if parent_element else "Unk"
                        cont_id=parent_element.get("Self","Unk") if parent_element else "Unk"
                        if not any(i['uri']==uri and i['image_element_id']==element_id for i in page_data["images"]):
                            page_data["images"].append({
                                "uri":uri,"image_element_id":element_id, "container_element_tag":cont_tag,
                                "container_element_id":cont_id, "global_bounds":item_global_aabb,
                                "item_transform":item_local_matrix_str 
                            })
                            # print(f"{indent}🖼️ Added Image Link: '{uri}' (ID: {element_id}) to Page ID: {assigned_page_id}.")


    if element_tag_local in ["Group", "Rectangle", "Oval", "Polygon"]: # Ensure containers are recursed
        for child_element in element:
            process_spread_element_recursively(child_element, element, item_global_matrix,
                                               geometric_pages, pages_content_map,
                                               stories_dir, story_cache, depth + 1)

# --- get_page_content_from_spread, main, and other helpers remain the same ---
def get_page_content_from_spread(spread_path, stories_dir, story_cache):
    try:
        tree = ET.parse(spread_path); root = tree.getroot(); spread_element = root
        # Attempt to find the actual <Spread> content element if root is a wrapper
        if not spread_element.tag.endswith('Spread') or not any(c.tag.split('}')[-1]=='Page' for c in spread_element):
            candidate = spread_element.find('{*}Spread') # Common for idPkg:Spread wrapper
            if candidate is not None and any(c.tag.split('}')[-1]=='Page' for c in candidate):
                spread_element = candidate
            else: # Try finding any Spread tag if the first check failed broadly
                found_spreads = [el for el in root.iter() if el.tag.endswith('Spread') and any(c.tag.split('}')[-1]=='Page' for c in el)]
                if found_spreads:
                    spread_element = found_spreads[0] # Take the first one with Page children
                else:
                    print(f"  ❌ No main <Spread> with <Page> children in {os.path.basename(spread_path)}"); return {}
        # print(f"  Processing Spread Element: <{spread_element.tag.split('}')[-1]} Self='{spread_element.get('Self')}'>") # Less verbose
    except Exception as e: print(f"  ❌ Error parsing {spread_path}: {e}"); return {}

    current_spread_pages_content = {}; geometric_pages = []
    for page_el in spread_element.findall('.//{*}Page'): # Ensure we search within the identified spread_element
        pid=page_el.get("Self"); name=page_el.get("Name",f"UnkPage_{pid}")
        if not pid: continue
        page_info=PageGeometricInfo(pid,name,page_el.get("ItemTransform"),page_el.get("GeometricBounds"))
        geometric_pages.append(page_info)
        if pid not in current_spread_pages_content:
             current_spread_pages_content[pid] = {"name":name,"images":[],"texts":[]}
    if not geometric_pages: print(f"  ℹ️ No Page elements found in {os.path.basename(spread_path)}."); return {}

    sbm_str=spread_element.get("ItemTransform"); sbm=parse_transform_matrix(sbm_str)
    for child_el in spread_element: # Iterate direct children of the content <Spread>
        ct_local = child_el.tag.split('}')[-1]
        if ct_local in ["Page","FlattenerPreference","Properties"]: continue # Properties of Spread itself
        process_spread_element_recursively(child_el,spread_element,sbm,geometric_pages,
                                           current_spread_pages_content,stories_dir,story_cache,0)
    return current_spread_pages_content

def find_spread_files(spreads_dir):
    if not os.path.isdir(spreads_dir): return []
    return [os.path.join(spreads_dir, f) for f in os.listdir(spreads_dir) if f.startswith("Spread_") and f.endswith(".xml")]

def main():
    parser = argparse.ArgumentParser(description="Extract content per page from IDML.")
    parser.add_argument("idml_file", help="Path to .idml file")
    args = parser.parse_args()

    temp_dir_base = os.path.join(tempfile.gettempdir(), "idml_extractor_debug")
    os.makedirs(temp_dir_base, exist_ok=True) 
    temp_dir = tempfile.mkdtemp(prefix="idml_", dir=temp_dir_base)
    
    print(f"⏳ Extracting '{args.idml_file}' to: {temp_dir}")
    if not extract_idml(args.idml_file, temp_dir):
        if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
        return
    
    spreads_dir = os.path.join(temp_dir, "Spreads")
    stories_dir = os.path.join(temp_dir, "Stories")
    
    if not os.path.isdir(spreads_dir):
        print(f"❌ Critical: 'Spreads' folder missing at '{spreads_dir}'.")
        if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
        return

    all_page_data_by_id = {} 
    story_cache = {} 

    spread_files = find_spread_files(spreads_dir)
    if not spread_files:
        print(f"❌ No spread files found in {spreads_dir}.")
        if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
        return
    
    for spread_file_path in spread_files:
        spread_name = os.path.basename(spread_file_path)
        print(f"\n📄 Processing Spread File: {spread_name}")
        content_from_this_spread = get_page_content_from_spread(spread_file_path, stories_dir, story_cache)
        for pid, pdata_in_spread in content_from_this_spread.items():
            if pid not in all_page_data_by_id:
                all_page_data_by_id[pid] = pdata_in_spread
            else: 
                for img_item in pdata_in_spread["images"]:
                    if not any(ex_img['uri'] == img_item['uri'] and ex_img['image_element_id'] == img_item['image_element_id'] 
                               for ex_img in all_page_data_by_id[pid]["images"]):
                        all_page_data_by_id[pid]["images"].append(img_item)
                current_tf_ids = {t['text_frame_id'] for t in all_page_data_by_id[pid]["texts"]}
                for text_item in pdata_in_spread["texts"]:
                    if text_item['text_frame_id'] not in current_tf_ids:
                        all_page_data_by_id[pid]["texts"].append(text_item)
                        current_tf_ids.add(text_item['text_frame_id'])

    print("\n\n--- ✨ Consolidated Content Per Page ✨ ---")
    if not all_page_data_by_id:
        print("No content found on any pages.")
    else:
        final_sorted_page_ids = sorted(all_page_data_by_id.keys(), 
                                       key=lambda pid_key: all_page_data_by_id[pid_key]["name"])
        for pid in final_sorted_page_ids:
            pdata = all_page_data_by_id[pid]
            print(f"\n--- Page \"{pdata['name']}\" (ID: {pid}) ---")
            if pdata["texts"]:
                print("📝 Texts:")
                for t_item in pdata["texts"]:
                    # content_prev = (t_item['content'][:70] + '...') if len(t_item['content']) > 70 else t_item['content']
                    bounds = t_item['global_bounds']
                    print(f"  - Story: {t_item['story_id']} (Frame: {t_item['text_frame_id']}) "
                          f"Bounds: x1={bounds['x1']:.1f},y1={bounds['y1']:.1f},x2={bounds['x2']:.1f},y2={bounds['y2']:.1f}")
                    print(f"    Content: \"{t_item}\"")
            else: print("📝 Texts: None")
            if pdata["images"]:
                print("🖼 Images:")
                for img_item in pdata["images"]:
                    bounds = img_item['global_bounds']
                    print(f"  - URI: {img_item['uri']} (ImageElem: {img_item['image_element_id']}, "
                          f"Container: <{img_item['container_element_tag']} ID:{img_item['container_element_id']}>)")
                    print(f"    Bounds: x1={bounds['x1']:.1f},y1={bounds['y1']:.1f},x2={bounds['x2']:.1f},y2={bounds['y2']:.1f}")
            else: print("🖼 Images: None")

    print("\n✅ Done. Cleaning up temporary directory...")
    try:
        shutil.rmtree(temp_dir)
        print(f"🗑️ Temporary directory {temp_dir} removed.")
    except Exception as e:
        print(f"⚠️ Error removing {temp_dir}: {e}. Please remove manually.")

if __name__ == "__main__":
    main()