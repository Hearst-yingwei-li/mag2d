import zipfile
import xml.etree.ElementTree as ET
import os
import argparse
import shutil
import tempfile
import re
import math
import json

# --- Helper Functions for Geometry and Transforms ---
def parse_transform_matrix(transform_str):
    if not transform_str: return (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
    try:
        parts = list(map(float, transform_str.split()))
        if len(parts) == 6: return tuple(parts)
    except ValueError: pass
    return (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)

def parse_geometric_bounds(bounds_str):
    if not bounds_str: return (0.0, 0.0, 0.0, 0.0)
    try:
        parts = list(map(float, bounds_str.split()))
        if len(parts) == 4: return tuple(parts) # y1, x1, y2, x2
    except ValueError: pass
    return (0.0, 0.0, 0.0, 0.0)

def multiply_matrices(m1, m2):
    a1, b1, c1, d1, tx1, ty1 = m1; a2, b2, c2, d2, tx2, ty2 = m2
    return (a1*a2+c1*b2, b1*a2+d1*b2, a1*c2+c1*d2, b1*c2+d1*d2, a1*tx2+c1*ty2+tx1, b1*tx2+d1*ty2+ty1)

def apply_transform_to_point(x, y, matrix):
    a, b, c, d, tx, ty = matrix
    return (a*x+c*y+tx, b*x+d*y+ty)

def get_global_corners(local_bounds, item_global_matrix): # Renamed from global_item_matrix for clarity
    y1, x1, y2, x2 = local_bounds
    return [
        apply_transform_to_point(x1, y1, item_global_matrix), apply_transform_to_point(x2, y1, item_global_matrix),
        apply_transform_to_point(x1, y2, item_global_matrix), apply_transform_to_point(x2, y2, item_global_matrix)
    ]

def get_axis_aligned_bounding_box(corners):
    if not corners: return {"x1":0.0,"y1":0.0,"x2":0.0,"y2":0.0}
    all_x=[p[0] for p in corners]; all_y=[p[1] for p in corners]
    return {"x1":min(all_x),"y1":min(all_y),"x2":max(all_x),"y2":max(all_y)}

def get_item_center(global_aabb):
    return ((global_aabb["x1"]+global_aabb["x2"])/2.0, (global_aabb["y1"]+global_aabb["y2"])/2.0)

def is_point_in_rect(px, py, rect_y1, rect_x1, rect_y2, rect_x2): # CORRECTED
    return (rect_x1 <= px <= rect_x2) and \
           (rect_y1 <= py <= rect_y2)

class PageGeometricInfo:
    def __init__(self, self_id, name, page_local_transform_str, page_local_bounds_str, spread_base_transform_matrix): # MODIFIED
        self.id = self_id
        self.name = name
        
        page_local_matrix = parse_transform_matrix(page_local_transform_str)
        self.local_bounds = parse_geometric_bounds(page_local_bounds_str) # These are page's own geometric bounds
        
        # Calculate the page's true global matrix by applying the spread's base transform to the page's local transform
        self.global_matrix = multiply_matrices(spread_base_transform_matrix, page_local_matrix)
        
        # Calculate global AABB of the page using the page's local bounds and its true global_matrix
        page_global_corners = get_global_corners(self.local_bounds, self.global_matrix)
        self.global_aabb = get_axis_aligned_bounding_box(page_global_corners)
        
        print(f"  Page '{self.name}' (ID: {self.id}): "
              # f"LocalTransform={page_local_matrix}, SpreadBaseM={spread_base_transform_matrix}, "
              f"FinalGlobalMatrix={tuple(f'{x:.2f}' for x in self.global_matrix)}, "
              f"Global AABB=(x1: {self.global_aabb['x1']:.2f}- x2: {self.global_aabb['x2']:.2f}, "
              f"y1: {self.global_aabb['y1']:.2f}- y2:{self.global_aabb['y2']:.2f})")


def get_local_bounds_from_path_geometry(tf_element):
    path_point_array = tf_element.find('.//{*}Properties/{*}PathGeometry/{*}GeometryPathType/{*}PathPointArray')
    if path_point_array is None:
        path_point_array = tf_element.find('.//{*}PathGeometry/{*}GeometryPathType/{*}PathPointArray')
        if path_point_array is None: return None
    all_x=[]; all_y=[]
    for path_point_type in path_point_array.findall('{*}PathPointType'):
        anchor_str = path_point_type.get("Anchor")
        if anchor_str:
            try:
                coords = list(map(float, anchor_str.split()))
                if len(coords) == 2: all_x.append(coords[0]); all_y.append(coords[1])
            except ValueError: continue
    if not all_x or not all_y: return None
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
            if tag_name == 'Content' and element.text: text_content_segments.append(element.text)
            elif tag_name == 'Br': text_content_segments.append('\n')
        full_text = "".join(text_content_segments)
        full_text = re.sub(r'\s*\n\s*','\n',full_text); full_text = re.sub(r'\n{2,}','\n',full_text).strip()
        if not full_text and os.path.exists(story_path): print(f"    ℹ️ Story {os.path.basename(story_path)}: No text content extracted.")
        return full_text
    except ET.ParseError as e: print(f"    ❌ ERROR PARSING STORY XML {os.path.basename(story_path)}: {e}"); return ""
    except Exception as e: print(f"    ❌ UNEXPECTED ERROR in get_story_text for {os.path.basename(story_path)}: {e}"); return ""

def find_page_for_item_center(item_id, item_tag, item_center_x, item_center_y, geometric_pages): # Added item_id and item_tag for logging
    print(f"    DEBUG find_page: Item <{item_tag} ID:{item_id}> Checking point ({item_center_x:.2f}, {item_center_y:.2f})") 
    for i, page_info in enumerate(geometric_pages):
        print(f"      DEBUG find_page: vs Page '{page_info.name}' (ID: {page_info.id}) AABB: x:[{page_info.global_aabb['x1']:.2f}-{page_info.global_aabb['x2']:.2f}], y:[{page_info.global_aabb['y1']:.2f}-{page_info.global_aabb['y2']:.2f}]")
        if is_point_in_rect(item_center_x, item_center_y,
                              page_info.global_aabb["y1"], page_info.global_aabb["x1"],
                              page_info.global_aabb["y2"], page_info.global_aabb["x2"]):
            print(f"      DEBUG find_page: MATCH! Item <{item_tag} ID:{item_id}> assigned to Page {page_info.id}") 
            return page_info.id
    print(f"    DEBUG find_page: No page match for Item <{item_tag} ID:{item_id}> at point ({item_center_x:.2f}, {item_center_y:.2f})") 
    return None

def process_spread_element_recursively(element, parent_element, current_accumulated_matrix,
                                       geometric_pages, pages_content_map,
                                       stories_dir, story_cache, depth=0):
    indent = "  " * (depth+1) 
    element_tag_local = element.tag.split('}')[-1]
    element_id = element.get("Self", "UnknownID")
    print(f"{indent}Processing <{element_tag_local} ID:{element_id}> LocalTransform: {element.get('ItemTransform')}") 

    item_local_matrix_str = element.get("ItemTransform")
    item_local_matrix = parse_transform_matrix(item_local_matrix_str)
    item_global_matrix = multiply_matrices(current_accumulated_matrix, item_local_matrix)
    print(f"{indent}  GlobalMatrix for {element_id}: {tuple(f'{x:.2f}' for x in item_global_matrix)}") 

    determined_local_bounds = (0.0,0.0,0.0,0.0)
    if element_tag_local == 'Image':
        gb_prop = element.find('.//{*}Properties/{*}GraphicBounds')
        if gb_prop is not None:
            try:
                gb_l,gb_t,gb_r,gb_b = (float(gb_prop.get(s,"0")) for s in ["Left","Top","Right","Bottom"])
                determined_local_bounds = (gb_t, gb_l, gb_b, gb_r) 
                # print(f"{indent}  Image ID: {element_id} using GraphicBounds: {determined_local_bounds}")
            except Exception as e: print(f"{indent}  ⚠️ Error parsing GraphicBounds for Image {element_id}: {e}")
        if determined_local_bounds == (0.0,0.0,0.0,0.0): 
            item_lbs_img = element.get("GeometricBounds"); 
            if item_lbs_img: determined_local_bounds = parse_geometric_bounds(item_lbs_img)
    elif element_tag_local == "TextFrame":
        path_gb = get_local_bounds_from_path_geometry(element)
        if path_gb: determined_local_bounds = path_gb; print(f"{indent}  TextFrame ID: {element_id} using PathGeometry bounds: {tuple(f'{x:.2f}' for x in determined_local_bounds)}")
        else:
            lbs_tf = element.get("GeometricBounds")
            if lbs_tf: determined_local_bounds = parse_geometric_bounds(lbs_tf)
            # else: print(f"{indent}  TextFrame ID: {element_id} no PathGeometry or GeoBounds. Local bounds 0.")
    else: determined_local_bounds = parse_geometric_bounds(element.get("GeometricBounds"))
    
    item_global_aabb={"x1":0.0,"y1":0.0,"x2":0.0,"y2":0.0} 
    item_center_x,item_center_y = item_global_matrix[4],item_global_matrix[5]
    if determined_local_bounds != (0.0,0.0,0.0,0.0):
        igc = get_global_corners(determined_local_bounds,item_global_matrix)
        item_global_aabb = get_axis_aligned_bounding_box(igc)
        item_center_x,item_center_y = get_item_center(item_global_aabb)
        # print(f"{indent}  Item ID: {element_id} LocalBounds: {determined_local_bounds} -> GlobalAABB & Center calculated")
    
    assigned_page_id = find_page_for_item_center(element_id, element_tag_local, item_center_x, item_center_y, geometric_pages)
    
    if element_tag_local == "TextFrame":
        story_id = element.get("ParentStory")
        # print(f"{indent}  TF {element_id} (Story: {story_id}), Center: ({item_center_x:.1f},{item_center_y:.1f}), AssignedPage: {assigned_page_id}")
        if assigned_page_id and story_id:
            if story_id not in story_cache:
                sfn=f"Story_{story_id}.xml"; sfp=os.path.join(stories_dir,sfn)
                if os.path.exists(sfp): story_cache[story_id]=get_story_text(sfp)
                else: story_cache[story_id]=""; print(f"{indent}    ❌ Story file {sfn} NOT FOUND for TF {element_id}")
            tc=story_cache.get(story_id,""); pd=pages_content_map.get(assigned_page_id)
            # print(f"{indent}    TF {element_id} on Page {assigned_page_id}: PageDataOk: {pd is not None}. TextLen: {len(tc)}. Story: {story_id}")
            if pd:
                if not any(tf['text_frame_id']==element_id for tf in pd["texts"]):
                    if tc:
                        pd["texts"].append({"story_id":story_id,"text_frame_id":element_id,"content":tc,"global_bounds":item_global_aabb,"item_transform":item_local_matrix_str})
                        print(f"{indent}    ✅ Added TextFrame ID: {element_id} (Story: {story_id}) to Page ID: {assigned_page_id}.")
                    elif story_id in story_cache: print(f"{indent}    ℹ️ TF {element_id} (Story: {story_id}) has empty text_content, not adding.")
        elif story_id: print(f"{indent}  ⚠️ TF {element_id} (Story {story_id}) at C:({item_center_x:.1f},{item_center_y:.1f}) NOT assigned to page.")
    elif element_tag_local == "Image":
        if assigned_page_id:
            le=element.find('.//{*}Link')
            if le is not None:
                uri=le.get("LinkResourceURI")
                if uri:
                    pd=pages_content_map.get(assigned_page_id)
                    if pd:
                        ctag=parent_element.tag.split('}')[-1] if parent_element else "Unk"; cid=parent_element.get("Self","Unk") if parent_element else "Unk"
                        if not any(i['uri']==uri and i['image_element_id']==element_id for i in pd["images"]):
                            pd["images"].append({"uri":uri,"image_element_id":element_id,"container_element_tag":ctag,"container_element_id":cid,"global_bounds":item_global_aabb,"item_transform":item_local_matrix_str})
                            # print(f"{indent}  🖼️ Added Image Link: '{uri}' (ID: {element_id}) to Page ID: {assigned_page_id}.")
    if element_tag_local in ["Group","Rectangle","Oval","Polygon"]:
        for child in element: process_spread_element_recursively(child,element,item_global_matrix,geometric_pages,pages_content_map,stories_dir,story_cache,depth+1)

def get_page_content_from_spread(spread_path, stories_dir, story_cache): # MODIFIED
    try:
        tree = ET.parse(spread_path); root = tree.getroot(); spread_element = root
        if not spread_element.tag.endswith('Spread') or not any(c.tag.split('}')[-1]=='Page' for c in spread_element):
            candidate = spread_element.find('{*}Spread') 
            if candidate is not None and any(c.tag.split('}')[-1]=='Page' for c in candidate): spread_element = candidate
            else:
                found_spreads = [el for el in root.iter() if el.tag.endswith('Spread') and any(c.tag.split('}')[-1]=='Page' for c in el)]
                if found_spreads: spread_element = found_spreads[0]
                else: print(f"  ❌ No main <Spread> with <Page> children in {os.path.basename(spread_path)}"); return {}
    except Exception as e: print(f"  ❌ Error parsing {spread_path}: {e}"); return {}, []

    current_spread_pages_content = {}; geometric_pages = []
    spread_base_matrix_str = spread_element.get("ItemTransform") # Get Spread's own transform
    spread_base_matrix = parse_transform_matrix(spread_base_matrix_str)
    print(f"  Processing Spread '{spread_element.get('Self')}'. Spread Base Matrix: {tuple(f'{x:.2f}' for x in spread_base_matrix)}")

    for page_el in spread_element.findall('.//{*}Page'):
        pid=page_el.get("Self"); name=page_el.get("Name",f"UnkPage_{pid}")
        if not pid: continue
        # Pass spread_base_matrix to PageGeometricInfo
        page_info=PageGeometricInfo(pid, name, page_el.get("ItemTransform"), page_el.get("GeometricBounds"), spread_base_matrix)
        geometric_pages.append(page_info)
        if pid not in current_spread_pages_content:
             current_spread_pages_content[pid] = {"name":name,"images":[],"texts":[]}
    if not geometric_pages: print(f"  ℹ️ No Page elements found in {os.path.basename(spread_path)}."); return {}
    
    for child_el in spread_element: 
        ct_local = child_el.tag.split('}')[-1]
        if ct_local in ["Page","FlattenerPreference","Properties"]: continue
        # Pass spread_base_matrix as the initial current_accumulated_matrix for its direct children
        process_spread_element_recursively(child_el,spread_element,spread_base_matrix,geometric_pages,
                                           current_spread_pages_content,stories_dir,story_cache,0)
    return current_spread_pages_content, geometric_pages

def find_spread_files(spreads_dir):
    if not os.path.isdir(spreads_dir): return []
    return [os.path.join(spreads_dir, f) for f in os.listdir(spreads_dir) if f.startswith("Spread_") and f.endswith(".xml")]

def main():
    parser = argparse.ArgumentParser(description="Extract content per page from IDML.")
    parser.add_argument("idml_file", help="Path to .idml file")
    parser.add_argument("--output_json", help="Path to .idml file")
    args = parser.parse_args()
    idml_file = args.idml_file
    output_file_path = args.output_json
    extract(idml_file=idml_file,output_file_path=output_file_path)
    pass

def extract(idml_file,output_file_path):
    temp_dir_base = os.path.join(tempfile.gettempdir(), "idml_extractor_debug")
    os.makedirs(temp_dir_base, exist_ok=True)
    temp_dir = tempfile.mkdtemp(prefix="idml_", dir=temp_dir_base)
    print(f"⏳ Extracting '{idml_file}' to: {temp_dir}")
    if not extract_idml(idml_file, temp_dir):
        if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
        return
    spreads_dir = os.path.join(temp_dir, "Spreads"); stories_dir = os.path.join(temp_dir, "Stories")
    if not os.path.isdir(spreads_dir):
        print(f"❌ Critical: 'Spreads' folder missing at '{spreads_dir}'.")
        if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
        return
    all_page_data_by_id = {}
    story_cache = {} 
    all_page_geometries = {}
    spread_files = find_spread_files(spreads_dir)
    if not spread_files:
        print(f"❌ No spread files found in {spreads_dir}.")
        if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
        return
    for spread_file_path in spread_files:
        spread_name = os.path.basename(spread_file_path)
        print(f"\n📄 Processing Spread File: {spread_name}")
        content_from_this_spread, geoinfo_from_this_spread = get_page_content_from_spread(spread_file_path, stories_dir, story_cache)
        
        for page_info in geoinfo_from_this_spread: # Store geometry info
            if page_info.id not in all_page_geometries:
                all_page_geometries[page_info.id] = page_info
                
        for pid, pdata_in_spread in content_from_this_spread.items():
            if pid not in all_page_data_by_id: all_page_data_by_id[pid] = pdata_in_spread
            else: 
                for img_item in pdata_in_spread["images"]:
                    if not any(ex_img['uri']==img_item['uri'] and ex_img['image_element_id']==img_item['image_element_id'] for ex_img in all_page_data_by_id[pid]["images"]):
                        all_page_data_by_id[pid]["images"].append(img_item)
                current_tf_ids = {t['text_frame_id'] for t in all_page_data_by_id[pid]["texts"]}
                for text_item in pdata_in_spread["texts"]:
                    if text_item['text_frame_id'] not in current_tf_ids:
                        all_page_data_by_id[pid]["texts"].append(text_item); current_tf_ids.add(text_item['text_frame_id'])
    print("\n\n--- ✨ Consolidated Content Per Page ✨ ---")
    if not all_page_data_by_id: print("No content found on any pages.")
    else:
        final_sorted_page_ids = sorted(all_page_data_by_id.keys(), key=lambda pid_key: all_page_data_by_id[pid_key]["name"])
        for pid in final_sorted_page_ids:
            pdata = all_page_data_by_id[pid]; 
            print(f"\n--- Page \"{pdata['name']}\" (ID: {pid}) ---")
            page_geo_info = all_page_geometries.get(pid) # Get geometry info for this page
            if not page_geo_info:
                print(f"⚠️ Warning: Geometry info not found for Page ID {pid}. Skipping coordinate conversion for this page.")
                page_offset_x, page_offset_y = 0, 0
            else:
                page_global_bounds = page_geo_info.global_aabb
                page_offset_x = page_global_bounds["x1"]
                page_offset_y = page_global_bounds["y1"]
                
            if pdata["texts"]:
                print("📝 Texts:")
                for t_item in pdata["texts"]:
                    cp = (t_item['content'][:70] + '...') if len(t_item['content']) > 70 else t_item['content']
                    b = t_item['global_bounds']
                    # print(f"  - Story: {t_item['story_id']} (Frame: {t_item['text_frame_id']}) Bounds: x1={b['x1']:.1f},y1={b['y1']:.1f},x2={b['x2']:.1f},y2={b['y2']:.1f}")
                    print(f"  - Story: {t_item['story_id']}")
                    print(f"    Content: \"{cp.replace(chr(10), r'\\n')}\"")
            else: print("📝 Texts: None")
            if pdata["images"]:
                print("🖼 Images:")
                for img_item in pdata["images"]:
                    b = img_item['global_bounds']
                    print(f"  - URI: {img_item['uri']} (ImageElem: {img_item['image_element_id']}, Container: <{img_item['container_element_tag']} ID:{img_item['container_element_id']}>)")
                    print(f"    Bounds: x1={b['x1']:.1f},y1={b['y1']:.1f},x2={b['x2']:.1f},y2={b['y2']:.1f}")
            else: print("🖼 Images: None")
    print("\n✅ Done. Cleaning up temporary directory...")
    
    # --- Prepare data for JSON output ---
    output_data = {"pages": []}
    if all_page_data_by_id:
        final_sorted_page_ids = sorted(all_page_data_by_id.keys(), 
                                       key=lambda pid_key: all_page_data_by_id[pid_key]["name"])
        for pid in final_sorted_page_ids:
            pdata = all_page_data_by_id[pid]
            
            page_geo_info = all_page_geometries.get(pid)
            if not page_geo_info:
                print(f"⚠️ Warning: Geometry info not found for Page ID {pid}. Skipping coordinate conversion for this page.")
                page_offset_x, page_offset_y = 0, 0
            else:
                page_global_bounds = page_geo_info.global_aabb
                page_offset_x = page_global_bounds["x1"]
                page_offset_y = page_global_bounds["y1"]
                page_width = page_global_bounds["x2"] - page_global_bounds["x1"] 
                page_height = page_global_bounds["y2"] - page_global_bounds["y1"] 
                
            page_output = {
                "page_name": pdata["name"], # Using 'page_name' as key
                "page_id": pid, # Also including the internal ID
                "images": [],
                "texts": []
            }
            for img_item in pdata["images"]:
                item_global_bounds = img_item["global_bounds"]
                page_output["images"].append({
                    "URI": img_item["uri"],
                    "ImageId": img_item['image_element_id'],
                    "Bounds": { # Convert to page-relative coordinates
                        "x1": item_global_bounds["x1"] - page_offset_x, "y1": item_global_bounds["y1"] - page_offset_y,
                        "x2": item_global_bounds["x2"] - page_offset_x, "y2": item_global_bounds["y2"] - page_offset_y,
                    },
                    "page_width":page_width,
                    "page_height":page_height
                })
            for t_item in pdata["texts"]:
                item_global_bounds = t_item["global_bounds"]
                page_output["texts"].append({
                    "Content": t_item["content"],
                    "TextId":t_item['text_frame_id'],
                    "Bounds": { # Convert to page-relative coordinates
                        "x1": item_global_bounds["x1"] - page_offset_x, "y1": item_global_bounds["y1"] - page_offset_y,
                        "x2": item_global_bounds["x2"] - page_offset_x, "y2": item_global_bounds["y2"] - page_offset_y,
                    },
                    "page_width":page_width,
                    "page_height":page_height
                })
            output_data["pages"].append(page_output)

    # --- Write to JSON file ---
    # json_file_path = 'output.json'
    try:
        with open(output_file_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, ensure_ascii=False, indent=4)
        print(f"\n💾 Successfully wrote output to {output_file_path}")
    except IOError as e:
        print(f"\n❌ Error writing JSON to {output_file_path}: {e}")
        
    try: shutil.rmtree(temp_dir); print(f"🗑️ Temporary directory {temp_dir} removed.")
    except Exception as e: print(f"⚠️ Error removing {temp_dir}: {e}. Please remove manually.")

if __name__ == "__main__":
    main()