import zipfile
import xml.etree.ElementTree as ET
import os
import argparse
import shutil
import tempfile
import re
import math

# --- Helper Functions for Geometry and Transforms (mostly from previous version) ---


def parse_transform_matrix(transform_str):
    if not transform_str:
        return (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
    try:
        parts = list(map(float, transform_str.split()))
        if len(parts) == 6:
            return tuple(parts)
    except ValueError:
        pass
    print(
        f"⚠️ Warning: Could not parse transform string: '{transform_str}'. Using identity."
    )
    return (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


def parse_geometric_bounds(bounds_str):
    if not bounds_str:
        return (0.0, 0.0, 0.0, 0.0)
    try:
        parts = list(map(float, bounds_str.split()))
        if len(parts) == 4:
            return tuple(parts)  # y1, x1, y2, x2
    except ValueError:
        pass
    print(
        f"⚠️ Warning: Could not parse geometric bounds: '{bounds_str}'. Using zero bounds."
    )
    return (0.0, 0.0, 0.0, 0.0)


def multiply_matrices(m1, m2):
    a1, b1, c1, d1, tx1, ty1 = m1
    a2, b2, c2, d2, tx2, ty2 = m2
    a_res = a1 * a2 + c1 * b2
    b_res = b1 * a2 + d1 * b2
    c_res = a1 * c2 + c1 * d2
    d_res = b1 * c2 + d1 * d2
    tx_res = a1 * tx2 + c1 * ty2 + tx1
    ty_res = b1 * tx2 + d1 * ty2 + ty1
    return (a_res, b_res, c_res, d_res, tx_res, ty_res)


def apply_transform_to_point(x, y, matrix):
    a, b, c, d, tx, ty = matrix
    new_x = a * x + c * y + tx
    new_y = b * x + d * y + ty
    return new_x, new_y


def get_global_corners(local_bounds, global_item_matrix):
    """Calculates the four corner points of an item in global coordinates."""
    y1, x1, y2, x2 = local_bounds
    corners = [
        apply_transform_to_point(x1, y1, global_item_matrix),  # Top-left
        apply_transform_to_point(x2, y1, global_item_matrix),  # Top-right
        apply_transform_to_point(x1, y2, global_item_matrix),  # Bottom-left
        apply_transform_to_point(x2, y2, global_item_matrix),  # Bottom-right
    ]
    return corners  # List of (x,y) tuples


def get_axis_aligned_bounding_box(corners):
    """Calculates the axis-aligned bounding box from a list of corner points."""
    if not corners:
        return {"x1": 0, "y1": 0, "x2": 0, "y2": 0}
    all_x = [p[0] for p in corners]
    all_y = [p[1] for p in corners]
    return {"x1": min(all_x), "y1": min(all_y), "x2": max(all_x), "y2": max(all_y)}


def get_item_center(global_aabb):
    """Calculates the center of an axis-aligned bounding box."""
    return (
        (global_aabb["x1"] + global_aabb["x2"]) / 2.0,
        (global_aabb["y1"] + global_aabb["y2"]) / 2.0,
    )


def is_point_in_rect(px, py, rect_y1, rect_x1, rect_y2, rect_x2):
    return (rect_x1 <= px <= rect_x2) and (rect_y1 <= py <= rect_y2)


class PageGeometricInfo:
    def __init__(self, self_id, name, page_matrix_str, page_bounds_str):
        self.id = self_id
        self.name = name
        self.matrix = parse_transform_matrix(
            page_matrix_str
        )  # This is the page's ItemTransform
        self.local_bounds = parse_geometric_bounds(page_bounds_str)

        # Page's global bounds are its local bounds transformed by its own ItemTransform
        page_global_corners = get_global_corners(self.local_bounds, self.matrix)
        self.global_aabb = get_axis_aligned_bounding_box(
            page_global_corners
        )  # Stored as {"x1":val, ...}

        print(
            f"<<  Page '{self.name}' (ID: {self.id}): \npage_bounds_str = {page_bounds_str}"
            f"Global AABB=(x1: {self.global_aabb['x1']:.2f} x2: {self.global_aabb['x2']:.2f}, "
            f"y1: {self.global_aabb['y1']:.2f} y2: {self.global_aabb['y2']:.2f})  >>\n"
        )


# --- Main Extraction Logic --- (extract_idml, get_story_text are same)
def extract_idml(idml_path, extract_to):
    try:
        with zipfile.ZipFile(idml_path, "r") as zip_ref:
            zip_ref.extractall(extract_to)
        print(f"✅ Successfully extracted IDML to {extract_to}")
        return True
    except Exception as e:
        print(f"❌ Extraction failed: {e}")
        return False


def get_story_text(story_path):
    text_content_segments = []
    try:
        tree = ET.parse(story_path)
        root = tree.getroot()
        for element in root.iter():
            tag_name = element.tag.split("}")[-1]
            if tag_name == "Content":
                if element.text:
                    text_content_segments.append(element.text)
            elif tag_name == "Br":
                text_content_segments.append("\n")
        full_text = "".join(text_content_segments)
        full_text = re.sub(r"\s*\n\s*", "\n", full_text)
        full_text = re.sub(r"\n{2,}", "\n", full_text).strip()
        # Keep logging minimal here, focus on spread processing logs
        # if not full_text:
        #     print(f"    ℹ️ No text content found in story: {os.path.basename(story_path)}")
        # else:
        #     print(f"    📝 Extracted text from story: {os.path.basename(story_path)} (approx {len(full_text)} chars)")
        return full_text
    except ET.ParseError as e:
        print(f"    ❌ Error parsing story XML {story_path}: {e}")
    except Exception as e:
        print(f"    ❌ Unexpected error in get_story_text for {story_path}: {e}")
    return ""


def find_page_for_item_center(item_center_x, item_center_y, geometric_pages):
    """Finds which page an item's center point falls into."""
    for page_info in geometric_pages:
        if is_point_in_rect(
            item_center_x,
            item_center_y,
            page_info.global_aabb["y1"],
            page_info.global_aabb["x1"],
            page_info.global_aabb["y2"],
            page_info.global_aabb["x2"],
        ):
            return page_info.id
    return None


# Modified function signature
def process_spread_element_recursively(
    element,
    parent_element,
    current_accumulated_matrix,
    geometric_pages,
    pages_content_map,
    stories_dir,
    story_cache,
    depth=0,
):
    indent = "  " * (depth + 2)
    element_tag_local = element.tag.split("}")[-1]
    element_id = element.get("Self", "UnknownID")

    item_local_matrix_str = element.get("ItemTransform")
    item_local_matrix = parse_transform_matrix(item_local_matrix_str)
    item_global_matrix = multiply_matrices(
        current_accumulated_matrix, item_local_matrix
    )
    # For <Image> there is no 'GeometricBounds' property, however can be found in <Properties> -> <GraphicBounds Left="0" Top="0" Right="961.92" Bottom="1442.88"/>
    # print(f"+++ element tag local = {element_tag_local}")
    if element_tag_local == "Image":
        graphic_bounds_prop = element.find(
            ".//{*}Properties/{*}GraphicBounds"
        )  # Relative XPath to find GraphicBounds
        if graphic_bounds_prop is not None:
            try:
                gb_left = float(graphic_bounds_prop.get("Left", "0"))
                gb_top = float(graphic_bounds_prop.get("Top", "0"))
                gb_right = float(graphic_bounds_prop.get("Right", "0"))
                gb_bottom = float(graphic_bounds_prop.get("Bottom", "0"))
                # Override actual_local_bounds with GraphicBounds values
                item_local_bounds = (gb_top, gb_left, gb_bottom, gb_right)
            except (ValueError, TypeError) as e:
                print(
                    f"{indent}⚠️ Error parsing GraphicBounds for Image ID: {element_id} - {e}. Attributes: {graphic_bounds_prop.attrib}"
                )
        # else: # Optional: for debugging if GraphicBounds tag itself is not found
        # print(f"{indent}ℹ️ Image ID: {element_id} did not find GraphicBounds element under Properties.")
    else:
        item_local_bounds_str = element.get("GeometricBounds")
        item_local_bounds = parse_geometric_bounds(
                item_local_bounds_str
            )  # y1, x1, y2, x2

    # item_global_aabb = {"x1": 0.0, "y1": 0.0, "x2": 0.0, "y2": 0.0}
    # item_center_x, item_center_y = item_global_matrix[4], item_global_matrix[5] # Default to transformed origin

    ##

    item_global_corners = []
    item_global_aabb = {
        "x1": 0.0,
        "y1": 0.0,
        "x2": 0.0,
        "y2": 0.0,
    }  # Ensure float for consistency
    # item_center_x, item_center_y = 0.0, 0.0 # Initialize as float
    item_center_x, item_center_y = (
        item_global_matrix[4],
        item_global_matrix[5],
    )  # Default to transformed origin
    # print(f'-------- item center x == {item_center_x}  item center y == {item_center_y}')
    # Calculate global corners and AABB
    if item_local_bounds != (0.0, 0.0, 0.0, 0.0) or element_tag_local in ["Group"]:
        item_global_corners = get_global_corners(item_local_bounds, item_global_matrix)
        item_global_aabb = get_axis_aligned_bounding_box(item_global_corners)
        # *** This is where item_center_x and item_center_y should be defined ***
        item_center_x, item_center_y = get_item_center(item_global_aabb)

    # Now item_center_x and item_center_y are defined before being used
    assigned_page_id = find_page_for_item_center(
        item_center_x, item_center_y, geometric_pages
    )

    # --- Handle Image (often inside Rectangle, Oval, Polygon) ---
    if element_tag_local == "Image":
        if assigned_page_id:
            link_element = element.find(".//{*}Link")
            if link_element is not None:
                uri = link_element.get("LinkResourceURI")
                if uri:
                    page_data = pages_content_map.get(assigned_page_id)
                    if page_data:
                        container_tag = "Unknown"
                        container_id = "Unknown"
                        if (
                            parent_element is not None
                        ):  # Check if parent_element was passed
                            container_tag = parent_element.tag.split("}")[-1]
                            container_id = parent_element.get("Self", "Unknown")

                        if not any(
                            img["uri"] == uri and img["image_element_id"] == element_id
                            for img in page_data["images"]
                        ):
                            page_data["images"].append(
                                {
                                    "uri": uri,
                                    "image_element_id": element_id,
                                    "container_element_tag": container_tag,  # Use parent_element info
                                    "container_element_id": container_id,  # Use parent_element info
                                    "global_bounds": item_global_aabb,
                                    "item_transform": item_local_matrix_str,
                                }
                            )
                            print(
                                f"{indent}🖼️ Added Image Link: '{uri}' (ID: {element_id}) to Page ID: {assigned_page_id}. Container: <{container_tag} ID:{container_id}> Bounds: x1={item_global_aabb['x1']:.1f}, y1={item_global_aabb['y1']:.1f}"
                            )

    elif element_tag_local == "TextFrame":
        story_id = element.get("ParentStory")
        # --- Start Crucial Debug Prints for TextFrames ---
        print(
            f"{indent}Processing TextFrame ID: {element_id}, StoryID: {story_id}, CalculatedCenter: ({item_center_x:.1f},{item_center_y:.1f}), AssignedPageID: {assigned_page_id}"
        )
        # --- End Crucial Debug Prints ---
        if assigned_page_id and story_id:
            if story_id not in story_cache:
                story_filename = f"Story_{story_id}.xml"
                story_file_path = os.path.join(stories_dir, story_filename)
                # print(f"{indent}  Attempting to load story: {story_filename}") # Keep this less verbose unless needed
                if os.path.exists(story_file_path):
                    story_cache[story_id] = get_story_text(
                        story_file_path
                    )  # Ensure get_story_text logs its own errors
                else:
                    print(
                        f"{indent}  ❌ Story file {story_filename} NOT FOUND for TextFrame ID: {element_id}"
                    )
                    story_cache[story_id] = ""

            text_content = story_cache.get(story_id, "")
            page_data = pages_content_map.get(assigned_page_id)

            # --- Start Crucial Debug Prints for TextFrames ---
            print(
                f"{indent}  For TextFrame {element_id} on Page {assigned_page_id}: Page data OK: {page_data is not None}. Text content len: {len(text_content)}. Story: {story_id}"
            )
            # --- End Crucial Debug Prints ---

            if page_data:
                is_tf_added = any(
                    tf["text_frame_id"] == element_id for tf in page_data["texts"]
                )
                if not is_tf_added:
                    if text_content:  # Only add if there's actual content
                        page_data["texts"].append(
                            {
                                "story_id": story_id,
                                "text_frame_id": element_id,
                                "content": text_content,
                                "global_bounds": item_global_aabb,
                                "item_transform": item_local_matrix_str,
                            }
                        )
                        print(
                            f"{indent}  ✅ Added TextFrame ID: {element_id} (Story: {story_id}) to Page ID: {assigned_page_id}."
                        )
                    elif (
                        story_id in story_cache
                    ):  # text_content is empty but story_id was processed
                        print(
                            f"{indent}  ℹ️ TextFrame ID: {element_id} (Story: {story_id}) has empty text_content from cache (story empty or parse error), not adding."
                        )
        elif story_id:
            print(
                f"{indent}⚠️ TextFrame ID: {element_id} (Story: {story_id}) at center ({item_center_x:.1f},{item_center_y:.1f}) was NOT assigned to any page."
            )
        pass
    # --- Recursively process children ---
    if element_tag_local in ["Group", "Rectangle", "Oval", "Polygon"]:
        for child_element in element:
            process_spread_element_recursively(
                child_element,
                element,  # Pass current element as the parent for the child
                item_global_matrix,
                geometric_pages,
                pages_content_map,
                stories_dir,
                story_cache,
                depth + 1,
            )


def get_page_content_from_spread(spread_path, stories_dir, story_cache):
    try:
        tree = ET.parse(spread_path)
        root = tree.getroot()

        spread_element = root
        if root.tag.endswith("Spread") and not any(
            child.tag.split("}")[-1] == "Page" for child in root
        ):
            actual_spread_candidate = root.find("{*}Spread")
            if actual_spread_candidate is not None:
                spread_element = actual_spread_candidate
            else:
                if not any(child.tag.split("}")[-1] == "Page" for child in root):
                    print(
                        f"  ❌ Error: Could not find the main <Spread> element with <Page> children in {os.path.basename(spread_path)}"
                    )
                    return {}
        print(
            f"  Processing Spread Element: <{spread_element.tag.split('}')[-1]} Self='{spread_element.get('Self')}'>"
        )
    except Exception as e:
        print(f"  ❌ Error parsing or finding spread element in {spread_path}: {e}")
        return {}

    current_spread_pages_content = {}
    geometric_pages = []

    for page_el in spread_element.findall(".//{*}Page"):
        pid = page_el.get("Self")
        name = page_el.get("Name", f"UnnamedPage_{pid}")
        if not pid:
            print(f"  ⚠️ Found a Page element without a 'Self' ID. Skipping.")
            continue

        page_info = PageGeometricInfo(
            pid, name, page_el.get("ItemTransform"), page_el.get("GeometricBounds")
        )
        geometric_pages.append(page_info)
        if pid not in current_spread_pages_content:
            current_spread_pages_content[pid] = {
                "name": name,
                "images": [],
                "texts": [],
            }
    if not geometric_pages:
        print(
            f"  ℹ️ No Page elements with geometric info found. Cannot associate content."
        )
        return {}

    spread_base_matrix_str = spread_element.get("ItemTransform")
    spread_base_matrix = parse_transform_matrix(spread_base_matrix_str)
    # print(f'>> spread element length == {len(spread_element)}')
    for child_el in spread_element:
        child_tag_local = child_el.tag.split("}")[-1]
        # print(f'>> -- child_el --- {child_tag_local}')
        if child_tag_local in [
            "Page",
            "FlattenerPreference",
            "Properties",
        ]:  # Skip already processed or non-content metadata
            continue
        process_spread_element_recursively(
            child_el,
            spread_element,
            spread_base_matrix,
            geometric_pages,
            current_spread_pages_content,
            stories_dir,
            story_cache,
        )
    # print(f'<<<<<<<<<<<< spread pages content:\n{current_spread_pages_content}')
    return current_spread_pages_content


# --- Main function and argument parsing (similar to before) ---
def find_spread_files(spreads_dir):
    if not os.path.isdir(spreads_dir):
        print(f"❌ Spreads directory not found at: {spreads_dir}")
        return []
    files = [
        os.path.join(spreads_dir, f)
        for f in os.listdir(spreads_dir)
        if f.startswith("Spread_") and f.endswith(".xml")
    ]
    if not files:
        print(f"ℹ️ No spread XML files (Spread_*.xml) found in {spreads_dir}")
    # else: print(f"Found spread files in {spreads_dir}: {files}") # Verbose
    return files


def find_story_files(stories_dir):
    if not os.path.isdir(stories_dir):
        # print(f"❌ Stories directory not found at: {stories_dir}") # Can be optional if no text
        return []
    files = [
        f
        for f in os.listdir(stories_dir)
        if f.startswith("Story_") and f.endswith(".xml")
    ]
    # if not files: print(f"ℹ️ No story XML files (Story_*.xml) found in {stories_dir}") # Verbose
    # else: print(f"Found story files in {stories_dir}: {files[:5]}..." if len(files) > 5 else files) # Verbose
    return files


def main():
    parser = argparse.ArgumentParser(
        description="Extract content per page from IDML, including text frame bounding boxes."
    )
    parser.add_argument("idml_file", help="Path to .idml file")
    args = parser.parse_args()

    temp_dir_base = os.path.join(tempfile.gettempdir(), "idml_extractor_debug")
    os.makedirs(temp_dir_base, exist_ok=True)
    temp_dir = tempfile.mkdtemp(prefix="idml_", dir=temp_dir_base)

    print(f"⏳ Extracting '{args.idml_file}' to: {temp_dir}")
    if not extract_idml(args.idml_file, temp_dir):
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        return

    # print(f"📂 Contents of temp_dir '{temp_dir}': {os.listdir(temp_dir)}") # Verbose

    spreads_dir = os.path.join(temp_dir, "Spreads")
    stories_dir = os.path.join(temp_dir, "Stories")

    if not os.path.isdir(spreads_dir):
        print(f"❌ Critical: 'Spreads' folder missing at '{spreads_dir}'.")
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        return
    # find_story_files(stories_dir) # Optional: for debugging available stories

    all_page_data_by_id = {}
    story_cache = {}

    spread_files = find_spread_files(spreads_dir)
    if not spread_files:
        print("❌ No spread files found.")
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        return

    for spread_file_path in spread_files:
        spread_name = os.path.basename(spread_file_path)
        print(f"\n📄 Processing Spread File: {spread_name}")

        content_from_this_spread = get_page_content_from_spread(
            spread_file_path, stories_dir, story_cache
        )

        for pid, pdata_in_spread in content_from_this_spread.items():
            if pid not in all_page_data_by_id:
                all_page_data_by_id[pid] = {
                    "name": pdata_in_spread["name"],
                    "images": list(pdata_in_spread["images"]),
                    "texts": list(pdata_in_spread["texts"]),
                }
            else:
                print(
                    f"  🔄 Aggregating content for Page ID {pid} (Name: {pdata_in_spread['name']})"
                )
                for img_item in pdata_in_spread["images"]:  # Images are now dicts
                    # Crude check for duplication based on URI and image element ID
                    if not any(
                        existing_img["uri"] == img_item["uri"]
                        and existing_img["image_element_id"]
                        == img_item["image_element_id"]
                        for existing_img in all_page_data_by_id[pid]["images"]
                    ):
                        all_page_data_by_id[pid]["images"].append(img_item)

                current_tf_ids = {
                    t["text_frame_id"] for t in all_page_data_by_id[pid]["texts"]
                }
                for text_item in pdata_in_spread["texts"]:  # Texts are now dicts
                    if text_item["text_frame_id"] not in current_tf_ids:
                        all_page_data_by_id[pid]["texts"].append(text_item)
                        current_tf_ids.add(text_item["text_frame_id"])

    print("\n\n--- ✨ Consolidated Content Per Page ✨ ---")
    if not all_page_data_by_id:
        print("No content found on any pages after processing all spreads.")
    else:
        final_sorted_page_ids = sorted(
            all_page_data_by_id.keys(),
            key=lambda pid_key: all_page_data_by_id[pid_key]["name"],
        )

        for pid in final_sorted_page_ids:
            pdata = all_page_data_by_id[pid]
            print(f"\n--- Page \"{pdata['name']}\" (ID: {pid}) ---")
            if pdata["texts"]:
                print("📝 Texts:")
                for t_idx, t_item in enumerate(pdata["texts"]):
                    content_prev = (
                        (t_item["content"][:70] + "...")
                        if len(t_item["content"]) > 70
                        else t_item["content"]
                    )
                    content_prev = content_prev.replace("\n", "\\n")
                    bounds = t_item["global_bounds"]
                    print(
                        f"  [{t_idx+1}] Story ID: {t_item['story_id']} (Frame: {t_item['text_frame_id']})"
                    )
                    print(f'      Content: "{content_prev}"')
                    print(
                        f"      GlobalBounds: x1={bounds['x1']:.2f}, y1={bounds['y1']:.2f}, x2={bounds['x2']:.2f}, y2={bounds['y2']:.2f}"
                    )
                    # print(f"      ItemTransform: {t_item['item_transform']}") # Optional: for debugging
            else:
                print("📝 Texts: None")

            if pdata["images"]:
                print("🖼 Images:")
                for img_item in pdata["images"]:
                    bounds = img_item["global_bounds"]
                    print(f"  - URI: {img_item['uri']}")
                    print(
                        f"    ImageElemID: {img_item['image_element_id']}, Container: <{img_item['container_element_tag']} ID:{img_item['container_element_id']}>"
                    )
                    print(
                        f"    GlobalBounds: x1={bounds['x1']:.2f}, y1={bounds['y1']:.2f}, x2={bounds['x2']:.2f}, y2={bounds['y2']:.2f}"
                    )
                    # print(f"    ItemTransform: {img_item['item_transform']}") # Optional: for debugging
            else:
                print("🖼 Images: None")

    print("\n✅ Done. Cleaning up temporary directory...")
    try:
        shutil.rmtree(temp_dir)
        print(f"🗑️ Temporary directory {temp_dir} removed.")
    except Exception as e:  # Use 'as e'
        print(
            f"⚠️ Error removing temporary directory {temp_dir}: {e}. Please remove manually."
        )


if __name__ == "__main__":
    main()
