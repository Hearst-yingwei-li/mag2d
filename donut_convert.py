import json
from pathlib import Path

# === Input & Output Paths ===
input_json = Path("label_studio_test.json")
output_json = Path("donut_finetune_ready.json")

# === Load Label Studio JSON ===
with open(input_json, "r", encoding="utf-8") as f:
    label_studio_data = json.load(f)

# === Convert to Donut Format ===
donut_dataset = []

for task in label_studio_data:
    image_path = task["data"]["image"].split("/")[-1]
    page_annotations = []

    # Maps for annotation info
    bbox_map = {}     # id → bounding box with label
    text_map = {}     # id → textarea values (if not already attached)
    relations = []    # list of relations

    for item in task["annotations"][0]["result"]:
        item_type = item.get("type")

        # --- Rectangle labels (bounding boxes) ---
        if item_type == "rectanglelabels" and "id" in item:
            item_id = item["id"]
            label = item["value"]["rectanglelabels"][0]
            bbox_map[item_id] = {
                "label": label,
                "bbox": {
                    "x": item["value"]["x"],
                    "y": item["value"]["y"],
                    "width": item["value"]["width"],
                    "height": item["value"]["height"],
                },
                "text": ""  # Placeholder, might be updated later
            }

        # --- Text attached to region ---
        elif item_type == "textarea" and "id" in item:
            item_id = item["id"]
            text_content = item["value"]["text"][0] if isinstance(item["value"]["text"], list) else item["value"]["text"]
            if item_id in bbox_map:
                bbox_map[item_id]["text"] = text_content
            else:
                text_map[item_id] = text_content

        # --- Relationships ---
        elif item_type == "relation":
            relations.append({
                "from": item["from_id"],
                "to": item["to_id"],
                "type": item["labels"][0] if "labels" in item and item["labels"] else ""
            })

    # --- Assemble Annotation List ---
    for obj_id, data in bbox_map.items():
        page_annotations.append({
            "label": data["label"],
            "text": data.get("text", ""),
            "bbox": data["bbox"],
            "id": obj_id
        })

    # --- Build Final JSON for this page ---
    donut_item = {
        "file_name": image_path,
        "annotations": page_annotations,
        "relations": relations
    }

    donut_dataset.append(donut_item)

# === Save to Donut Format JSON ===
with open(output_json, "w", encoding="utf-8") as f:
    json.dump(donut_dataset, f, ensure_ascii=False, indent=2)

print(f"✅ Converted file saved to: {output_json}")
