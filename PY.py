import pandas as pd
import xml.etree.ElementTree as ET
from pathlib import Path

records = []

root_folder = Path("data")

for xml_file in root_folder.rglob("*.xml"):
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        record = {}

        for item in root.iter():
            if item.attrib.get("Name"):
                record[item.attrib["Name"]] = item.text

        record["SourceFile"] = xml_file.name

        records.append(record)

    except Exception as e:
        print(f"Skipped {xml_file}: {e}")

df = pd.DataFrame(records)

if "ApplicationFolderID" in df.columns:
    df["ApplicationID"] = df["ApplicationFolderID"]

df.to_csv(
    "output.csv",
    index=False,
    encoding="utf-8-sig"
)

print(f"Processed {len(df)} records")