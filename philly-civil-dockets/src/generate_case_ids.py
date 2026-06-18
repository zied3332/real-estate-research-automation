import pandas as pd
from pathlib import Path

START_CASE_ID = 141202168
END_CASE_ID = 141202300

OUTPUT_FILE = Path("data/input/case_ids.xlsx")

OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)

case_ids = list(range(START_CASE_ID, END_CASE_ID + 1))

df = pd.DataFrame({
    "case_id": case_ids
})

df.to_excel(OUTPUT_FILE, index=False)

print(f"Generated {len(case_ids)} Case IDs")
print(f"Saved to: {OUTPUT_FILE}")