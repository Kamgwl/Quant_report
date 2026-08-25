import json

transcript_path = r"C:\Users\Administrator\.gemini\antigravity-ide\brain\1d3fb230-091f-4de8-888f-421a8298bf7b\.system_generated\logs\transcript.jsonl"

print("Searching transcript for STOCK FO:")
print("-" * 80)
with open(transcript_path, "r", encoding="utf-8") as f:
    for line in f:
        try:
            data = json.loads(line)
            content = str(data)
            if "STOCK FO" in content:
                print(f"Step {data.get('step_index')}: {data.get('type')} from {data.get('source')}")
        except Exception as e:
            pass
