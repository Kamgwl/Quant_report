import json

transcript_path = r"C:\Users\Administrator\.gemini\antigravity-ide\brain\1d3fb230-091f-4de8-888f-421a8298bf7b\.system_generated\logs\transcript.jsonl"

print("Searching transcript for matching/excel user messages:")
print("-" * 80)
with open(transcript_path, "r", encoding="utf-8") as f:
    for line in f:
        try:
            data = json.loads(line)
            if data.get("source") == "USER_EXPLICIT" or data.get("type") == "USER_INPUT":
                content = data.get("content", "")
                if any(kw in content.lower() for kw in ["match", "excel", "report", "difference"]):
                    print(f"Step {data.get('step_index')}: {content}")
        except Exception as e:
            pass
