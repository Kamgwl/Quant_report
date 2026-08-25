import json

transcript_path = r"C:\Users\Administrator\.gemini\antigravity-ide\brain\1d3fb230-091f-4de8-888f-421a8298bf7b\.system_generated\logs\transcript.jsonl"

with open(transcript_path, "r", encoding="utf-8") as f:
    for line in f:
        try:
            data = json.loads(line)
            if data.get("step_index") == 816:
                print(json.dumps(data, indent=2))
                break
        except Exception as e:
            pass
