"""Build web/index.html by inlining the notebook's JSON export into the page template.

Run from the tideline folder after the notebook has run:  python web/build_web.py
"""
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)


def main():
    with open(os.path.join(ROOT, "outputs", "web_data.json"), encoding="utf-8") as f:
        data = json.load(f)
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    with open(os.path.join(HERE, "template.html"), encoding="utf-8") as f:
        page = f.read()
    out = os.path.join(HERE, "index.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(page.replace("__TIDELINE_DATA__", payload))
    print(f"wrote {out} ({os.path.getsize(out) / 1024:.0f} KB)")


if __name__ == "__main__":
    main()
