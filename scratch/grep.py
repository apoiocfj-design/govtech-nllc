import sys
path = r"C:\Users\CFJ-SECOM\.gemini\antigravity\scratch\govtech-nllc\frontend\index.html"
terms = sys.argv[1].split("|")
for i, line in enumerate(open(path, encoding="utf-8"), 1):
    if any(t in line for t in terms):
        print(i, line.rstrip()[:200])
