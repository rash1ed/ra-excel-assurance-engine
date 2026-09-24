from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]
TEXT_EXTS = {'.py', '.md', '.txt', '.toml', '.json'}
PATTERNS = {
    'windows_user_path': re.compile(r'C:\\\\Users\\\\', re.I),
    'personal_email': re.compile(r'rashidalbloushi5@gmail\\.com', re.I),
    'openai_like_token': re.compile(r'sk-[A-Za-z0-9_-]{10,}'),
    'github_pat': re.compile(r'ghp_[A-Za-z0-9]{20,}'),
}

findings = []
for path in ROOT.rglob('*'):
    if not path.is_file():
        continue
    if path.name == '.gitignore' or path.suffix.lower() in TEXT_EXTS:
        text = path.read_text(encoding='utf-8', errors='ignore')
        for label, pattern in PATTERNS.items():
            if pattern.search(text):
                findings.append((str(path.relative_to(ROOT)), label))

print('SCAN_FINDINGS', findings)
raise SystemExit(1 if findings else 0)
