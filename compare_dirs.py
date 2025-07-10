#!/usr/bin/env python3
import argparse
import json
import os
import subprocess
import tempfile
from collections import defaultdict

STATUS_ORDER = [
    'Added',
    'Removed',
    'Renamed',
    'Modified',
    'RenamedAndModified',
    'Unchanged',
]

ATTR_CONTENT = """*.docx diff=custom
*.doc diff=custom
*.xlsx diff=custom
*.xls diff=custom
*.pptx diff=custom
*.ppt diff=custom
*.pdf diff=custom
"""


def run_git_diff(dir1, dir2, attrs, textconv_path):
    with tempfile.NamedTemporaryFile('w', delete=False) as f:
        f.write(attrs)
        attr_file = f.name
    abs1 = os.path.abspath(dir1)
    abs2 = os.path.abspath(dir2)
    cmd = [
        'git',
        '-c', f'diff.custom.textconv={textconv_path}',
        '-c', f'core.attributesfile={attr_file}',
        'diff', '--no-index', '--textconv', '-M', '--name-status', abs1, abs2
    ]
    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    os.unlink(attr_file)
    return result.stdout


def get_diff_hunk(path1, path2, attrs, textconv_path):
    with tempfile.NamedTemporaryFile('w', delete=False) as f:
        f.write(attrs)
        attr_file = f.name
    abs1 = os.path.abspath(path1)
    abs2 = os.path.abspath(path2)
    cmd = [
        'git',
        '-c', f'diff.custom.textconv={textconv_path}',
        '-c', f'core.attributesfile={attr_file}',
        'diff', '--no-index', '--textconv', '-M', abs1, abs2
    ]
    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    os.unlink(attr_file)
    return result.stdout


def collect_files(root):
    files = []
    for base, _, names in os.walk(root):
        for n in names:
            files.append(os.path.relpath(os.path.join(base, n), root))
    return set(files)


def main():
    parser = argparse.ArgumentParser(description='Compare directories using git diff with textconv.')
    parser.add_argument('dir1')
    parser.add_argument('dir2')
    parser.add_argument('--json', action='store_true', help='Output in JSON format')
    args = parser.parse_args()

    dir1 = os.path.abspath(args.dir1)
    dir2 = os.path.abspath(args.dir2)

    diff_output = run_git_diff(dir1, dir2, ATTR_CONTENT, os.path.abspath('textconv.py'))
    changes = []
    changed_paths = set()

    for line in diff_output.strip().splitlines():
        if not line:
            continue
        parts = line.split('\t')
        status = parts[0]
        if status.startswith('R'):
            score = int(status[1:])
            old = os.path.relpath(parts[1], dir1)
            new = os.path.relpath(parts[2], dir2)
            if score == 100:
                st = 'Renamed'
            else:
                st = 'RenamedAndModified'
            changes.append({'status': st, 'old': old, 'new': new})
            changed_paths.add(old)
            changed_paths.add(new)
        elif status == 'M':
            path = os.path.relpath(parts[1], dir1)
            changes.append({'status': 'Modified', 'path': path})
            changed_paths.add(path)
        elif status == 'A':
            path = os.path.relpath(parts[1], dir2)
            changes.append({'status': 'Added', 'path': path})
            changed_paths.add(path)
        elif status == 'D':
            path = os.path.relpath(parts[1], dir1)
            changes.append({'status': 'Removed', 'path': path})
            changed_paths.add(path)

    all_files_dir1 = collect_files(dir1)
    all_files_dir2 = collect_files(dir2)
    unchanged = sorted((all_files_dir1 & all_files_dir2) - changed_paths)
    for p in unchanged:
        changes.append({'status': 'Unchanged', 'path': p})

    # get diffs for changed files
    for change in changes:
        st = change['status']
        if st in ('Modified', 'RenamedAndModified'):
            old = os.path.join(dir1, change.get('old', change.get('path')))
            new = os.path.join(dir2, change.get('new', change.get('path')))
            change['diff'] = get_diff_hunk(old, new, ATTR_CONTENT, os.path.abspath('textconv.py'))
        elif st == 'Added':
            new = os.path.join(dir2, change['path'])
            change['diff'] = get_diff_hunk(os.devnull, new, ATTR_CONTENT, os.path.abspath('textconv.py'))
        elif st == 'Removed':
            old = os.path.join(dir1, change['path'])
            change['diff'] = get_diff_hunk(old, os.devnull, ATTR_CONTENT, os.path.abspath('textconv.py'))
        elif st == 'Renamed':
            old = os.path.join(dir1, change['old'])
            new = os.path.join(dir2, change['new'])
            change['diff'] = get_diff_hunk(old, new, ATTR_CONTENT, os.path.abspath('textconv.py'))

    # group by status order
    grouped = defaultdict(list)
    for c in changes:
        grouped[c['status']].append(c)

    if args.json:
        out = {st: grouped.get(st, []) for st in STATUS_ORDER}
        print(json.dumps(out, indent=2, ensure_ascii=False))
    else:
        for st in STATUS_ORDER:
            items = grouped.get(st)
            if not items:
                continue
            print(f'## {st}')
            for item in items:
                if st in ('Renamed', 'RenamedAndModified'):
                    print(f"{item['old']} -> {item['new']}")
                else:
                    print(item.get('path'))
                if 'diff' in item:
                    print(item['diff'])

if __name__ == '__main__':
    main()
