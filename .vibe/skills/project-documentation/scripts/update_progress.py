"""
Script to update PROGRESS.md file with Epic/Story/Task structure

This file is part of xlManage.

xlManage is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

xlManage is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with xlManage.  If not, see <https://www.gnu.org/licenses/>.
"""

import argparse
import sys
from pathlib import Path
from datetime import datetime
from typing import Optional


PROGRESS_FILE = Path("PROGRESS.md")


def _find_section_index(lines: list[str], section_name: str) -> Optional[int]:
    """Find the index of a section header (## Epic, ### Story, etc.)"""
    for i, line in enumerate(lines):
        if section_name in line:
            return i
    return None


def _find_section_end(lines: list[str], start_index: int) -> int:
    """Find the end of a section (next section of same or higher level)"""
    level = len([c for c in lines[start_index] if c == "#"])
    for i in range(start_index + 1, len(lines)):
        if lines[i].strip().startswith("#"):
            line_level = len([c for c in lines[i] if c == "#"])
            if line_level <= level:
                return i
    return len(lines)


def add_epic(epic_name: str, objective: str = "") -> bool:
    """Add a new Epic to PROGRESS.md"""
    if not PROGRESS_FILE.exists():
        print(f"Error: {PROGRESS_FILE} not found")
        return False

    content = PROGRESS_FILE.read_text(encoding="utf-8")
    lines = content.split("\n")

    # Check if epic already exists
    if _find_section_index(lines, f"## {epic_name}") is not None:
        print(f"Error: Epic '{epic_name}' already exists")
        return False

    # Find where to insert (before "## En cours" or at end)
    insert_index = len(lines)
    en_cours_idx = _find_section_index(lines, "## En cours")
    if en_cours_idx is not None:
        insert_index = en_cours_idx

    epic_header = f"## {epic_name}"
    objective_line = f"### Objectif\n{objective}" if objective else "### Objectif\n"
    date = datetime.now().strftime("%Y-%m-%d")

    lines.insert(insert_index, "")
    lines.insert(insert_index, f"### Début estimé : {date}")
    lines.insert(insert_index, objective_line)
    lines.insert(insert_index, epic_header)

    PROGRESS_FILE.write_text("\n".join(lines), encoding="utf-8")
    print(f"✓ Epic '{epic_name}' added")
    return True


def add_story(epic_name: str, story_name: str, description: str = "") -> bool:
    """Add a new Story to an Epic"""
    if not PROGRESS_FILE.exists():
        print(f"Error: {PROGRESS_FILE} not found")
        return False

    content = PROGRESS_FILE.read_text(encoding="utf-8")
    lines = content.split("\n")

    epic_idx = _find_section_index(lines, f"## {epic_name}")
    if epic_idx is None:
        print(f"Error: Epic '{epic_name}' not found")
        return False

    epic_end = _find_section_end(lines, epic_idx)
    story_idx = _find_section_index(lines[epic_idx:epic_end], f"### {story_name}")
    if story_idx is not None:
        print(f"Error: Story '{story_name}' already exists in epic '{epic_name}'")
        return False

    # Insert story after Objective section
    objective_idx = _find_section_index(lines[epic_idx:epic_end], "### Objectif")
    insert_index = epic_idx + (objective_idx + 2 if objective_idx is not None else 1)

    story_header = f"### {story_name}"
    desc_line = f"- {description}" if description else ""

    lines.insert(insert_index, "")
    lines.insert(insert_index, desc_line)
    lines.insert(insert_index, story_header)

    PROGRESS_FILE.write_text("\n".join(lines), encoding="utf-8")
    print(f"✓ Story '{story_name}' added to epic '{epic_name}'")
    return True


def add_task_to_story(epic_name: str, story_name: str, task: str) -> bool:
    """Add a task to a Story"""
    if not PROGRESS_FILE.exists():
        print(f"Error: {PROGRESS_FILE} not found")
        return False

    content = PROGRESS_FILE.read_text(encoding="utf-8")
    lines = content.split("\n")

    epic_idx = _find_section_index(lines, f"## {epic_name}")
    if epic_idx is None:
        print(f"Error: Epic '{epic_name}' not found")
        return False

    epic_end = _find_section_end(lines, epic_idx)
    story_idx = _find_section_index(lines[epic_idx:epic_end], f"### {story_name}")
    if story_idx is None:
        print(f"Error: Story '{story_name}' not found in epic '{epic_name}'")
        return False

    story_abs_idx = epic_idx + story_idx
    story_end = _find_section_end(lines[epic_idx:epic_end], story_idx) + epic_idx

    # Find last task or insert after description
    insert_index = story_end
    for i in range(story_abs_idx + 1, story_end):
        if lines[i].strip().startswith("- ["):
            insert_index = i + 1

    task_line = f"- [ ] {task}"
    lines.insert(insert_index, task_line)

    PROGRESS_FILE.write_text("\n".join(lines), encoding="utf-8")
    print(f"✓ Task added to '{epic_name} > {story_name}'")
    return True


def complete_task(epic_name: str, story_name: str, task_pattern: str) -> bool:
    """Mark a task as completed"""
    if not PROGRESS_FILE.exists():
        print(f"Error: {PROGRESS_FILE} not found")
        return False

    content = PROGRESS_FILE.read_text(encoding="utf-8")
    lines = content.split("\n")

    epic_idx = _find_section_index(lines, f"## {epic_name}")
    if epic_idx is None:
        print(f"Error: Epic '{epic_name}' not found")
        return False

    epic_end = _find_section_end(lines, epic_idx)
    story_idx = _find_section_index(lines[epic_idx:epic_end], f"### {story_name}")
    if story_idx is None:
        print(f"Error: Story '{story_name}' not found in epic '{epic_name}'")
        return False

    story_abs_idx = epic_idx + story_idx
    story_end = _find_section_end(lines[epic_idx:epic_end], story_idx) + epic_idx

    # Find and complete task
    found = False
    for i in range(story_abs_idx + 1, story_end):
        if task_pattern.lower() in lines[i].lower() and "- [ ]" in lines[i]:
            lines[i] = lines[i].replace("- [ ]", "- [x]")
            found = True
            break

    if not found:
        print(f"Error: Task matching '{task_pattern}' not found")
        return False

    PROGRESS_FILE.write_text("\n".join(lines), encoding="utf-8")
    print(f"✓ Task completed in '{epic_name} > {story_name}'")
    return True


def add_completed(description: str) -> bool:
    """Add an entry to the Terminées section"""
    if not PROGRESS_FILE.exists():
        print(f"Error: {PROGRESS_FILE} not found")
        return False

    content = PROGRESS_FILE.read_text(encoding="utf-8")
    lines = content.split("\n")

    terminées_idx = _find_section_index(lines, "## Terminées")
    if terminées_idx is None:
        print("Error: Section 'Terminées' not found")
        return False

    insert_index = terminées_idx + 1
    date = datetime.now().strftime("%Y-%m-%d")
    entry = f"- {date} : {description}"

    lines.insert(insert_index, entry)

    PROGRESS_FILE.write_text("\n".join(lines), encoding="utf-8")
    print(f"✓ Added to Terminées: {description}")
    return True


def list_epics() -> bool:
    """List all Epics in PROGRESS.md"""
    if not PROGRESS_FILE.exists():
        print(f"Error: {PROGRESS_FILE} not found")
        return False

    content = PROGRESS_FILE.read_text(encoding="utf-8")
    lines = content.split("\n")

    print("\nEpics:")
    for line in lines:
        if line.startswith("## Epic"):
            print(f"  {line[3:]}")

    return True


def main():
    parser = argparse.ArgumentParser(
        description="Update PROGRESS.md with Epic/Story/Task structure"
    )
    subparsers = parser.add_subparsers(dest="command", help="Command to execute")

    # List epics
    subparsers.add_parser("list", help="List all epics")

    # Add epic
    epic_parser = subparsers.add_parser("add-epic", help="Add a new Epic")
    epic_parser.add_argument("name", help="Epic name (format: Epic X - Name)")
    epic_parser.add_argument("--objective", help="Epic objective")

    # Add story
    story_parser = subparsers.add_parser("add-story", help="Add a new Story to an Epic")
    story_parser.add_argument("epic", help="Epic name")
    story_parser.add_argument("name", help="Story name")
    story_parser.add_argument("--description", help="Story description")

    # Add task
    task_parser = subparsers.add_parser("add-task", help="Add a task to a Story")
    task_parser.add_argument("epic", help="Epic name")
    task_parser.add_argument("story", help="Story name")
    task_parser.add_argument("task", help="Task description")

    # Complete task
    complete_parser = subparsers.add_parser(
        "complete-task", help="Mark a task as completed"
    )
    complete_parser.add_argument("epic", help="Epic name")
    complete_parser.add_argument("story", help="Story name")
    complete_parser.add_argument("pattern", help="Task pattern to match")

    # Add to Terminées
    completed_parser = subparsers.add_parser(
        "completed", help="Add entry to Terminées section"
    )
    completed_parser.add_argument("description", help="Description of completed item")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return 1

    if args.command == "list":
        return 0 if list_epics() else 1
    elif args.command == "add-epic":
        return 0 if add_epic(args.name, args.objective) else 1
    elif args.command == "add-story":
        return 0 if add_story(args.epic, args.name, args.description) else 1
    elif args.command == "add-task":
        return 0 if add_task_to_story(args.epic, args.story, args.task) else 1
    elif args.command == "complete-task":
        return 0 if complete_task(args.epic, args.story, args.pattern) else 1
    elif args.command == "completed":
        return 0 if add_completed(args.description) else 1

    return 1


if __name__ == "__main__":
    sys.exit(main())
