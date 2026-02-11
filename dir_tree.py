import os

# Set the root directory of your Django project
ROOT_DIR = "D:\\deck_automator"  # <-- Replace with your actual path

# Folders you want to skip diving into (but still show their names)
SKIP_DIRS = {
    "__pycache__", "locale", "zoneinfo", "pip", "db", "cache",
    "checks", "Include", "Lib", "Scripts"
}

def print_tree(current_path, prefix=""):
    try:
        entries = sorted(os.listdir(current_path))
    except PermissionError:
        return

    for index, entry in enumerate(entries):
        full_path = os.path.join(current_path, entry)
        is_last = (index == len(entries) - 1)
        branch = "└── " if is_last else "├── "
        new_prefix = "    " if is_last else "│   "

        if os.path.isdir(full_path):
            print(prefix + branch + entry + "/")
            if entry in SKIP_DIRS:
                continue  # Show the folder, but don't go inside
            print_tree(full_path, prefix + new_prefix)
        else:
            print(prefix + branch + entry)

# Run it
print(f"Project structure for: {ROOT_DIR}\n")
print_tree(ROOT_DIR)
