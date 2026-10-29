# dir_utils

Small Python utilities for directories, files, and ZIP archives. The module requires Python 3.10 or later.

## Installation

This repository is intended to be included in other projects as a Git submodule. From the root of the parent project:

```bash
git submodule add https://github.com/RandomGgames/dir_utils
```

Install `send2trash` to use `send_to_recycle_bin`:

```bash
python -m pip install send2trash
```

On Windows, install `pywin32` as well to use `get_windows_details`. Both dependencies are only needed for those helpers; the rest of the module can be used without them.

## Usage

Import utilities from the package:

```python
from dir_utils import ensure_dir, generate_hash, get_files_list, zip_files

files = get_files_list("input", recursive=True)
zip_files(files, "archive.zip")
digest = generate_hash("archive.zip")
```

## Utilities

- `ensure_dir(path)` creates the target directory and returns its resolved `Path`. Existing files and non-existing paths with a suffix are treated as file paths, so their parent directory is created.
- `send_to_recycle_bin(path)` sends an existing file or directory to the operating system's trash. Requires `send2trash`.
- `dir_is_empty(path)` checks whether an existing directory is empty. Raises `FileNotFoundError` or `NotADirectoryError` for invalid paths.
- `get_file_data(path)` returns `(modified_time_ns, creation_or_change_time_ns, size_bytes)`. The second value is the platform's creation time where available; otherwise it is the platform's change time.
- `zip_files(files, zip_path, compresslevel=6, error_on_missing=False, overwrite=True)` creates or replaces a ZIP archive. Missing inputs are skipped by default; set `error_on_missing=True` to raise an error. Files are stored by basename, and directory contents are stored under the directory's basename.
- `get_files_list(root, recursive=False)` returns file paths from a directory. An empty directory returns an empty list.
- `generate_hash(file_path, algorithm="sha256")` returns a file's hexadecimal digest.
- `extract_file_from_zip(zip_path, file_name)` returns the named archive member's contents as bytes.
- `list_files_in_zip(zip_path, directory)` returns `Path` objects for files under the named archive directory.
- `get_windows_details(file_path, max_columns=512)` returns non-empty Windows Shell Details properties as a dictionary. Requires Windows and `pywin32`.

## Submodule Maintenance

Clone a parent project and its submodules with:

```bash
git clone --recurse-submodules <project-repository-url>
```

For a project already cloned, initialize its submodules with:

```bash
git submodule update --init --recursive
```

To move this submodule to the latest remote commit:

```bash
git submodule update --remote
```
