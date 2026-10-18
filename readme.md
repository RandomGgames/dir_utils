# Text File Utils

A Python module for directories, files, and zips.

## Usage

This repository is intended to be included in other projects as a Git submodule.

From the root directory of an existing project, run:

```bash
git submodule add https://github.com/RandomGgames/dir_utils
```

This adds the repository to your project.

You can then import the utilities directly:

```python
from dir_utils import send_to_recycle_bin
```

## Cloning a Project with the Submodule

When cloning a project that contains this repository as a submodule, use:

```bash
git clone --recurse-submodules <project-repository-url>
```

If the project has already been cloned, initialize the submodule with:

```bash
git submodule update --init --recursive
```

## Updating the Utilities

The parent project tracks a specific commit of this repository. To update to the latest remote commit:

```bash
git submodule update --remote
```
