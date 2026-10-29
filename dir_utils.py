"""
Functions for working with directories, zips, and files.
"""

from pathlib import Path
import hashlib
import logging
import os
import platform
import tempfile
import zipfile

logger = logging.getLogger(__name__)

try:
    import send2trash
except ImportError:
    send2trash = None

try:
    import win32com.client
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False


def ensure_dir(path: str | Path) -> Path:
    """
    Ensures that the directory (or parent directory if a file path is passed) exists.

    Args:
        path: Directory path or file path.

    Returns:
        The resolved Path object of the directory.
    """
    path = Path(path).resolve()
    target_dir = path.parent if (path.is_file() or (not path.exists() and path.suffix)) else path

    try:
        if not target_dir.exists():
            target_dir.mkdir(parents=True)
            logger.debug("Created directory: \"%s\"", target_dir)

        return target_dir

    except OSError as e:
        logger.exception("Failed to create directory: \"%s\"", target_dir)
        raise e


def send_to_recycle_bin(path: str | Path) -> None:
    """
    Sends a file or directory to the OS recycle bin / trash.
    """
    path = Path(path)

    if not path.exists():
        raise FileNotFoundError(f"Path does not exist: \"{path}\"")

    try:
        if send2trash is None:
            raise ModuleNotFoundError("send_to_recycle_bin requires the send2trash package.")
        send2trash.send2trash(str(path))
        logger.info("Sent \"%s\" to recycle bin.", path)

    except Exception as e:
        logger.exception("Failed to send path to recycle bin: \"%s\"", path)
        raise e


def dir_is_empty(path: str | Path) -> bool:
    """
    Checks if a directory is empty efficiently.
    """
    path = Path(path)

    if not path.exists():
        raise FileNotFoundError(f"Path does not exist: \"{path}\"")

    if not path.is_dir():
        raise NotADirectoryError(f"Path is not a valid directory: \"{path}\"")

    try:
        is_empty = next(path.iterdir(), None) is None
        return is_empty

    except OSError:
        logger.exception("Failed to check if directory is empty: \"%s\"", path)
        raise


def get_file_data(file_path: str | Path) -> tuple[int, int, int]:
    """Gets file stats safely across different operating systems.

    Args:
        file_path (str or Path): The path to the file whose statistics are to be retrieved.

    Returns:
        tuple[int, int, int]: Modified time in nanoseconds, creation time when available (otherwise the platform's
            change time) in nanoseconds, and size in bytes.

    Raises:
        OSError: If an error occurs while retrieving the file data.
    """
    path = Path(file_path)

    try:
        stat = path.stat()

        if hasattr(stat, 'st_birthtime_ns'):
            modified_time, created_time = stat.st_mtime_ns, stat.st_birthtime_ns
        else:
            modified_time = stat.st_mtime_ns
            created_time = getattr(stat, 'st_ctime_ns', stat.st_mtime_ns)

        size_in_bytes = stat.st_size
        logger.debug("Retrieved file data for %s (size: %d bytes)", path, size_in_bytes)

        return modified_time, created_time, size_in_bytes

    except OSError as e:
        logger.error("Failed to get file data for %s: %s", path, e)
        raise


def zip_files(files: list[str | Path], zip_path: str | Path, compresslevel: int = 6, error_on_missing: bool = False, overwrite: bool = True) -> None:
    """
    Creates a zip archive from a list of file or directory paths safely using a temporary file.

    Args:
        files (list[str or Path]): A list of file or directory paths to be included in the zip archive.
        zip_path (str or Path): The path where the zip archive will be created.
        compresslevel (int, optional): Compression level for the zip archive. Defaults to 6.
        error_on_missing (bool, optional): Whether to raise an error if any source files are missing. Defaults to False.
        overwrite (bool, optional): Whether to overwrite an existing zip file at the specified path. Defaults to True.

    Returns:
        None

    Raises:
        FileExistsError: If the destination archive already exists and `overwrite=False`.
        FileNotFoundError: If a source path is missing and `error_on_missing=True`.
        Exception: For any other unexpected errors.
    """
    try:
        zip_path = Path(zip_path)
        if not overwrite and zip_path.exists():
            raise FileExistsError(f"Destination archive already exists and overwrite=False: {zip_path}")

        zip_path.parent.mkdir(parents=True, exist_ok=True)
        temp_file: Path | None = None

        try:
            with tempfile.NamedTemporaryFile(
                dir=zip_path.parent,
                prefix=f"{zip_path.name}.tmp_",
                delete=False
            ) as tmp_f:
                temp_file = Path(tmp_f.name)

                with zipfile.ZipFile(tmp_f, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=compresslevel) as zipf:
                    for src_path in files:
                        src_path = Path(src_path)

                        if not src_path.exists():
                            if error_on_missing:
                                raise FileNotFoundError(f"Source path missing: {src_path}")
                            logger.warning("Skipping missing source path: %s", src_path)
                            continue

                        if src_path.is_dir():
                            base_folder_name = src_path.name
                            for item in src_path.rglob("*"):
                                if item.is_file():
                                    relative_path = item.relative_to(src_path)
                                    item_arcname = Path(base_folder_name) / relative_path
                                    zipf.write(item, arcname=str(item_arcname))
                        else:
                            zipf.write(src_path, arcname=src_path.name)

            temp_file.replace(zip_path)
            logger.info("Successfully created archive: %s", zip_path)

        except Exception as e:
            logger.exception("Failed to create zip archive %s", zip_path)
            raise
        finally:
            if temp_file and temp_file.exists():
                try:
                    temp_file.unlink()
                except OSError:
                    logger.exception("Failed to clean up temporary zip file %s", temp_file)

    except Exception as e:
        logger.error("An error occurred: %s", str(e))
        raise


def get_files_list(root: str | Path, recursive: bool = False) -> list[Path]:
    """
    Retrieves a list of file paths from the specified directory.

    Args:
        root (str or Path): The path to the directory from which files are retrieved.
        recursive (bool, optional): If True, returns all files in subdirectories as well. Defaults to False.

    Returns:
        list[Path]: A list of all file paths found in the specified directory and its subdirectories if `recursive` is True.

    Raises:
        NotADirectoryError: Raised if the provided path does not exist or is not a directory.
    """
    root = Path(root)

    if not root.is_dir():
        raise NotADirectoryError(f"Expected a directory, got: {root}")

    try:
        files = []
        directories = [root]

        while directories:
            current_dir = directories.pop()
            with os.scandir(current_dir) as entries:
                for entry in entries:
                    if recursive and entry.is_dir(follow_symlinks=False):
                        directories.append(Path(entry.path))
                    elif entry.is_file():
                        files.append(Path(entry.path))

        logger.debug("Found %s files in %s (recursive=%s)", len(files), root, recursive)
        return files

    except Exception:
        logger.exception("Failed to get files list for %s", root)
        raise


def generate_hash(file_path: str | Path, algorithm: str = "sha256") -> str:
    """
    Generates a hexadecimal hash for a file using optimal performance methods.

    Args:
        file_path (str or Path): The path to the file whose hash is to be generated.
        algorithm (str, optional): The hash algorithm to use. Defaults to 'sha256'.

    Returns:
        str: A hexadecimal representation of the hash.

    Raises:
        FileNotFoundError: If the specified file does not exist.
        ValueError: If an invalid hash algorithm is provided.
    """
    try:
        file_path = Path(file_path)
        if not file_path.is_file():
            raise FileNotFoundError(f"File does not exist: {file_path}")

        logger.debug("Generating %s hash for %s...", algorithm, file_path)

        # Determine the appropriate hashing function based on Python version
        if hasattr(hashlib, 'file_digest'):
            try:
                with file_path.open("rb") as file:
                    hexd = hashlib.file_digest(file, algorithm).hexdigest()
            except ValueError as e:
                raise ValueError(f"Invalid hash algorithm: {e}") from e
        else:
            h = hashlib.new(algorithm)
            with file_path.open("rb") as f:
                for chunk in iter(lambda: f.read(1024 * 1024), b""):
                    h.update(chunk)
            hexd = h.hexdigest()

    except Exception as e:
        logger.exception("Failed to generate hash for %s", file_path)
        raise

    logger.debug("Generated hash %s for %s", hexd, file_path)
    return hexd


def extract_file_from_zip(zip_path: str | Path, file_name: str) -> bytes:
    """
    Extracts a specific file from a zip archive and returns it as bytes.

    Args:
        zip_path (str or Path): The path to the zip archive.
        file_name (str): The name of the file within the zip archive to extract.

    Returns:
        bytes: The content of the extracted file in binary format.

    Raises:
        FileNotFoundError: If the specified file is not found in the zip archive.
        Exception: For any other unexpected errors during extraction.
    """
    try:
        with Path(zip_path).open('rb') as zip_file:
            with zipfile.ZipFile(zip_file) as z:
                if file_name not in z.namelist():
                    raise FileNotFoundError(f"File not found in the zip archive: \"{file_name}\"")

                data = z.read(file_name)

        return data

    except Exception as e:
        logger.exception("Failed to extract or read file from zip")
        raise


def list_files_in_zip(zip_path: str | Path, directory: str) -> list[Path]:
    """
    Retrieves a list of files from a directory inside a zip archive.

    Args:
        zip_path (str | Path): Path to the zip archive.
        directory (str): Directory inside the archive to search.

    Returns:
        list[Path]: A list of file paths found inside the archive directory.

    Raises:
        FileNotFoundError:
            If the archive does not exist or the directory is not found
            within the archive.
        zipfile.BadZipFile:
            If the archive is not a valid zip file.
    """
    try:
        zip_path = Path(zip_path)

        if not zip_path.is_file():
            raise FileNotFoundError(f"Zip archive does not exist: {zip_path}")

        directory = directory.replace("\\", "/").rstrip("/") + "/"

        logger.debug("Listing files from %s under %s", zip_path, directory)
        with zipfile.ZipFile(zip_path) as z:
            names = z.namelist()

            if not any(name.startswith(directory) for name in names):
                raise FileNotFoundError(f'Directory not found in archive: "{directory}"')

            files = [Path(name) for name in names if name.startswith(directory) and not name.endswith("/")]

        logger.debug("Found %s files under %s in archive %s", len(files), directory, zip_path)

        return files

    except Exception:
        logger.exception("Failed to list files from archive %s under %s", zip_path, directory)
        raise


def get_windows_details(file_path: str | Path, max_columns: int = 512) -> dict[str, str]:
    """
    Returns Windows 'Details' tab properties using the Shell Property System (Windows-only).
    """
    p = Path(file_path)
    if not p.exists():
        raise FileNotFoundError(p)
    if not HAS_WIN32 or platform.system() != "Windows":
        raise OSError("get_windows_details requires Windows and the pywin32 library.")

    shell = win32com.client.Dispatch("Shell.Application")
    folder = shell.NameSpace(str(p.parent))
    if folder is None:
        raise RuntimeError(f"Shell.NameSpace failed for parent: {p.parent}")
    item = folder.ParseName(p.name)
    if item is None:
        raise RuntimeError(f"Shell.ParseName failed for item: {p}")

    props: dict[str, str] = {}
    blanks = 0

    for i in range(max_columns):
        header = folder.GetDetailsOf(None, i)
        if not header:
            blanks += 1
            if blanks >= 25:
                break
            continue
        blanks = 0

        value = folder.GetDetailsOf(item, i)
        if isinstance(value, str):
            value = value.strip()
        if value:
            props[str(header).strip()] = str(value)

    return props
