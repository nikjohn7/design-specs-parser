"""Workbook loading and validation utilities.

This module provides safe loading of Excel workbooks with proper
error handling for invalid, corrupt, or unsupported files.
"""

import io
import zipfile
from typing import TYPE_CHECKING

import openpyxl
from openpyxl import Workbook
from openpyxl.utils.exceptions import InvalidFileException

if TYPE_CHECKING:
    pass


class WorkbookLoadError(Exception):
    """Exception raised when a workbook cannot be loaded.

    This exception is raised for various loading failures including:
    - Empty file data
    - Corrupt or invalid ZIP structure
    - Invalid Excel file format
    - Password-protected files
    - Other unexpected errors during loading

    Attributes:
        message: Human-readable error description
        detail: Additional technical details (optional)
    """

    def __init__(self, message: str, detail: str | None = None):
        """Initialize WorkbookLoadError.

        Args:
            message: Brief error message describing what went wrong
            detail: Additional error details or technical information
        """
        self.message = message
        self.detail = detail
        super().__init__(message)

    def __str__(self) -> str:
        """Return string representation of the error."""
        if self.detail:
            return f"{self.message}: {self.detail}"
        return self.message


def load_workbook_safe(file_bytes: bytes) -> Workbook:
    """Safely load an Excel workbook from raw bytes.

    This function validates and loads an Excel (.xlsx) file from bytes,
    handling various error conditions gracefully.

    Args:
        file_bytes: Raw bytes of the Excel file

    Returns:
        Workbook: Loaded openpyxl Workbook object

    Raises:
        WorkbookLoadError: If the file cannot be loaded due to:
            - Empty file data
            - Corrupt or invalid ZIP structure (xlsx files are ZIP archives)
            - Invalid Excel file format
            - Password-protected files
            - Other unexpected errors

    Example:
        >>> with open("schedule.xlsx", "rb") as f:
        ...     file_bytes = f.read()
        >>> wb = load_workbook_safe(file_bytes)
        >>> print(wb.sheetnames)
        ['Sheet1', 'Sheet2']
    """
    # Validate input is not empty
    if not file_bytes:
        raise WorkbookLoadError(
            message="Empty file",
            detail="The uploaded file contains no data"
        )

    # Check minimum file size (ZIP files need at least a few bytes for header)
    # A valid xlsx file should be at least ~100 bytes (empty workbook)
    if len(file_bytes) < 100:
        raise WorkbookLoadError(
            message="Invalid file",
            detail=f"File too small ({len(file_bytes)} bytes) to be a valid Excel workbook"
        )

    # Create a file-like object from bytes
    file_stream = io.BytesIO(file_bytes)

    try:
        # Load workbook with openpyxl
        # data_only=False preserves formulas (needed for schedule name extraction)
        # read_only=False allows full access to merged cells and other features
        workbook = openpyxl.load_workbook(
            file_stream,
            data_only=False,
            read_only=False,
        )
        return workbook

    except zipfile.BadZipFile as e:
        # xlsx files are ZIP archives - this error means corrupt or not xlsx
        raise WorkbookLoadError(
            message="Invalid file format",
            detail="File is not a valid Excel workbook (corrupt or not .xlsx format)"
        ) from e

    except InvalidFileException as e:
        # openpyxl-specific validation failure
        error_str = str(e).lower()

        # Check for password protection
        if "password" in error_str or "encrypted" in error_str:
            raise WorkbookLoadError(
                message="Password-protected file",
                detail="Cannot open password-protected Excel files"
            ) from e

        # Generic invalid file
        raise WorkbookLoadError(
            message="Invalid Excel file",
            detail=str(e)
        ) from e

    except PermissionError as e:
        # Unlikely with BytesIO but handle defensively
        raise WorkbookLoadError(
            message="Permission error",
            detail="Unable to read file data"
        ) from e

    except MemoryError as e:
        # File too large to load into memory
        raise WorkbookLoadError(
            message="File too large",
            detail="The file is too large to process"
        ) from e

    except KeyError as e:
        # Missing required Excel file components (e.g., [Content_Types].xml)
        # This happens when the file is a valid ZIP but not an Excel file
        raise WorkbookLoadError(
            message="Invalid Excel file",
            detail="File is a valid ZIP archive but not a valid Excel workbook (missing required components)"
        ) from e

    except Exception as e:
        # Catch-all for unexpected errors
        # Log the actual exception type for debugging
        error_type = type(e).__name__
        raise WorkbookLoadError(
            message="Failed to load workbook",
            detail=f"Unexpected error ({error_type}): {str(e)}"
        ) from e
