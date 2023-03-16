"""Test cases for the import_ldefin2d_mischelpers function."""
import sys
from os import path

import pytest

from kompas_3d_wrapper import import_ldefin2d_mischelpers


KOMPAS_21_PYTHONWIN = (
    r"C:\ProgramData\ASCON\KOMPAS-3D\21\Python 3\App\Lib\site-packages\pythonwin"
)


def test_import_kompas_ldefin2d_mischelpers() -> None:
    """Test the import ldefin2d mischelpers function."""
    if path.exists(KOMPAS_21_PYTHONWIN):
        # Test if the function with valid input
        ldefin2d, mischelpers = import_ldefin2d_mischelpers(KOMPAS_21_PYTHONWIN)
        assert isinstance(ldefin2d, object)
        assert isinstance(mischelpers, object)


def test_invalid_path() -> None:
    """Test the import ldefin2d mischelpers function with invalid input path."""
    # Test if the function raises an error when provided with invalid input path
    with pytest.raises(FileNotFoundError) as excinfo:
        ldefin2d, mischelpers = import_ldefin2d_mischelpers(
            "C:/Users/ABC/Documents/kompas_pythonwin"
        )
    assert "Kompas pythonwin not found" in str(excinfo.value)


def test_wrong_path() -> None:
    """Test if the function raises ImportError when either modules can't be imported."""
    if path.exists(KOMPAS_21_PYTHONWIN + "/pywin"):
        with pytest.raises(ImportError) as excinfo:
            ldefin2d, mischelpers = import_ldefin2d_mischelpers(
                KOMPAS_21_PYTHONWIN + "/pywin"
            )
        assert "Failed importing" in str(excinfo.value)


def test_sys_path() -> None:
    """Test if the sys.path is restored after execution."""
    if path.exists(KOMPAS_21_PYTHONWIN):
        original_sys_path = sys.path
        ldefin2d, mischelpers = import_ldefin2d_mischelpers(KOMPAS_21_PYTHONWIN)
        assert sys.path == original_sys_path
