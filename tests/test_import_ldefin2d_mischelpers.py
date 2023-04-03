"""Test cases for the import_ldefin2d_mischelpers function."""
import sys

import pytest

from kompas_3d_wrapper import Kompas


kompas = Kompas()


def test_import_kompas_ldefin2d_mischelpers() -> None:
    """Test the import ldefin2d mischelpers function."""
    # Test if the function with valid input
    ldefin2d, mischelpers = kompas.import_ldefin2d_mischelpers()
    assert isinstance(ldefin2d, object)
    assert isinstance(mischelpers, object)


def test_invalid_path() -> None:
    """Test the import ldefin2d mischelpers function with invalid input path."""
    # Test if the function raises an error when provided with invalid input path
    with pytest.raises(FileNotFoundError) as excinfo:
        kompas.import_ldefin2d_mischelpers("C:/Users/ABC/Documents/kompas_pythonwin")
    assert "Kompas pythonwin not found" in str(excinfo.value)


def test_sys_path() -> None:
    """Test if the sys.path is restored after execution."""
    original_sys_path = sys.path
    kompas.import_ldefin2d_mischelpers()
    assert sys.path == original_sys_path
