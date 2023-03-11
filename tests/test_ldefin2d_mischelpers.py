"""Test cases for the import_kompas_ldefin2d_mischelpers function."""
import os

import pytest

from kompas_3d_wrapper.__main__ import import_kompas_ldefin2d_mischelpers


KOMPAS_21_PYTHONWIN = (
    r"C:\ProgramData\ASCON\KOMPAS-3D\21\Python 3\App\Lib\site-packages\pythonwin"
)


def test_import_kompas_ldefin2d_mischelpers():
    """Test the import_kompas_ldefin2d_mischelpers function."""
    if os.path.exists(KOMPAS_21_PYTHONWIN):
        # Test if the function with valid input
        ldefin2d, mischelpers = import_kompas_ldefin2d_mischelpers(KOMPAS_21_PYTHONWIN)
        assert isinstance(ldefin2d, object)
        assert isinstance(mischelpers, object)

    # Test if the function raises an error when provided with invalid input path
    with pytest.raises(FileNotFoundError) as excinfo:
        ldefin2d, mischelpers = import_kompas_ldefin2d_mischelpers(
            "C:/Users/ABC/Documents/kompas_pythonwin"
        )
    assert "Kompas pythonwin not found" in str(excinfo.value)

    if os.path.exists(KOMPAS_21_PYTHONWIN + "/pywin"):
        # Test if the function raises ImportError exception
        # when either of the modules cannot be imported
        with pytest.raises(ImportError) as excinfo:
            ldefin2d, mischelpers = import_kompas_ldefin2d_mischelpers(
                KOMPAS_21_PYTHONWIN + "/pywin"
            )
        assert "Kompas LDFine2D and miscHelpers modules not found" in str(excinfo.value)

    if os.path.exists(KOMPAS_21_PYTHONWIN):
        # Test if sys.path is restored after execution
        original_sys_path = os.sys.path
        ldefin2d, mischelpers = import_kompas_ldefin2d_mischelpers(KOMPAS_21_PYTHONWIN)
        assert os.sys.path == original_sys_path
