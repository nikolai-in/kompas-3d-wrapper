"""Test for the kompas_3d_wrapper test_start_kompas_not_running function."""
from os import path

from kompas_3d_wrapper import start_kompas_if_not_running


KOMPAS_21_DIR = r"C:\Program Files\ASCON\KOMPAS-3D v21 Study\Bin"
KOMPAS_21_EXECUTABLE = "kStudy.exe"


def test_start_kompas_not_running() -> None:
    """Test for the function start_kompas_if_not_running."""
    if path.exists(KOMPAS_21_DIR):
        result = start_kompas_if_not_running(KOMPAS_21_DIR, KOMPAS_21_EXECUTABLE)

        # check that the function returns True/False correctly
        assert isinstance(result, bool)
