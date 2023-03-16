"""Тесты для различных функций."""
from os import path

import pytest

from kompas_3d_wrapper import find_exe_by_file_extension
from kompas_3d_wrapper import get_pythonwin_path


def test_find_exe_by_file_extension_success():
    """Проверка успешного поиска исполняемого файла по расширению файла."""
    bat = find_exe_by_file_extension(".bat")
    assert "%1" in bat


def test_find_exe_by_file_extension_nonexistent_extension_error():
    """Проверка ошибки при поиске по несуществующему расширению файла."""
    with pytest.raises(FileNotFoundError):
        find_exe_by_file_extension(".IHATEKOMPAS")


def test_get_pythonwin():
    """Проверка получения пути к pythonwin."""
    cdm = find_exe_by_file_extension(".cdm")
    cdm = cdm.replace(path.basename(cdm), "")
    pythonwin = get_pythonwin_path()
    assert path.basename(cdm) in pythonwin
