"""Main module."""
import logging
import shlex
import subprocess  # noqa: S404
import sys
import winreg
from os import path
from types import ModuleType

import psutil
import pythoncom
from win32com.client import Dispatch
from win32com.client import gencache


def find_exe_by_file_extension(file_extension: str) -> str:
    """Находит исполняемый файл по расширению файла.

    Args:
        file_extension (str): Расширение файла

    Returns:
        str: Путь к исполняемому файлу

    Raises:
        FileNotFoundError: Если не удалось найти исполняемый файл

    Example:
    >>> try:
    >>>    cdm = find_exe_by_file_extension('.cdm')
    >>>    assert 'KOMPAS-3D'in cdm
    >>> except FileNotFoundError:
    >>>    logging.warning('КОМПАС-3D с поддержкой макросов не установлен')
    >>>    pass
    """
    try:
        with winreg.OpenKey(
            winreg.HKEY_CLASSES_ROOT, file_extension, 0, winreg.KEY_READ
        ) as key:
            class_name, key_type = winreg.QueryValueEx(key, "")

            if key_type != winreg.REG_SZ:
                raise FileNotFoundError(f"Нет ключа {file_extension} в реестре")

        with winreg.OpenKey(
            winreg.HKEY_CLASSES_ROOT,
            rf"{class_name}\shell\open\command",
            0,
            winreg.KEY_READ,
        ) as key:
            command_val, _ = winreg.QueryValueEx(key, "")
            exe_path = shlex.split(command_val)[0]
            return exe_path.strip('"')
    except FileNotFoundError as file_not_found_err:
        raise FileNotFoundError(
            f"Не удалось найти исполняемый файл по расширению файла {file_extension!r}"
        ) from file_not_found_err


def get_pythonwin_path() -> str:
    """Возвращает путь к папке pythonwin КОМПАС-3D.

    Пытается определить путь на основе расширения файла '.cdm'.

    Returns:
        str: Путь к папке pythonwin КОМПАС-3D

    Raises:
        FileNotFoundError: Если не удалось найти папку pythonwin КОМПАС-3D
        FileNotFoundError: Если не удалось найти библиотеки python компаса

    Example:
        >>> pythonwin = get_pythonwin_path()
        >>> assert 'Python 3' in pythonwin
    """
    file_ext = ".cdm"
    py_scripter_path = find_exe_by_file_extension(file_ext)

    if "Python 3" not in py_scripter_path:
        raise FileNotFoundError("Библиотеки Python КОМПАС-3D не найдены")

    pythonwin_path = path.join(
        py_scripter_path.replace(path.basename(py_scripter_path), ""),
        "App\\Lib\\site-packages\\pythonwin",
    )

    if not path.exists(pythonwin_path):
        raise FileNotFoundError("папка pythonwin КОМПАС-3D не найдена")

    logging.debug(f"Pythonwin: {pythonwin_path}")

    return pythonwin_path


pythonwin_path = get_pythonwin_path()


def import_ldefin2d_mischelpers(
    kompas_pythonwin: str = pythonwin_path,
) -> tuple[ModuleType, ModuleType]:
    """Импортирует модули LDefin2D и MiscellaneousHelpers из папки pythonwin КОМПАС-3D.

    Args:
        kompas_pythonwin (str, optional): Путь к папке pythonwin КОМПАС-3D

    Returns:
        tuple[ModuleType, ModuleType]: Модули LDefin2D, MiscellaneousHelpers

    Raises:
        FileNotFoundError: Если папка pythonwin не найдена
        ImportError: Если не удалось импортировать модули LDefin2D или miscHelpers

    Example:
        >>> LDefin2D, miscHelpers = import_ldefin2d_mischelpers()
        >>> assert  'ALL_OBJ' in dir(LDefin2D)
        >>> assert  'Dispatch' in dir(miscHelpers)

    """
    # Проверяем, существует ли папка pythonwin
    if not path.exists(kompas_pythonwin):
        raise FileNotFoundError(
            f"Kompas pythonwin not found at {kompas_pythonwin!r}. "
            + "Please install Kompas with macro support."
        )

    sys.path.append(kompas_pythonwin)

    # Удаляем модули LDefin2D и MiscellaneousHelpers из sys.modules, если они там есть
    sys.modules.pop("MiscellaneousHelpers", None)
    sys.modules.pop("LDefin2D", None)

    try:
        # Импортируем модули LDefin2D и MiscellaneousHelpers как miscHelpers
        import LDefin2D  # pyright: reportMissingImports=false
        import MiscellaneousHelpers as miscHelpers
    except ImportError as import_err:
        # Если не удалось импортировать модули, то возвращаем исключение
        raise ImportError(f"Failed importing {import_err.name}") from ImportError
    finally:
        # Убеждаемся, что папка pythonwin удалена из sys.path
        sys.path.remove(kompas_pythonwin)

    return LDefin2D, miscHelpers


def get_kompas_constants() -> ModuleType:
    """Get KOMPAS-3D constants."""
    try:
        constants = gencache.EnsureModule(
            "{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0
        ).constants
        return constants
    except Exception as e:
        raise Exception("Failed to get Kompas constants: " + str(e)) from Exception


def get_kompas_constants_3d() -> ModuleType:
    """Get KOMPAS-3D constants."""
    try:
        constants = gencache.EnsureModule(
            "{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0
        ).constants
        return constants
    except Exception as e:
        raise Exception("Failed to get Kompas constants 3d: " + str(e)) from Exception


def get_kompas_api7() -> tuple[any, type]:
    """Get KOMPAS-3D COM API version 7."""
    try:
        module = gencache.EnsureModule(
            "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0
        )
        api = module.IKompasAPIObject(
            Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(
                module.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch
            )
        )
        return module, api
    except Exception as e:
        raise Exception(
            "Failed to get Kompas COM API version 7: " + str(e)
        ) from Exception


def get_kompas_api5() -> tuple[any, type]:
    """Get KOMPAS-3D COM API version 5."""
    try:
        module = gencache.EnsureModule(
            "{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0
        )
        api = module.KompasObject(
            Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(
                module.KompasObject.CLSID, pythoncom.IID_IDispatch
            )
        )
        return module, api
    except Exception as e:
        raise Exception(
            "Failed to get Kompas COM API version 5: " + str(e)
        ) from Exception


def is_process_running(process_name: str) -> bool:
    """Check if a process with a given name is currently running."""
    for proc in psutil.process_iter():
        try:
            if proc.name() == process_name:
                logging.debug(f"Process {process_name} is running")
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    logging.debug(f"Process {process_name} is not running")
    return False


def start_kompas(kompas_exe_path: str) -> bool:
    """Start KOMPAS-3D if it's not already running."""
    if is_process_running(path.basename(kompas_exe_path)):
        return True

    subprocess.Popen(kompas_exe_path)  # noqa: S603
    logging.debug(f"Started Kompas at {kompas_exe_path}")
    return False
