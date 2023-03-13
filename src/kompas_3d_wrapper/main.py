"""Main module."""
import subprocess  # noqa: S404
import sys
from os import path
from types import ModuleType

import psutil
import pythoncom
from win32com.client import Dispatch
from win32com.client import gencache


def import_kompas_ldefin2d_mischelpers(
    kompas_pythonwin: str,
) -> tuple[ModuleType, ModuleType]:
    """Import Kompas LDefin2D and miscHelpers modules."""
    # Check if the kompas_pythonwin path exists, otherwise raise FileNotFoundError
    if not path.exists(kompas_pythonwin):
        raise FileNotFoundError(
            "Kompas pythonwin not found. Please install Kompas with macro support."
        ) from None

    sys.path.append(kompas_pythonwin)

    # Test if MiscellaneousHelpers and LDefin2D already exist in the sys.modules
    if "MiscellaneousHelpers" in sys.modules:
        # If the MiscellaneousHelpers module already exists, remove it
        del sys.modules["MiscellaneousHelpers"]
    if "LDefin2D" in sys.modules:
        # If the LDefin2D module already exists remove it
        del sys.modules["LDefin2D"]

    try:
        # Import LDefin2D and MiscellaneousHelpers as miscHelpers modules
        import LDefin2D  # pyright: reportMissingImports=false
        import MiscellaneousHelpers as miscHelpers
    except ImportError:
        # If either of the modules above cannot be imported, raise the same ImportError
        raise ImportError(
            "Kompas LDFine2D and miscHelpers modules not found."
        ) from ImportError
    finally:
        # Ensure that the sys.path is restored
        sys.path.remove(kompas_pythonwin)

    # Return the imported modules as a tuple after the try-finally block has completed
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
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return False


def start_kompas(kompas_exe_path: str) -> bool:
    """Start KOMPAS-3D if it's not already running."""
    if is_process_running(path.basename(kompas_exe_path)):
        return True

    subprocess.Popen(kompas_exe_path)  # noqa: S603
    return False
