"""Command-line interface."""
import subprocess  # noqa: S404
import sys
import time
from os import path
from types import ModuleType

import click
import psutil
import pythoncom
from win32com.client import Dispatch
from win32com.client import gencache


KOMPAS_21_DIR = r"C:\Program Files\ASCON\KOMPAS-3D v21 Study\Bin"
KOMPAS_21_EXECUTABLE = "kStudy.exe"
KOMPAS_21_PYTHONWIN = (
    r"C:\ProgramData\ASCON\KOMPAS-3D\21\Python 3\App\Lib\site-packages\pythonwin"
)


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


def get_kompas_api7() -> tuple[any, type, ModuleType]:
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
        const = gencache.EnsureModule(
            "{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0
        ).constants
        return module, api, const
    except Exception as e:
        raise Exception(
            "Failed to get Kompas COM API version 7: " + str(e)
        ) from Exception


def get_kompas_api5() -> tuple[any, type, ModuleType]:
    """Get KOMPAS-3D COM API version 5."""
    try:
        module = gencache.EnsureModule(
            "{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0
        )
        api = module.IKompasAPIObject(
            Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(
                module.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch
            )
        )
        const = gencache.EnsureModule(
            "{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0
        ).constants
        return module, api, const
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
        return False

    subprocess.Popen(kompas_exe_path)  # noqa: S603
    return True


@click.command()
@click.version_option()
def main() -> None:
    """Kompas 3D Wrapper."""
    if path.exists(KOMPAS_21_PYTHONWIN):
        try:
            is_running: bool = not start_kompas(
                path.join(KOMPAS_21_DIR, KOMPAS_21_EXECUTABLE)
            )

            time.sleep(5)

            module7, api7, const7 = get_kompas_api7()  # Подключаемся к API7
            app7 = api7.Application  # Получаем основной интерфейс
            app7.Visible = True  # Показываем окно пользователю (если скрыто)
            app7.HideMessage = (
                const7.ksHideMessageNo  # Отвечаем НЕТ на любые вопросы программы
            )
            print(f"Application Name: {app7.ApplicationName(FullName=True)}")

            if not is_running:
                app7.Quit()
        except Exception as e:
            print(f"Error occurred: {e}")
    else:
        print("Kompas 3D not found. Please install Kompas 3D with macro support.")


if __name__ == "__main__":
    main(prog_name="kompas-3d-wrapper")  # pragma: no cover
