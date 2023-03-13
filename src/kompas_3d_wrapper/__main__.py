"""Command-line interface."""
import time
from os import path

import click

from kompas_3d_wrapper.main import get_kompas_api7
from kompas_3d_wrapper.main import get_kompas_constants
from kompas_3d_wrapper.main import start_kompas


KOMPAS_21_DIR = r"C:\Program Files\ASCON\KOMPAS-3D v21 Study\Bin"
KOMPAS_21_EXECUTABLE = "kStudy.exe"
KOMPAS_21_PYTHONWIN = (
    r"C:\ProgramData\ASCON\KOMPAS-3D\21\Python 3\App\Lib\site-packages\pythonwin"
)


@click.command()
@click.version_option()
def main() -> None:
    """Kompas 3D Wrapper."""
    if not path.exists(KOMPAS_21_PYTHONWIN):
        print("Kompas 3D not found. Please install Kompas 3D with macro support.")
        return

    try:
        is_running: bool = start_kompas(path.join(KOMPAS_21_DIR, KOMPAS_21_EXECUTABLE))

        time.sleep(5)

        module7, api7 = get_kompas_api7()
        const = get_kompas_constants()
        app7 = api7.Application
        app7.Visible = True
        app7.HideMessage = const.ksHideMessageNo

        print(f"Application Name: {app7.ApplicationName(FullName=True)}")

        if not is_running:
            app7.Quit()

    except Exception as e:
        print(f"Error occurred: {e}")


if __name__ == "__main__":
    main(prog_name="kompas-3d-wrapper")  # pragma: no cover
