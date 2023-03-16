"""Command-line interface."""
import time

import click

from kompas_3d_wrapper.main import get_kompas_api7
from kompas_3d_wrapper.main import get_kompas_constants
from kompas_3d_wrapper.main import start_kompas


@click.command()
@click.version_option()
def main() -> None:
    """Kompas 3D Wrapper."""
    try:
        is_running: bool = start_kompas()

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
