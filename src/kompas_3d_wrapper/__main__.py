"""Command-line interface."""
import sys
from os import path

import click


KOMPAS_21_STUDY = r"C:\Program Files\ASCON\KOMPAS-3D v21 Study\Bin\kStudy.exe"
KOMPAS_21_PYTHONWIN = (
    r"C:\ProgramData\ASCON\KOMPAS-3D\21\Python 3\App\Lib\site-packages\pythonwin"
)


def import_kompas_ldefin2d_mischelpers(kompas_pythonwin: str) -> tuple:
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


@click.command()
@click.version_option()
def main() -> None:
    """Kompas 3D Wrapper."""
    ldefin2d, misc_helpers = import_kompas_ldefin2d_mischelpers(KOMPAS_21_PYTHONWIN)
    print(type(ldefin2d))
    print(type(misc_helpers))


if __name__ == "__main__":
    main(prog_name="kompas-3d-wrapper")  # pragma: no cover
