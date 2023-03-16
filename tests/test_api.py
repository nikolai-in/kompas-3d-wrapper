"""Test kompas_3d_wrapper api."""
import time

from kompas_3d_wrapper import get_kompas_api5
from kompas_3d_wrapper import get_kompas_api7
from kompas_3d_wrapper import get_kompas_constants
from kompas_3d_wrapper import start_kompas


def test_kompas_api7() -> None:
    """Test kompas api7."""
    is_running: bool = start_kompas()

    time.sleep(5)

    const = get_kompas_constants()
    _, api7 = get_kompas_api7()

    app7 = api7.Application
    app7.Visible = True
    app7.HideMessage = const.ksHideMessageNo

    app_name = app7.ApplicationName(FullName=True)

    if not is_running:
        app7.Quit()

    assert (
        "КОМПАС-3D v21 Учебная версия (Не для коммерческого использования)" in app_name
    )


def test_kompas_api5() -> None:
    """Test kompas api5."""
    is_running: bool = start_kompas()

    time.sleep(5)

    _, api5 = get_kompas_api5()

    api7 = api5.ksGetApplication7()

    app_name = api7.ApplicationName(FullName=True)

    if not is_running:
        api5.Quit()

    assert (
        "КОМПАС-3D v21 Учебная версия (Не для коммерческого использования)" in app_name
    )
