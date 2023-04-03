"""Test kompas_3d_wrapper api."""
from kompas_3d_wrapper import Kompas


def test_kompas_api7() -> None:
    """Test kompas api7."""
    kompas = Kompas()

    const = kompas.constants
    api7 = kompas.api7

    app7 = api7.Application
    app7.Visible = True
    app7.HideMessage = const.ksHideMessageNo

    app_name = app7.ApplicationName(FullName=True)

    if not kompas.was_running:
        app7.Quit()

    assert (
        "КОМПАС-3D v21 Учебная версия (Не для коммерческого использования)" in app_name
    )


def test_kompas_api5() -> None:
    """Test kompas api5."""
    kompas = Kompas()

    _, api5 = kompas.get_kompas_api5()

    api7 = api5.ksGetApplication7()

    app_name = api7.ApplicationName(FullName=True)

    if not kompas.was_running:
        api5.Quit()

    assert (
        "КОМПАС-3D v21 Учебная версия (Не для коммерческого использования)" in app_name
    )
