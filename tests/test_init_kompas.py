"""Test import kompas_3d_wrapper."""
import kompas_3d_wrapper


def test_import_modules():
    """Test if functions are imported."""
    for function in [
        "import_kompas_ldefin2d_mischelpers",
        "start_kompas_if_not_running",
        "get_kompas_api5",
        "get_kompas_api7",
    ]:
        assert function in dir(kompas_3d_wrapper)
