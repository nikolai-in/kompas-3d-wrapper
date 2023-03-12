"""Test import kompas_3d_wrapper."""
import kompas_3d_wrapper


def test_import_kompas_ldefin2d_mischelpers():
    """Test if the import_kompas_ldefin2d_mischelpers function is imported."""
    assert "import_kompas_ldefin2d_mischelpers" in dir(kompas_3d_wrapper)
