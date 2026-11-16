import pytest

from archive_service import folder_name, pst_name


def test_pst_and_folder_names():
    assert pst_name("2024") == "ADR.pst"
    assert pst_name("2021") == "ADR_2021.pst"
    assert folder_name("2022") == "Archives 2022"


def test_rejects_unknown_year():
    with pytest.raises(ValueError):
        pst_name("2020")
