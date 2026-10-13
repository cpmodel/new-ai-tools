from cpat_model.components.carbon_pricing.existing_ets import (
    ExistingETS
)

from cpat_model.components.carbon_pricing.existing_ct import load_do_ct_ets_exist


def test_set_eu_ets():
    obj = ExistingETS.__new__(ExistingETS)
    set_eu_ets = obj._ExistingETS__set_eu_ets # pylint: disable=protected-access

    # No countries:
    selected_countries = ['ZZZ', 'AUS', 'AAA']
    do_ct_ets_exist = load_do_ct_ets_exist(selected_countries)

    set_eu_ets(do_ct_ets_exist)
    assert isinstance(obj.eu_ets, list)
    assert not obj.eu_ets


def test_set_do_ets_forecast():
    obj = ExistingETS.__new__(ExistingETS)
    set_do_ets_forecast = obj._ExistingETS__set_do_ets_forecast # pylint: disable=protected-access

    set_do_ets_forecast(True)
    assert obj.do_ets_forecast is True

    set_do_ets_forecast(False)
    assert obj.do_ets_forecast is False
