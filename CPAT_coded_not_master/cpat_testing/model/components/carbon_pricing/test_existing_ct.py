from cpat_model.components.carbon_pricing.existing_ct import ExistingCT

def test_set_do_ct_forecast():
    obj = ExistingCT.__new__(ExistingCT)  # creates instance without calling __init__
    set_do_ct_forecast = obj._ExistingCT__set_do_ct_forecast # pylint: disable=protected-access

    set_do_ct_forecast(True)
    assert obj.do_ct_forecast is True

    set_do_ct_forecast(False)
    assert obj.do_ct_forecast is False
