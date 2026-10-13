import pandas as pd
import numpy as np

from cpat_model.inputs.dashboard_inputs import EnergySectorReform
import cpat_model.constants as c

class Phaseouts:
    """
    'Phaseouts of existing policies (subsidies, taxes, and price controls)' from
    'Policies' section (v361)
    CPAT Excel: table start baseline: 1816, scenario: 6806

    No 'Phaseout factor - PPAs (engineer model)' implemented.
    """
    # Public attributes:
    phaseout_producer: pd.DataFrame
    phaseout_consumer: pd.DataFrame
    phaseout_subsidy_tax: pd.DataFrame

    def __init__(
            self,
            simulation_years: tuple[int, int],
            scenario_type: str,
            # Dashboard inputs:
            energy_sector_reform: EnergySectorReform,
            is_pc_phaseout_baseline: bool,
            ) -> None:
        self.__years = np.arange(simulation_years[0] - 1, simulation_years[1] + 1)
        self.__columns = self.__years.astype(str)

        self.phaseout_producer = self.__get_phaseout_producer(energy_sector_reform)
        self.phaseout_consumer = self.__get_phaseout_consumer(energy_sector_reform)
        self.phaseout_subsidy_tax = self.__get_phaseout_subsidy_tax(
            scenario_type, is_pc_phaseout_baseline, energy_sector_reform
        )


    def __get_phaseout_producer(
            self,
            energy_sector_reform: EnergySectorReform
            ) -> pd.DataFrame:
        """
        'Phaseout factor - fossil fuels (producer)' (v361)
        CPAT Excel: baseline 1817, scenario 6807

        Scenario specific
        returns dims <simulation_years[0] - 1, simulation_years[1]>
        """
        neagative_share = 1.0 - energy_sector_reform['ffs_phaseout_share_prod']
        max_value = max(neagative_share, 1.0)

        phaseout_producer = pd.DataFrame(
            [np.full(len(self.__columns), max_value, dtype=float)],
            columns=self.__columns
        )

        if not energy_sector_reform['is_ffs_phaseout_prod']:
            return phaseout_producer

        decrease_each_year = (
            - energy_sector_reform['ffs_phaseout_share_prod']
            / energy_sector_reform['ffs_phaseout_prod']
        )

        mask = self.__years > energy_sector_reform['ffs_start_prod']
        values = (
            1 + decrease_each_year * (self.__years[mask] - energy_sector_reform['ffs_start_prod'])
        )
        # for years > than ffs_start_prod:
        #phaseout_producer = phaseout_producer.copy()
        #phaseout_producer.values[0, mask] = np.maximum(values, neagative_share)
        phaseout_producer = phaseout_producer.copy()
        phaseout_producer.iloc[0, mask] = np.maximum(values, neagative_share)
        return phaseout_producer


    def __get_phaseout_consumer(
            self,
            energy_sector_reform: EnergySectorReform
            ) -> pd.DataFrame:
        """
        'Phaseout factor - fossil fuels (consumer)' (v361)
        CPAT Excel: baseline 1818, scenario 6808

        Scenario specific
        returns dims <simulation_years[0] - 1, simulation_years[1]>
        """
        neagative_share = 1.0 - energy_sector_reform['ffs_phaseout_share_cons']
        max_value = max(neagative_share, 1.0)

        phaseout_consumer = pd.DataFrame(
            [np.full(len(self.__columns), max_value, dtype=float)],
            columns=self.__columns
        )

        if not energy_sector_reform['is_ffs_phaseout_cons']:
            return phaseout_consumer

        decrease_each_year = (
            - energy_sector_reform['ffs_phaseout_share_cons']
            / energy_sector_reform['ffs_phaseout_cons']
        )

        mask = self.__years > energy_sector_reform['ffs_start_cons']
        values = (
            1 + decrease_each_year * (self.__years[mask] - energy_sector_reform['ffs_start_cons'])
        )
        # for years > than ffs_start_cons:

        phaseout_consumer = phaseout_consumer.copy()
        phaseout_consumer.iloc[0, mask] = np.maximum(values, neagative_share)
        #phaseout_consumer.values[0, mask] = np.maximum(values, neagative_share)

        return phaseout_consumer


    def __get_phaseout_subsidy_tax(
            self,
            scenario_type: str,
            is_pc_phaseout_baseline: bool, # TODO: input Dashboard!$V$76
            energy_sector_reform: EnergySectorReform
            ) -> pd.DataFrame:
        """
        'Phaseout factor - price controls (subsidy)' (v361)
        and
        'Phaseout factor - price controls (tax)' (v361)
        CPAT Excel: baseline 1820,1821, scenario 6810,6811

        Both have the same formulas in CPAT Excel

        Scenario specific
        returns dims <simulation_years[0] - 1, simulation_years[1]>
        """

        neagative_share = 1.0 - energy_sector_reform['ffs_phaseout_share_cons']
        max_value = max(neagative_share, 1.0)

        phaseout_subsidy_tax = pd.DataFrame(
            [np.full(len(self.__columns), max_value, dtype=float)],
            columns=self.__columns
        )

        baseline_condition = (scenario_type  != c.BASELINE) or is_pc_phaseout_baseline
        if not (energy_sector_reform['is_pc_phaseout'] and baseline_condition):
            return phaseout_subsidy_tax

        decrease_each_year = (
            - energy_sector_reform['ffs_phaseout_share_cons']
            / energy_sector_reform['pc_phaseout']
        )

        mask = self.__years > energy_sector_reform['pc_start']
        values = (
            1 + decrease_each_year * (self.__years[mask] - energy_sector_reform['pc_start'])
        )
        # for years > than pc_start:
        #phaseout_subsidy_tax.values[0, mask] = np.maximum(values, neagative_share)
        phaseout_subsidy_tax = phaseout_subsidy_tax.copy()
        phaseout_subsidy_tax.iloc[0, mask] = np.maximum(values, neagative_share)
        return phaseout_subsidy_tax
