# Documentation for the Sisepuede team

## Environment setup
This project uses Conda for environment management. The instructions below show an example of how to set up the environment using Conda and the provided `environment.yml` file.

Create the Conda environment from the project root:
```
conda env create -f environment.yml
```

Activate it:
```
conda activate cpat_sisepuede
```

Install the project requirements:
```
pip install -r requirements.txt
```

To update an existing environment after modifying `environment.yml`:
```
conda env update -f environment.yml --prune
```
To update the pip requirements:
```
pip install -r requirements.txt --upgrade
```


## Project structure
Folders located in the root directory:

```
├── cpat_data
├── cpat_documentation
├── cpat_model
├── cpat_testing
```

Files contained in the root folder are:
- `run_model.py` - entry point for running the model
- `config.py` - used to define scenarios for `run_model.py` and contain `DATA_PATH` for the input data. The keys in `CONFIG` define the scenario names, and the values define the `DashboardInputs` that overwrite the default inputs for each scenario. Also holds `DATA_PATH` for the input data.
- `requirements.txt` - listing all Python dependencies, including `pytest` for tests.

Data input files are not placed in the github repository. All of them are stored in a shared folder named `Mitigation Project Data Files` <br />
`README.txt` file, located in `Mitigation Project Data Files`, contains instructions on where the files should be added within the project.

## Data structure
Data handled in the model uses the `pd.DataFrame` type (or `pd.Sereis`), typically indexed with 3 main index levels defined as `ID_COL_NAMES = ['CountryCode', 'SectorCode', 'FuelCode']`.
Not all data require all three index levels; when a level is redundant or not optimal for calculations.
For time series data frames - columns representing years are stored as strings, e.g. `'2025'`. 

Index and columns of the data are marked at the bottom of docstrings each dosctring as "(index), columns", eg `(c, s), 'Scenario', <simulation_years[0] - 1, simulation_years[1]>`, meaning it has `MultiIndex` with 2 index levels `'CountryCode'` and `'SectorCode'`, 1 column named `Scenario` and rest of the columns representing time series for years in range of `<simulation_years[0] - 1, simulation_years[1]>`. In index description `c` relates to `CountryCode`, `s` to `SectorCode` and `f` to `FuelCode`.

## cpat_model
This folder contains all the logic for the CPAT algorithm. Code for different sections of the model is organized in separate subfolders:
- `components` - contains individual building blocks of the model. Further divided into subfolders based on CPAT Excel sections
- `inputs` - includes logic for data inputs and dashboard inputs (originally based on Dashboard section in CPAT Excel)
- `mappings`
- `constants.py` - located directly inside `cpat_model` folder

### Inputs
Right now there are 2 types of inputs:
- dashboard inputs with logic in `cpat_model/inputs/dashboard_inputs.py`. `DashboardInputs` have their default values defined in constructor. For each scenario they can be overwritten based on `config.py` file values.
- data inputs with logic in `cpat_model/inputs/input_data.py`, they are loaded once and are filtered based on `simulation_years` and `selected_countries` parameters. Any further filtering/modyfiyng of the input data that requires dashboard inputs is usually performed later on in the algorithm. Some of the data inputs are preprocessed with scripts located in `cpat_processing/mitigation` folder.

Constant `SCENARIOS`, located in `config.py` file, defines scenarios for which `run_model` will be run. Keys define name of the scenario that can be used later on for results identification. For each scenario `{...}` holds parameters that will overwrite default parameters defined in `DashboardInputs` constructor.

### Mappings
All the mappings are included in `cpat_model/mappings/mappings.xlsx`, with different mappings stored in separate sheets. These are then loaded and handled in `cpat_model/mappings/mapping.py`.

### Constants
For now most of the constants are located in and are loaded in other files as `import cpat_model.constants as c`. They will be divided later on, based on their purpose.


## Input data folder structure and data reading
All the data used in the model resides in `cpat_data/` folder that has the following structure:
```
├── cpat_data
│   ├── new_data
│   ├── new_data_pkl
```

The target format for data files is `.pkl.bz2`. Currently, many data input files are still stored as `.csv` in `cpat_data/new_data/`, but these are temporary and will be migrated to  `.pkl.bz2` format.
Data loading is handled in `cpat_model/inputs/input_data.py`. All functions that are used to load the data should have `[Data Loading]` tag in their docstring.


## Running the model
The only way to run the mitigation module now is by running `python run_model.py`.
Dashboard inputs (manual inputs) can be adjusted in `config.py` file by modifying `SCENARIOS`. Each key, value pair in `SCENARIOS` dictionary represents name of a scenario and parameters that overwrite the default dashboard input values for this specific scenario.

Parameters that can be adjusted from the code are: `selected_countries` and `last_simulation_year`. Please do not change `base_year = 2022`.

Saving the results is not handled by `run_model.py` and can be easily implemented in a custom way by identifying the scenario via the keys in `config.SCENARIOS` and saving the desired variables or tables to CSV/XLSX files. Please see `Example on how to save output variables` comment in the `run_model.py` file.

## Testing
The `cpat_testing` directory contains unit tests for the project logic and mirrors the project structure. The `cpat_testing/model/components` folder includes unit tests for functions defined in `cpat_model/components`. Subdirectories within `cpat_testing/model/components` follow the same structure as `cpat_model/components` .

Some unit tests are not included in this repository because they require country-specific input data that is not distributed here.

Originally, several tests in `cpat_testing/model/components` relied on expected DataFrames stored as `.csv` files in `cpat_model/model/data_expected`. The path to this directory is defined in `conftest.py` under the constant `EXPECTED_DATA_PATH`. These tests have been removed from this repository, but the directory structure and configuration remain available. Therefore, If you need to implement similar tests that compare results against predefined DataFrames, you can store the expected `.csv` files in this directory and reuse the existing path configuration.

`pytest`:
Please name test files `test_{tested_file}.py` and try to define test functions names so they indicate what they check, eg: `def test_{tested_function}(...):` or `def test_{tested_functionality}(...):`.
Please mind that by default, `pytest` looks for files that match the pattern `test_*.py` or `*_test.py`, as well as test functions that start with `test_`.

`../conftest.py` holds configuration for all `cpat_testing` dir. Used to ensure that the root directory is always added to the Python path whenever pytest is ran. We could also hold here fixtures used accross multiple files.

For running all tests:
```
pytest cpat_testing
```
For running tests only in a given dir, eg in `model`:
```
pytest cpat_testing/model
```
For running just 1 test:
```
pytest -k "test_my_func"
```
Adding `-s` (`pytest -s -k "test_my_func"`) disables output capturing, allowing print statements to appear in the terminal. (for test debugging)

IMPORTANT: if `Access is denied.` occurs, please use `python -m pytest`, e.g.:
```
python -m pytest cpat_testing
```

## Power model - example of a component structure
The Power model logic is located in `cpat_model/components/power` folder. `../power.py` file holds the main class `Power` with instances of Power model components as attributes.
`__init__` creates instances of all components and calculates all values that can be determined before iterating over `simulation_years`. `calculate_power_year` updates power attributes year by year in a loop over `simulation_years` located in `run_model.py`
Please see docstrings and comments in `../power.py`, `../data.py`, `../variable_cost.py`, `../demand.py`, `../investment.py` and `../generation.py` for more information about the Power components.

## Notes
Please try to use type hinting. <br />
Please try to use docstrings. <br />
Please use feature branches and create Pull Requests. <br />
Please try to follow coding standards: https://peps.python.org/pep-0008/ <br />
Using a linter is recommended (Pylint)
