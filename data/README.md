# Data Source & Schema

## Dataset
- **Source:** [DataCo Smart Supply Chain for Big Data Analysis (Kaggle)](https://www.kaggle.com/datasets/shashwatwork/dataco-smart-supply-chain-for-big-data-analysis)
- **File:** `DataCoSupplyChainDataset.csv` (180,519 rows x 53 columns; ~95 MB, `latin1` encoded)
- **Time Horizon:** January 2015 – January 2018 (37 months; full-year FY2015–FY2017 used for annualized P&L baseline)

## How to Run Locally
1. Download `DataCoSupplyChainDataset.csv` from Kaggle and place it inside this `data/` folder.
2. Run the automated pipeline from the repository root:
   ```bash
   pip install -r requirements.txt
   python src/01_data_cleaning_and_diagnostics.py
   python src/02_build_ebitda_model.py
   ```
3. All aggregated CSVs, executive PNG charts, and the dynamic 4-tab Excel model (`deliverables/EBITDA_Transformation_Model.xlsx`) will regenerate automatically.
