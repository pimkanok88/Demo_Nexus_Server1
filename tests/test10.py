import pandas as pd
from pages.10_กรอกรายรับ import filter_income_data_by_project

def test_filter_income_data_by_project():
    data = {
        "รหัสโครงการวิจัย": ["A001", "A002", "A001"],
        "จำนวนเงิน": [1000, 2000, 1500],
    }
    df = pd.DataFrame(data)
    filtered_df = filter_income_data_by_project(df, "A001")

    assert len(filtered_df) == 2
    assert filtered_df["จำนวนเงิน"].sum() == 2500
