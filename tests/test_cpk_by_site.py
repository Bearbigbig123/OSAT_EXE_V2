import pandas as pd

from spc_cpk_dashboard import calculate_cpk, get_site_columns, prepare_cpk_data


def test_site_columns_are_sorted_numerically():
    frame = pd.DataFrame(columns=['point_val', 'site10', 'site2', 'site1', 'site_3'])
    assert get_site_columns(frame) == ['site1', 'site2', 'site10']


def test_by_site_pools_all_site_measurements():
    frame = pd.DataFrame({
        'point_time': ['2026-01-01', '2026-01-02'],
        'point_val': [100.0, 101.0],
        'site1': [9.8, 9.9],
        'site2': [10.1, 10.2],
        'site3': [10.0, None],
    })

    normalized = prepare_cpk_data(frame, 'by_site')

    assert len(normalized) == 5
    assert set(normalized['site']) == {'site1', 'site2', 'site3'}
    assert normalized['point_val'].tolist() == [9.8, 9.9, 10.1, 10.2, 10.0]


def test_point_value_mode_does_not_use_site_values():
    frame = pd.DataFrame({
        'point_time': ['2026-01-01', '2026-01-02'],
        'point_val': [10.0, 10.2],
        'site1': [1.0, 100.0],
    })

    normalized = prepare_cpk_data(frame, 'point_val')

    assert normalized['point_val'].tolist() == [10.0, 10.2]
    assert normalized['site'].tolist() == ['point_val', 'point_val']


def test_cpk_uses_pooled_site_values():
    frame = pd.DataFrame({
        'point_time': ['2026-01-01', '2026-01-02'],
        'point_val': [50.0, 60.0],
        'site1': [9.8, 9.9],
        'site2': [10.1, 10.2],
    })
    chart_info = {'Characteristics': 'Nominal', 'USL': 11.0, 'LSL': 9.0}

    result = calculate_cpk(prepare_cpk_data(frame, 'by_site'), chart_info)

    assert result['Cpk'] == 1.826
