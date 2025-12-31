import numpy as np


def get_iqr_outlier(data, multiplier: float = 1.5):
    if not data or len(data) < 4:
        return []
    q1 = np.percentile(data, 25)
    q2 = np.percentile(data, 75)
    iqr = q2 - q1
    lower_bound = q1 - multiplier * iqr
    upper_bound = q2 + multiplier * iqr

    outlier = [x for x in data if x < lower_bound or x > upper_bound]

    return outlier