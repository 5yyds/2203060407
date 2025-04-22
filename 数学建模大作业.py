import numpy as np
from scipy.stats import norm
import pandas as pd

# 数据
a = 104
L1 = 43248
b = 12
L2 = 10673

# 计算相对危险度
RR_hat = (a / L1) / (b / L2)

# 计算 Ka
Ka = ((a * (L1 + L2) - (a + b) * L1) ** 2) / ((a + b) * L1 * L2)

# 计算 95% 置信区间
z = norm.ppf(0.975)
lower_bound = RR_hat ** (1 - z / np.sqrt(Ka))
upper_bound = RR_hat ** (1 + z / np.sqrt(Ka))

# 将结果保存到 DataFrame
result_df = pd.DataFrame({
    '指标': ['相对危险度 (RR) 估计值', '总体相对危险度的 95% 可信区间下限', '总体相对危险度的 95% 可信区间上限'],
    '值': [RR_hat, lower_bound, upper_bound]
})

# 将结果保存到 Excel 文件
result_df.to_excel('relative_risk_result.xlsx', index=False)

print("结果已保存到 relative_risk_result.xlsx 文件中。")
   