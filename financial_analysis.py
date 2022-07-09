import pandas as pd
from docxtpl import DocxTemplate

asset = pd.read_csv('balance.csv',index_col=0)
profit = pd.read_csv('profit.csv',index_col=0)

#定位总资产和总负债
total_asset = asset.loc['资产总计'].values
total_debt= asset.loc["负债合计"].values

#计算流动比率
#liq_ratio = asset.loc["流动资产合计"].values/asset.loc["流动负债合计"].values
#计算速动比率
#qui_ratio = (asset.loc["流动资产合计"].values - asset.loc["存货"].values) / asset.loc["流动负债合计"].values
#计算资产负债率
debt_asset = asset.loc["负债合计"].values * 100/asset.loc["资产总计"].values
#计算全部债务
#debt = asset.loc["长期借款"].values + asset.loc["应付债券"].values + asset.loc["交易性金融负债"].values + asset.loc["应付票据"].values + asset.loc["一年内到期的非流动负债"].values
#计算全部债务资本比率
#debt_capital = debt * 100 / (debt + asset.loc["所有者权益合计"].values)
#计算毛利率
income = profit.loc["营业总收入"].values
margin = profit.loc["利润总额"].values
margin_income = margin / income
cost = income - margin

#将计算结果写入dataframe
#Tableindex=["总资产","总负债","流动比率","速动比率","资产负债率%","债务资本比率%","营业收入","营业成本","毛利率%"]
#data = pd.DataFrame([total_asset, total_debt, liq_ratio, qui_ratio, debt_asset, debt_capital, income, cost, margin],
#                            index=Tableindex)

Tableindex=["总资产","总负债","资产负债率%","营业总收入","营业总成本","毛利率%"]
data = pd.DataFrame([total_asset, total_debt, debt_asset, income, cost, margin_income],
                           index=Tableindex)
data = round(data,2)
data.to_excel('data.xlsx',header=['20220331','20211231','20201231','20191231'],index_label=["项目"])


