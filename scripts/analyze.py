import pandas as pd
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
matplotlib.rcParams['axes.unicode_minus'] = False ## 마이나스 '-' 표시 제대로 출력
 
from sklearn.linear_model import LinearRegression
import statsmodels.api as sm

file_path = 'data/output/특정품목 조달 내역_2022.xlsx'
data = pd.read_excel(file_path, index_col=0)

# make new columns
data['용량'] = data[['냉방용량','난방용량']].max(axis=1)

# delete the row with the 0 cost
data = data[data['계약금액'] != 0]
data = data[data['장비금액'] != 0]
data = data[data['용량'] != 0]

# make new columns
data['장비금액비율'] = data['장비금액']/data['계약금액']

# calculate the lower and upper bounds based on the specified quantile value
q1 = data['용량'].quantile(0.25)
q3 = data['용량'].quantile(0.75)
iqr = q3 - q1
lower_bound = q1 - (1.5 * iqr)
upper_bound = q3 + (1.5 * iqr)

# filter out the outlier rows
data = data[(data['용량'] >= lower_bound) & (data['용량'] <= upper_bound)]

# create a linear regression model
reg = LinearRegression().fit(data[['용량']], data['계약금액'])

# use statsmodels to identify influential points
influence = sm.OLS(data['계약금액'], sm.add_constant(data['용량'])).fit().get_influence()
outliers = influence.summary_frame().loc[influence.summary_frame().cooks_d > 4/len(data)]

# remove the influential points from the dataframe
data = data.drop(outliers.index)

# re-fit the linear regression model without the influential points
reg = LinearRegression().fit(data[['용량']], data['계약금액'])

# print the coefficients
print('Intercept:', reg.intercept_)
print('Coefficient:', reg.coef_[0])

# plot the scatter plot and linear regression line
fig, ax = plt.subplots()
ax.scatter(data['용량'], data['계약금액'], color='blue')
ax.plot(data['용량'], reg.predict(data[['용량']]), color='red')
ax.set_xlabel('Capacity(kW)')
ax.set_ylabel('Total Cost(KRW)')
ax.set_title('Scatter plot with linear regression line')
plt.show()