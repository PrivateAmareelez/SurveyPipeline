import pandas as pd

data_IOC = pd.read_csv('IOC.csv')
data_EC = pd.read_csv('EC.csv')
data = data_IOC.iloc[:, 8:232].append(data_EC.iloc[:, 8:232])

data.loc['mean'] = data.mean()

cols = ['ID', 'Mean Score', 'Num. of Non-Feasible', 'Num. of Feasible', "All Marks"]
result = pd.DataFrame(columns=cols)
for i in range(0, len(data.columns), 2):
    problem, feasible = data.columns[i], data.columns[i + 1]
    temp = [[problem, data[problem]['mean'], (data[feasible] == 'No').sum(), (data[feasible] == 'Yes').sum(),
             list(data[problem][:-1])]]
    result = result.append(pd.DataFrame(temp, columns=cols), ignore_index=True)

result.to_html('result.html')
result.to_csv('result.csv')
result.to_excel('result.xlsx')
