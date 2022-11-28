import pandas as pd
from itertools import combinations


xls = pd.read_excel('Book1.xlsx')
xls = xls.sort_values(by=['name'])
dict1 = dict(zip(xls.name, xls.width))
target = 10

def closest_subset(s, items):
    return dict(sorted(min((
        c
        for i in range(1, len(items) + 1)
        for c in combinations(items, i)
    ), key=lambda x: abs(s - sum([v2 for v1, v2 in x])))))
    
solution = closest_subset(target, dict1.items())
df = pd.DataFrame(data=solution, index=[0])

with pd.ExcelWriter("kapidata.xlsx", engine="openpyxl", mode="a", if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name="name")

print(f'{solution} with a sum of {round(sum(solution.values()),2)} cm, at a count of {len(solution.values())} books')





