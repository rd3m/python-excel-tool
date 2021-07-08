import pandas as pd

excel_file = '../python/orders.xlsx'
orders = pd.read_excel(excel_file)

df = pd.DataFrame(orders, columns=['Item', 'Cost', 'Qty'])
array = df.to_numpy()

for row in array:
    if row[1] == 96:
        print(row)

# print(array[0][1])