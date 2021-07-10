import pandas as pd
import math

excel_file = '../python/orders.xlsx'
orders = pd.read_excel(excel_file)

df = pd.DataFrame(orders, columns=['Item', 'Cost', 'Qty'])
array = df.to_numpy()

newSheet = [['Item', 'Cost', 'Qty', 'Total Cost']]

for i in range(len(array[:])):

    # Box of items
    if not math.isnan(array[i][2]) and 'BOX' in array[i][0] and not '-S-' in array[i][0]:
        for j in range(5):
            cost = 11
            if j == 0 or j == 4:
                qty = 1 * array[i][2]
            else:
                qty = 2 * array[i][2]
            newSheet.append([array[i+j+1][0], cost, qty, cost*qty ])

    # Sunglasses
    if not math.isnan(array[i][2]) and '-S-' in array[i][0]:
        cost = 14
        qty = 8 * array[i][2]
        item = array[i][0]
        newSheet.append([item.split('-BOX OF')[0], cost, qty, cost*qty])
    
    #Individual items
    if not math.isnan(array[i][2]) and not 'BOX' in array[i][0]:
        cost = array[i][1]
        qty = array[i][2]
        newSheet.append([array[i][0], cost, qty, cost*qty])


# for row in newSheet:
#     print(row)

df2 = pd.DataFrame(newSheet)
writer = pd.ExcelWriter('test.xlsx')
df2.to_excel(writer, sheet_name='po', header=False, index=False)
writer.save()