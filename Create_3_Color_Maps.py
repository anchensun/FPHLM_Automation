import pandas as pd

#Import A1_output.csv from your_run_name/pr_unmatched/env/Results/ and
#select "first row contains field names", without primary key.
a1_output_path = '/a/mitch.cs.fiu.edu./disk/mitch-b/dmis-research/Anchen/2019/a1/processing/pr_unmatched/env/Results/A1_output.csv'

a1_output = pd.read_csv(a1_output_path)

#update the zipcode and add leading zeroes
a1_output['ZipCode'] = a1_output['ZipCode'].apply(lambda x: '{0:0>5}'.format(x))

#Separate by construction type
#Note: Manufactured can also be Mobile
const_types = ['Frame', 'Masonry', 'Manufactured']
for c in const_types:
    curr = a1_output[a1_output['Cons_type'] == c ]
    res = curr[['ZipCode', 'Structure_Loss_Cost']].copy()
    res['Structure_Loss_Cost'] = res['Structure_Loss_Cost']*1000
    res.to_excel('./A1_ConstType_{}.xls'.format(c), index=False)

