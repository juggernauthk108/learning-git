import pandas as pd
import xlsxwriter

excel = pd.read_excel(r'C:\Users\tejas\Desktop\pandas\data.xlsx',engine='openpyxl',header=0)
data = pd.DataFrame(excel)

#list of unique environments
envs = data.Env.unique()
print("Printing envs : {}".format(envs))

count=0

wb = xlsxwriter.Workbook("output.xlsx")
worksheet = wb.add_worksheet()

matrix = []

for e in envs:
   
    violations = data[(data.Env==e)].Violation.unique()
    
    for v in violations:
        lineItem = []
        arn =   data[(data.Env==e) & (data.Violation==v)].ARN
        
        summary = "infra | tr | 234273 | AWS cloud"
        DID = data[(data.Env==e)].DID.unique()[0]
        criticality = "HIGH"
        resource = ""
        for a in arn:
            resource += "* " + a + "\n"

        description = "AWS DID: {}\nCriticality: {}\nResources Affected: \n{}".format(DID,criticality,resource)
        print("#"*30)
        count = count + 1

        lineItem.append(summary)
        lineItem.append(DID)
        lineItem.append(criticality)
        lineItem.append(description)

        matrix.append(lineItem)

print("Total ticket need to be raised : {}".format(count))

df = pd.DataFrame(matrix)
writer = pd.ExcelWriter('test.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='welcome', index=False)
writer.save()
