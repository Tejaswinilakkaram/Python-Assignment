import openpyxl
import pandas as pd

xls = pd.ExcelFile(r"C:\Users\LENOVO\OneDrive\Desktop\Agile Recrute LLP\Python Assignment Input.xlsx")
df1 = pd.read_excel(xls, sheet_name="(Input) User IDs")
df2 = pd.read_excel(xls, sheet_name="(Input) Rigorbuilder RAW")
df3 = pd.merge(df1, df2, left_on='Name',right_on='name')
df3=df3.drop(['S No_x','S No_y','name','uid','Unnamed: 5','Statement 1','Statement 2','Statement 3','Statement 4','Statement 5','Statement 6','Statement 7','Statement 8','Statement 9','Statement 10','Statement 11','Statement 12','Statement 13','Statement 14','Statement 15','Statement 16'],axis=1)
df4=df3.groupby('Team Name')['total_statements'].mean().reset_index()
df5=df3.groupby('Team Name')['total_reasons'].mean().reset_index()
df6 = pd.merge(df5, df4, left_on='Team Name',right_on='Team Name')
df6=df6.sort_values(['total_statements','total_reasons'],ascending=False)
df6.index=[x for x in range(1,len(df6.values)+1)]
df6.index.name='Team Rank'
df6.rename(columns={"Team Name_x":"Thinking Teams Leaderboard"},inplace=True)
df6.to_excel('output1.xlsx')
print(df6)



df3=df3.sort_values(['total_statements','total_reasons'],ascending=False)
df3.index=[x for x in range(1,len(df2.values)+1)]
df3.index.name='Rank'
df3.rename(columns={"Use ID":"UID","total_statements":"No.of Statements","total_reasons":"No.of Reasons"},inplace=True)
df3.to_excel('output2.xlsx')
print(df3)
