import docx
import pandas as pd
import sys

df1 = pd.read_excel("data/results.xlsx")

persons_raw_data= [] 
for index, row in df1.iterrows(): 
    person_raw_data = [row.Name, row.A, row.C, row.E, row.F, row.G, 
    row.H, row.I, row.L, row.M, row.O, row.Q1, row.Q3, 
    row.Q4, row.G1, row.G2, row.G3, row.G4, row.G5]
    persons_raw_data.append(person_raw_data)

df2 = pd.read_excel("data/report_data.xlsx")

report_data= [] 
for index, row in df2.iterrows(): 
    raw_data = [row.A, row.C, row.E, row.F, row.G, 
    row.H, row.I, row.L, row.M, row.O, row.Q1, row.Q3, 
    row.Q4, row.G1, row.G2, row.G3, row.G4, row.G5]
    report_data.append(raw_data)


for person in persons_raw_data:
    doc = docx.Document()
    for index, item in enumerate(person):
        if person.index(item)==0:
            name = item
            heading = f"{name} DBE Kişilik Envanteri Raporu".format(name)
            doc.add_heading(heading, 0)
            doc_uri = f"reports/{name}.docx".format(name)
        else:
            factor_name = report_data[0][index-1]
            factor_title = f"{factor_name} ({item})".format(factor_name, item)
            header = doc.add_paragraph()
            header.add_run((factor_title)).bold = True
            if item < 36:
                doc.add_paragraph(report_data[1][index-1])
            if item > 35 and item < 66:
                doc.add_paragraph(report_data[2][index-1])
            if item > 65:
                doc.add_paragraph(report_data[3][index-1])
    doc.save(doc_uri)