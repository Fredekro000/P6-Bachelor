''' GPT Prompt
importer datasæt fra brest (eXcel ark)

find annotations i excel ark:
ID
Position (picture coding)
Days(duration)
Overall score

information i eXcel arket:
information i toppen er ligegyldig for nu
første item starter på række 9 og hvert item fylder 8 linjer + 4 padding
næste items starter på 21, 33, 45
der er 5 pages som korresponderer til en række i matricen (Line1, Line2, ...)
hver page har 4 items(6 observationer per item(vi er kun interesset i den første)):
ID står i det første felt(fx. A9) (kolonne A)
Position kan findes ved at finde ID i listen for picture coding (står nedenunder)
dagene for observationer er: 36, 62, 90, 122, 155, 184 (første observation er i kolonne C)
overall score er i C16:H16 (igen bare kig på kolonnne C)

picture coding(liste):
kolonner er nummereret 1 - 4
rækker er nummereret A - E
[[BB2, BN3, BP1, BD1], [BA3, BC2, BA1, BE1], [BC3, BE2, BN2, BB1], [BD2, BP2, BB3, BC1], [BA2, BE3, BD3, BN1]]

lav en string af annotations order (eksempel nedenunder):
'BB2 - A1 - 36 - 0' ('{ID} - {Position} - {Days} - {Overall score}')

put alle items i en liste
print listen til sidst

'''
import pandas as pd

# Indlæs Excel-filen
file_path = "CEPE 2012 inspection AF Efficacy Atlantic.xlsx"
xls = pd.ExcelFile(file_path)

# Picture coding liste
picture_coding = [
    ["BB2", "BN3", "BP1", "BD1"],
    ["BA3", "BC2", "BA1", "BE1"],
    ["BC3", "BE2", "BN2", "BB1"],
    ["BD2", "BP2", "BB3", "BC1"],
    ["BA2", "BE3", "BD3", "BN1"]
]

# Kortlæg ID til position i Picture Coding
position_map = {id_: f"{chr(65 + r)}{c+1}" for r, row in enumerate(picture_coding) for c, id_ in enumerate(row)}

# Definer observationsdage og rækkeoffsets
observation_days = [36, 62, 90, 122, 155, 184]
start_row = 9 - 2  # Første item starter her
item_spacing = 12  # 8 linjer + 4 padding

# Liste til resultater
annotations = []

# Gennemgå alle 5 ark
for sheet in ["Line1", "Line2", "Line3", "Line4", "Line5"]:
    df = xls.parse(sheet)

    for i in range(4):  # 4 items per ark
        row = start_row + (item_spacing * i)
        ID = str(df.iat[row, 0]).strip()  # ID fra kolonne A
        Position = position_map.get(ID, "Unknown")  # Find position i picture coding

        # Første observationsdag (36) i kolonne C
        Days = observation_days[0]
        Overall_score = df.iloc[row + 7, 2]  # Række 16 i kolonne C

        # Formatér annotation
        annotation = f"{ID} - {Position} - {Days} - {Overall_score}"
        annotations.append(annotation)

# Udskriv resultatet
print(annotations)

#christians file moving :)
