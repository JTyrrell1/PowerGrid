# Code by Sophie Tyrrell, Alastair Dobie, and Grace D'Abdank Kunicki
# Developed at HackSheffield 8 on november 11-12  2023
# Emails: jtyrrell1@sheffield.ac.uk, alastairdobie@gmail.com, gkunicki1@sheffield.ac.uk

import scipy
from scipy.optimize import linprog
import csv

file = open(r"C:\Users\JTyrr\hackathon\Power stations\data.csv")
csvreader = csv.reader(file)
rows = []
for row in csvreader:
    rows.append(row)


file2 = open(r"C:\Users\JTyrr\hackathon\Power stations\solar.csv")
solarreader = csv.reader(file2)
solar = []
for row in solarreader:
    solar.append(row)

output = open(r"C:\Users\JTyrr\hackathon\Power stations\output.txt", "a")

MaxCost = 0
TotalCost = 0
Battery = 0
i = 0
while (i < 17520):
    Other = rows[i][8]
    MaxSolar = float(solar[i][0])

    obj = [60,40,0]

    lhs_eq = [[1,1,1]];
    rhs_eq = [Other]

    bnd = [(0, 153),(0, float(MaxSolar)),(0, Battery)]

    opt = linprog(c=obj,A_eq=lhs_eq, b_eq=rhs_eq, bounds=bnd, method="revised simplex")

    Cost = (153 * 60) + (float(MaxSolar) * 40)

    if Cost > MaxCost:
        MaxCost = Cost
        HighestCombo = opt.x
    Battery = Battery - float(opt.x[2])
    TotalCost = TotalCost + Cost

    Energy = (153 + MaxSolar)
    if (Energy + float(opt.x[2])) > float(Other):
        Extra = (Energy + float(opt.x[2])) - float(Other)
        if Battery + Extra < 6755:
            Battery = Extra + Battery
        else:
            Battery = 6755

    print(opt.x, " Cost: ", Cost, " Charge: ", Battery, " Other: ", Other)
    print()
    output.write(str(opt.x))
    output.write("\n")
    i = i + 1

    if Battery < 0:
        break

    if float(Other) != (float(opt.x[0]) + float(opt.x[1]) + float(opt.x[2])):
        break

AverageCost = TotalCost / 17521
print("Max cost is: ", MaxCost, " with ", HighestCombo)
print("Average cost is: ", AverageCost)
