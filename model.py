# Code by Sophie Tyrrell, Alastair Dobie, and Grace D'Abdank Kunicki
# Developed at HackSheffield 8 on november 11-12  2023
# Emails: jtyrrell1@sheffield.ac.uk, alastairdobie@gmail.com, gkunicki1@sheffield.ac.uk

import openpyxl

wb = openpyxl.load_workbook("Copy of Example Dispatch Model.xlsx", data_only=True)
ws = wb["MODEL TO BUILD"]


class Model():
    OFFSET = 21
    def __init__(self,file, windFactor, solarFactor):
        self.windFactor = windFactor
        self.solarFactor = solarFactor*1000/(2*32.4)
        self.time = 0
        self.cost = 0
        self.file = file
        self.demand = 0
        self.DC2 = {"Wind": 0,
                    "Solar": 0,
                     "Hydro": 0}
        self.DC4 = {"Wood": 0,
                    "Gas": 0,
                    "Diesel": 0,
                     "Coal": 0}
        self.runningTotals = {"Geo": 0,
                              "Wind": 0,
                              "Solar": 0,
                              "Hydro": 0,
                              "Hydro Discharge": 0,
                              "Wood": 0,
                              "Gas": 0,
                              "Diesel": 0,
                              "Coal": 0,
                              "Demand": 0,
                              "CO2 Sim": 0}
        
        self.costs = {"Geo": 70,
                              "Wind": 60,
                              "Solar": 40,
                              "Hydro": 50,
                              "Wood": 150,
                              "Gas": 100,
                              "Diesel": 150,
                              "Coal": 200}
        
        self.co2Costs = {"Geo": 38,
                              "Wind": 9,
                              "Solar": 42,
                              "Hydro": 10,
                              "Hydro Discharge": 10,
                              "Wood": 14,
                              "Gas": 443,
                              "Diesel": 778,
                              "Coal": 960}
        self.storage = 0 #hydro
        self.maxStorage = 20000
        self.carbonIntensity = 0

    def incrementTime(self):
        self.time += 1

    def getTime(self):
        return self.time
    
    def getMaxTime(self):
        return 17520
    
    def getInitialDemand(self): #Initial demand at timestep
        self.demand = self.file["L"+str(self.time+Model.OFFSET)].internal_value
        self.runningTotals["Demand"] += self.demand
        return self.demand
    
    def updateDemand(self):
        pass

    def calculateSimulatedCO2(self, initialDemand):
        self.runningTotals["CO2 Sim"] += self.carbonIntensity * initialDemand 

    def getDC1(self):
        return ("Geo" , self.file["I"+str(self.time+Model.OFFSET)].internal_value)
    
    def setDC2(self):
        self.DC2["Hydro"] = self.file["E"+str(self.time+Model.OFFSET)].internal_value
        self.DC2["Solar"] = self.file["P"+str(self.time+Model.OFFSET)].internal_value * self.solarFactor#multiply by solar factor
        self.DC2["Wind"] = self.file["F"+str(self.time+Model.OFFSET)].internal_value  * self.windFactor #multiply by wind factor

    def setDC4(self):
        self.DC4["Wood"] = self.file["J"+str(self.time+Model.OFFSET)].internal_value
        self.DC4["Gas"] = self.file["G"+str(self.time+Model.OFFSET)].internal_value
        self.DC4["Coal"] = self.file["H"+str(self.time+Model.OFFSET)].internal_value
        self.DC4["Diesel"] = self.file["K"+str(self.time+Model.OFFSET)].internal_value

    def dispatch(self, source):
        powerOutput, overGeneration = self.checkOvergenerated(source[1])
        self.demand -= powerOutput
        self.carbonIntensity += powerOutput * self.co2Costs[source[0]]
        self.runningTotals[source[0]] += powerOutput
        if self.storage + overGeneration > self.maxStorage:
            self.curtail((self.storage + overGeneration - self.maxStorage), source[0])
            self.storage = self.maxStorage
        else:
            self.storage += overGeneration

    def dispatchDC3(self):
        if self.storage > 0:
            powerOutput = min(self.storage,self.demand)
            self.demand -= powerOutput
            self.storage -= powerOutput
            self.runningTotals["Hydro Discharge"] += powerOutput

    def checkOvergenerated(self,powerGenerated):
        powerOutput = min(powerGenerated,self.demand)
        overGeneration = powerGenerated - powerOutput
        return (powerOutput, overGeneration)

    #check for full storage

    def isDemandMet(self):
        return self.demand == 0
    
    def curtail(self, amount,sourceName):
        self.cost += (amount)*self.costs[sourceName]*2
    
    def step(self):
        self.incrementTime()
        self.getInitialDemand()
        initialDemand = self.demand
        if initialDemand > 0:
            self.carbonIntensity = 0
        self.dispatch(self.getDC1())
        if not self.isDemandMet():
            self.setDC2()
            for source in self.DC2.items():
                if not self.isDemandMet():
                    self.dispatch(source)
            if not self.isDemandMet():
                self.dispatchDC3()
                if not self.isDemandMet():
                    self.setDC4()
                    for source in self.DC4.items():
                        if not self.isDemandMet():
                            self.dispatch(source)
        if initialDemand > 0:
            self.carbonIntensity = self.carbonIntensity/initialDemand
            #print("carbon Intensity", self.carbonIntensity)
            self.calculateSimulatedCO2(initialDemand)
        else:
            self.runningTotals["CO2 Sim"] += self.carbonIntensity
                        
                    




m = Model(ws, 1.3, 9)
while m.getTime() < m.getMaxTime():
    m.step()

print(m.runningTotals.items())

for key in m.costs.keys():
    m.cost += (m.runningTotals[key]) *m.costs[key]*2

print("\n\n\n\n")
print("Cost", m.cost)
print("Average Cost", m.cost/(m.runningTotals["Demand"]))
print("Simulated CO2", m.runningTotals["CO2 Sim"]/m.runningTotals["Demand"])
print("Simulated fossil fuel power", ((m.runningTotals["Gas"]+m.runningTotals["Coal"]+m.runningTotals["Diesel"])/m.runningTotals["Demand"])*100, "%")
print("Demand Total",  m.runningTotals["Demand"]/2)
print("Estimated Demand Total", (sum(m.runningTotals.values()) -m.runningTotals["CO2 Sim"]- m.runningTotals["Demand"])/2)
print("Diff", - sum(m.runningTotals.values()) + m.runningTotals["CO2 Sim"] + (2*m.runningTotals["Demand"]))
print("Storage", m.storage/2)
