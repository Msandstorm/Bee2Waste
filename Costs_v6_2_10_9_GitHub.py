# -*- coding: utf-8 -*-
"""
Created on Tue Feb  2 18:17:29 2021

M A I N   A U T H O R: Mexitli Sandoval-Reyes

C O N T R I B U T I O N S:
Mexitli Sandoval-Reyes - Code writing, conceptualization, and methodology
Rui Semeano - MRF code & overall peer review
S. Carvalho - Landfill model conceptualization
Rui Semeano, P.Ferrão, A.Braga, S. Carvalho, A. Lorena - Conceptualization of the Waste Collection module
A.Braga, S. Carvalho, A. Lorena - Data collection
P.Ferrão, Rui He - Overall peer review
"""

#%%

"""
##  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
##  %%%%                                           %%%%
##  %%%%              L I B R A R Y                %%%%
##  %%%%                                           %%%%
##  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
"""

## GENERAL
def SpaceOrDashToUnderscore(list): # Within a list of values, convert spaces to Underscores  
    for i in range(len(list)):
        list[i] = list[i].replace(" ","_")
        list[i] = list[i].replace("-","_")
    return list


def ReadValueFromExcel(filename, sheet, column, row):
    """Read a single cell value from an Excel file"""
    return pd.read_excel(filename, sheet_name=sheet, skiprows=row - 1, usecols=column, nrows=1, header=None, names=["Value"]).iloc[0]["Value"]


"""
############     T E C H N O L O G Y   R E L A T E D     ###############
"""


"""
#######   Landfill   #######
"""


def KeepLastYearOnly(list):
    from itertools import dropwhile
    for i in range(len(CollectionYears)):
        list[i] = list[i].replace(" years","")
        list[i] = ''.join(dropwhile(lambda x: x not in '-', list[i]))
        list[i] = list[i].replace("-","")
    return list



# It calculates the Landfill emissions for flaring and ICE (collected and Uncollected), based on 1 ton of (wet) individual waste fraction of MSW
def LandGEM_TonEmissionsPerTonWaste(df_LandfillGasGeneration, WasteType, df_LandfillGasCollectionEff, df_General_Factors):
    
    # Getting the input data
    k = df_LandfillGasGeneration.at["k","Value"]
    MCF = df_LandfillGasGeneration.at["MCF","Value"]
    DOC = []
    for i in range(len(WasteType)):
        DOC.append(df_LandfillGasGeneration.at["DOC_" + WasteType[i],"Value"])
    DOC_ss = df_LandfillGasGeneration.at["DOC_ss","Value"]
    DOC_F = df_LandfillGasGeneration.at["DOC_F","Value"]
    F = df_LandfillGasGeneration.at["F","Value"]
    pCH4_Landfill = df_LandfillGasGeneration.at["pCH4_Landfill","Value"]
    d_CH4 = df_General_Factors.at["d_CH4","Value"]
    NMOC_conc = df_LandfillGasGeneration.at["NMOC_conc","Value"]
    Lifetime_Emissions = df_LandfillGasGeneration.at["LT_e","Value"]
    M_CH4 = df_General_Factors.at["M_CH4","Value"]
    M_CO2 = df_General_Factors.at["M_CO2","Value"]
    M_NMOC = df_General_Factors.at["M_NMOC","Value"]
    mT_g = df_General_Factors.at["mT_g","Value"]
    CollectionEff = []
    OxidationRate = []
    for j in range(Lifetime_Emissions):
        if j < int(df_LandfillGasCollectionEff.at[0, "CollectionYears"]):
            CollectionEff.append(df_LandfillGasCollectionEff.at[0,'Gas collection efficiency [%]'])
            OxidationRate.append(df_LandfillGasCollectionEff.at[0,'Oxidation rate [%]'])
        elif j >= int(df_LandfillGasCollectionEff.at[0, "CollectionYears"]) and j < int(df_LandfillGasCollectionEff.at[1, "CollectionYears"]):
            CollectionEff.append(df_LandfillGasCollectionEff.at[1,'Gas collection efficiency [%]'])
            OxidationRate.append(df_LandfillGasCollectionEff.at[1,'Oxidation rate [%]'])
        elif j >= int(df_LandfillGasCollectionEff.at[1, "CollectionYears"]) and j < int(df_LandfillGasCollectionEff.at[2, "CollectionYears"]):
            CollectionEff.append(df_LandfillGasCollectionEff.at[2,'Gas collection efficiency [%]'])
            OxidationRate.append(df_LandfillGasCollectionEff.at[2,'Oxidation rate [%]'])
        else:
            CollectionEff.append(df_LandfillGasCollectionEff.at[3,'Gas collection efficiency [%]'])
            OxidationRate.append(df_LandfillGasCollectionEff.at[3,'Oxidation rate [%]'])
            
    # Creating output files, for single stream
    L_0_ss = 0
    Q_LandfillEmissions_ss = pd.DataFrame(np.nan, index=["Total","Collected","Uncollected_Oxidized"], columns=["Q_CH4","Q_CO2","Q_NMOC","Q_LFG"])
    Q_CH4_perYear_ss = pd.DataFrame(np.nan, index=list(range(0, 100)),columns=['Total'])
    Q_CO2_perYear_ss = pd.DataFrame(np.nan, index=list(range(0, 100)), columns=['Total'])
    Q_NMOC_perYear_ss = pd.DataFrame(np.nan, index=list(range(0, 100)), columns=['Total'])
    Q_LFG_perYear_ss = pd.DataFrame(np.nan, index=list(range(0, 100)), columns=['Total'])

    # Creating output files, for waste types
    L_0 = []
    Q_LandfillEmissions = pd.DataFrame(np.nan, index=WasteType, columns=["Q_CH4","Q_CO2","Q_NMOC","Q_LFG"])
    Q_CH4_perYear = pd.DataFrame(np.nan, index=list(range(0, 100)), columns=WasteType)
    Q_CO2_perYear = pd.DataFrame(np.nan, index=list(range(0, 100)), columns=WasteType)
    Q_NMOC_perYear = pd.DataFrame(np.nan, index=list(range(0, 100)), columns=WasteType)
    Q_LFG_perYear = pd.DataFrame(np.nan, index=list(range(0, 100)), columns=WasteType)

    Q_LandfillEmissions_Collected = pd.DataFrame(np.nan, index=WasteType, columns=["Q_CH4","Q_CO2","Q_NMOC","Q_LFG"])
    Q_CH4_Collected_perYear = pd.DataFrame(np.nan, index=list(range(0, 100)), columns=WasteType)
    Q_CO2_Collected_perYear = pd.DataFrame(np.nan, index=list(range(0, 100)), columns=WasteType)
    Q_NMOC_Collected_perYear = pd.DataFrame(np.nan, index=list(range(0, 100)), columns=WasteType)

    Q_LandfillEmissions_Uncollected_Oxidized = pd.DataFrame(np.nan, index=WasteType, columns=["Q_CH4","Q_CO2","Q_NMOC","Q_LFG"])
    Q_CH4_Uncollected_perYear = pd.DataFrame(np.nan, index=list(range(0, 100)), columns=WasteType)
    Q_CO2_Uncollected_perYear = pd.DataFrame(np.nan, index=list(range(0, 100)), columns=WasteType)
    Q_NMOC_Uncollected_perYear = pd.DataFrame(np.nan, index=list(range(0, 100)), columns=WasteType)

    
    # Calculating L_0 and transforming from units from ton of CH4/ton of waste to m3 of CH4/ton of waste
      # For Single Stream
    L_0_ss = (MCF*DOC_ss*DOC_F*F*(16.04/12.011)) 
#    L_0_ss = (MCF*DOC_ss*DOC_F*F*(16.04/12.011))/(d_CH4/1000) # For Single Stream

      # For waste types
    for i in range(len(WasteType)):
        L_0.append(MCF*DOC[i]*DOC_F*F*(16.04/12.011))
#        L_0.append((MCF*DOC[i]*DOC_F*F*(16.04/12.011))/(d_CH4/1000))


    # Calculating Emissions (CH4, CO2, NMOC, and Total)
      # For Single Stream
    for j in range(Lifetime_Emissions):
        Q_CH4_dummy2 = []
        for l in np.arange (0,1,0.1):
            Q_CH4_dummy2.append(float(k)*float(L_0_ss)*float(1/10)*float(math.pow((math.e),(-k*(j+l)))))
        Q_CH4_perYear_ss.at[j,"Total"] = sum(Q_CH4_dummy2)
        Q_CO2_perYear_ss.at[j,"Total"] = sum(Q_CH4_dummy2)*((1-pCH4_Landfill)/pCH4_Landfill)*(M_CO2/M_CH4)
        Q_NMOC_perYear_ss.at[j,"Total"] = sum(Q_CH4_dummy2)/pCH4_Landfill*(M_NMOC/M_CH4)*(NMOC_conc/1000000)
        Q_LFG_perYear_ss.at[j,"Total"] = Q_CH4_perYear_ss.at[j,'Total'] + Q_CO2_perYear_ss.at[j,'Total'] + Q_NMOC_perYear_ss.at[j,'Total']

      # For waste types
    for i in range(len(WasteType)): 
#        Q_CH4_dummy = []
        for j in range(Lifetime_Emissions):
            Q_CH4_dummy2 = []
            for l in np.arange (0,1,0.1):
#                Q_CH4_dummy.append(float(k)*float(L_0[i])*float(1/10)*float(math.pow((math.e),(-k*(j+l)))))
                Q_CH4_dummy2.append(float(k)*float(L_0[i])*float(1/10)*float(math.pow((math.e),(-k*(j+l)))))
            Q_CH4_perYear.at[j,WasteType[i]] = sum(Q_CH4_dummy2)
            Q_CO2_perYear.at[j,WasteType[i]] = sum(Q_CH4_dummy2)*((1-pCH4_Landfill)/pCH4_Landfill)*(M_CO2/M_CH4)
            Q_NMOC_perYear.at[j,WasteType[i]] = sum(Q_CH4_dummy2)/pCH4_Landfill*(M_NMOC/M_CH4)*(NMOC_conc/1000000)
            Q_LFG_perYear.at[j,WasteType[i]] = Q_CH4_perYear.at[j,WasteType[i]] + Q_CO2_perYear.at[j,WasteType[i]] + Q_NMOC_perYear.at[j,WasteType[i]]


    # Creating the output file for Total Emissions (CH4, CO2, NMOC, and Total)
      # For single stream
    Q_LandfillEmissions_ss.at["Total","Q_CH4"] = sum(Q_CH4_perYear_ss.loc[:,"Total"])
    Q_LandfillEmissions_ss.at["Total","Q_CO2"] = sum(Q_CO2_perYear_ss.loc[:,"Total"])
    Q_LandfillEmissions_ss.at["Total","Q_NMOC"] = sum(Q_NMOC_perYear_ss.loc[:,"Total"])
    Q_LandfillEmissions_ss.at["Total","Q_LFG"] = sum(Q_LFG_perYear_ss.loc[:,"Total"])

      # For waste types    
    for i in range(len(WasteType)):
        Q_LandfillEmissions.at[WasteType[i],"Q_CH4"] = sum(Q_CH4_perYear.loc[:,WasteType[i]])
        Q_LandfillEmissions.at[WasteType[i],"Q_CO2"] = sum(Q_CO2_perYear.loc[:,WasteType[i]])
        Q_LandfillEmissions.at[WasteType[i],"Q_NMOC"] = sum(Q_NMOC_perYear.loc[:,WasteType[i]])
        Q_LandfillEmissions.at[WasteType[i],"Q_LFG"] = sum(Q_LFG_perYear.loc[:,WasteType[i]])


    # Creating the output file for collected Emissions (CH4, CO2, NMOC, and Total)   
      # For Single Stream
    for j in range(Lifetime_Emissions):
        Q_CH4_perYear_ss.at[j,"Collected"] =  Q_CH4_perYear_ss.at[j,"Total"] * CollectionEff[j]
        Q_CO2_perYear_ss.at[j,"Collected"] =  Q_CO2_perYear_ss.at[j,"Total"] * CollectionEff[j]
        Q_NMOC_perYear_ss.at[j,"Collected"] =  Q_NMOC_perYear_ss.at[j,"Total"] * CollectionEff[j]
    Q_LandfillEmissions_ss.at["Collected","Q_CH4"] = sum(Q_CH4_perYear_ss.loc[:,"Collected"])
    Q_LandfillEmissions_ss.at["Collected","Q_CO2"] = sum(Q_CO2_perYear_ss.loc[:,"Collected"])
    Q_LandfillEmissions_ss.at["Collected","Q_NMOC"] = sum(Q_NMOC_perYear_ss.loc[:,"Collected"])
    Q_LandfillEmissions_ss.at["Collected","Q_LFG"] =  Q_LandfillEmissions_ss.at["Collected","Q_CH4"] + Q_LandfillEmissions_ss.at["Collected","Q_CO2"] + Q_LandfillEmissions_ss.at["Collected","Q_NMOC"]

      # For waste types
    for i in range(len(WasteType)):
        for j in range(Lifetime_Emissions):
            Q_CH4_Collected_perYear.at[j,WasteType[i]] = Q_CH4_perYear.at[j,WasteType[i]] * CollectionEff[j]
            Q_CO2_Collected_perYear.at[j,WasteType[i]] = Q_CO2_perYear.at[j,WasteType[i]] * CollectionEff[j]
            Q_NMOC_Collected_perYear.at[j,WasteType[i]] = Q_NMOC_perYear.at[j,WasteType[i]] * CollectionEff[j]
        Q_LandfillEmissions_Collected.at[WasteType[i],"Q_CH4"] = sum(Q_CH4_Collected_perYear.loc[:,WasteType[i]])
        Q_LandfillEmissions_Collected.at[WasteType[i],"Q_CO2"] = sum(Q_CO2_Collected_perYear.loc[:,WasteType[i]])
        Q_LandfillEmissions_Collected.at[WasteType[i],"Q_NMOC"] = sum(Q_NMOC_Collected_perYear.loc[:,WasteType[i]])
        Q_LandfillEmissions_Collected.at[WasteType[i],"Q_LFG"] =  Q_LandfillEmissions_Collected.at[WasteType[i],"Q_CH4"] + Q_LandfillEmissions_Collected.at[WasteType[i],"Q_CO2"] + Q_LandfillEmissions_Collected.at[WasteType[i],"Q_NMOC"]


    # Creating the output file for UNCOLLECTED Emissions AND OXIDIZED CH4 (CH4, CO2, NMOC, and Total)
      # For Single Stream
    for j in range(Lifetime_Emissions):
        Q_CH4_perYear_ss.at[j,"Uncollected_Oxidized"] =  Q_CH4_perYear_ss.at[j,"Total"] * (1-CollectionEff[j]) * (1-OxidationRate[j])
        Q_CO2_perYear_ss.at[j,"Uncollected_Oxidized"] =  Q_CO2_perYear_ss.at[j,"Total"] * (1-CollectionEff[j]) + Q_CH4_perYear_ss.at[j,"Total"] * (1-CollectionEff[j]) * OxidationRate[j] * (M_CO2/M_CH4)
        Q_NMOC_perYear_ss.at[j,"Uncollected_Oxidized"] =  Q_NMOC_perYear_ss.at[j,"Total"] * (1-CollectionEff[j])
    Q_LandfillEmissions_ss.at["Uncollected_Oxidized","Q_CH4"] = sum(Q_CH4_perYear_ss.loc[:,"Uncollected_Oxidized"])
    Q_LandfillEmissions_ss.at["Uncollected_Oxidized","Q_CO2"] = sum(Q_CO2_perYear_ss.loc[:,"Uncollected_Oxidized"])
    Q_LandfillEmissions_ss.at["Uncollected_Oxidized","Q_NMOC"] = sum(Q_NMOC_perYear_ss.loc[:,"Uncollected_Oxidized"])
    Q_LandfillEmissions_ss.at["Uncollected_Oxidized","Q_LFG"] =  Q_LandfillEmissions_ss.at["Uncollected_Oxidized","Q_CH4"] + Q_LandfillEmissions_ss.at["Uncollected_Oxidized","Q_CO2"] + Q_LandfillEmissions_ss.at["Uncollected_Oxidized","Q_NMOC"]
    
      # For waste types
    for i in range(len(WasteType)):
        for j in range(Lifetime_Emissions):
            Q_CH4_Uncollected_perYear.at[j,WasteType[i]] = Q_CH4_perYear.at[j,WasteType[i]] * (1-CollectionEff[j]) * (1-OxidationRate[j])
            Q_CO2_Uncollected_perYear.at[j,WasteType[i]] = Q_CO2_perYear.at[j,WasteType[i]] * (1-CollectionEff[j]) + Q_CH4_perYear.at[j,WasteType[i]] * (1-CollectionEff[j]) * OxidationRate[j] * (M_CO2/M_CH4)
            Q_NMOC_Uncollected_perYear.at[j,WasteType[i]] = Q_NMOC_perYear.at[j,WasteType[i]] * (1-CollectionEff[j])
        Q_LandfillEmissions_Uncollected_Oxidized.at[WasteType[i],"Q_CH4"] = sum(Q_CH4_Uncollected_perYear.loc[:,WasteType[i]])
        Q_LandfillEmissions_Uncollected_Oxidized.at[WasteType[i],"Q_CO2"] = sum(Q_CO2_Uncollected_perYear.loc[:,WasteType[i]])
        Q_LandfillEmissions_Uncollected_Oxidized.at[WasteType[i],"Q_NMOC"] = sum(Q_NMOC_Uncollected_perYear.loc[:,WasteType[i]])
        Q_LandfillEmissions_Uncollected_Oxidized.at[WasteType[i],"Q_LFG"] =  Q_LandfillEmissions_Uncollected_Oxidized.at[WasteType[i],"Q_CH4"] + Q_LandfillEmissions_Uncollected_Oxidized.at[WasteType[i],"Q_CO2"] + Q_LandfillEmissions_Uncollected_Oxidized.at[WasteType[i],"Q_NMOC"]

               
    return Q_LandfillEmissions, Q_LandfillEmissions_Collected, Q_LandfillEmissions_Uncollected_Oxidized, Q_LandfillEmissions_ss



# It calculates the average leachate generated and collected in Landfill per year [m3/ton of waste]
def AverageLeachateCollected (df_LandfillLeachate, df_LandfillGasGeneration):
    
    AverageLeachateCollected = []
    Lifetime_Emissions = df_LandfillGasGeneration.at["LT_e","Value"]
    
    for j in range(Lifetime_Emissions):
        if j < int(df_LandfillLeachate.at[0, "CollectionYears"]):
            AverageLeachateCollected.append(df_LandfillLeachate.at[0,'Leachate Collected [m3/ton of waste]'])
        elif j >= int(df_LandfillLeachate.at[0, "CollectionYears"]) and j < int(df_LandfillLeachate.at[1, "CollectionYears"]):
            AverageLeachateCollected.append(df_LandfillLeachate.at[1,'Leachate Collected [m3/ton of waste]'])
        elif j >= int(df_LandfillLeachate.at[1, "CollectionYears"]) and j < int(df_LandfillLeachate.at[2, "CollectionYears"]):
            AverageLeachateCollected.append(df_LandfillLeachate.at[2,'Leachate Collected [m3/ton of waste]'])
        else:
            AverageLeachateCollected.append(df_LandfillLeachate.at[3,'Leachate Collected [m3/ton of waste]'])

    return sum(AverageLeachateCollected) / len(AverageLeachateCollected)



#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""   
##  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
##  %%%%                                           %%%%
##  %%%%          I N P U T    F I L E S           %%%%
##  %%%%                                           %%%%
##  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
"""

import pandas as pd
import numpy as np
import math 


filename = '0_InputData_v6_2_10_9_DataFrom3Drivers_AlmostFull.xlsx'
f = pd.ExcelFile(filename)


## GENERAL
df_General_Factors = f.parse('General', skiprows = 4, nrows=74,  usecols="A:D")
df_General_Factors = df_General_Factors.drop(labels=[11,12,13,14,16,17,18,19,20,21,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,51,52,53,64,65,66,67], axis=0)
df_General_Factors.set_index('Unnamed: 0', inplace=True)



## GENERATION
df_Generation = f.parse('Generation', skiprows = 4, nrows=15,  usecols="C:U")
df_Generation['Waste type'] = SpaceOrDashToUnderscore(df_Generation['Waste type']) # Convert Spaces to Underscores 
WasteType = df_Generation['Waste type'].to_list() # List to iterate
df_Generation.set_index('Waste type', inplace=True)
del df_Generation['Unnamed: 8'], df_Generation['Unnamed: 15']



## PROCESSES COSTS
df_Processes_Cost = f.parse('Processes_Cost', skiprows = 2, nrows=20,  usecols="B:H")
del df_Processes_Cost['Process']
ProcessType = df_Processes_Cost['Process_NameForCode'].to_list() # List to iterate
df_Processes_Cost.set_index('Process_NameForCode', inplace=True)



## REVENUES FROM COPRODUCTS
df_Revenues = f.parse('ValueAddedProducts_Revenues', skiprows = 2, nrows=43,  usecols="B:F")
save_CoproductName = df_Revenues['Coproduct'].to_list() # List to iterate
save_Process = df_Revenues['Process'].to_list() # List to iterate
save_Units = df_Revenues['Units'].to_list() # List to iterate
del df_Revenues['Process'], df_Revenues['Coproduct'], df_Revenues['Units']
CoproductType = df_Revenues['NameForCode'].to_list() # List to iterate
df_Revenues.set_index('NameForCode', inplace=True)



## COLLECTION
   ##### SETS
CollectionType = [] # List to iterate
df_Collection = f.parse('Collection_Cost', skiprows = 1, nrows=1,  usecols="C:H")
CollectionType.extend((df_Collection.at[0, 'Unnamed: 2'], df_Collection.at[0, 'Unnamed: 7']))
del df_Collection
df_Collection = f.parse('Collection_Cost', skiprows = 3, nrows=0,  usecols="C:G")
BinType = df_Collection.columns.values.tolist() # List to iterate
BinType = SpaceOrDashToUnderscore(BinType[:]) # Convert Spaces to Underscores
del df_Collection
   ##### COSTS
for i in range (len(CollectionType)):
    if CollectionType[i] == 'D2D':
        exec("df_Collection_Cost_%s = f.parse('Collection_Cost', skiprows = 3, nrows=3,  usecols='B:G')" %(CollectionType[i]))
        exec("df_Collection_Cost_%s.set_index('Bin type', inplace=True)" %(CollectionType[i]))
        exec("df_Collection_Cost_%s.columns = BinType" %(CollectionType[i]))

    elif CollectionType[i] == 'CS':
        exec("df_Collection_Cost_%s = f.parse('Collection_Cost', skiprows = 3, nrows=3,  usecols='B,H:L')" %(CollectionType[i]))
        exec("df_Collection_Cost_%s.set_index('Bin type', inplace=True)" %(CollectionType[i]))
        exec("df_Collection_Cost_%s.columns = BinType" %(CollectionType[i]))
   ##### COLLECTION FACTORS
for i in range (len(CollectionType)):
    exec("df_Collection_%s_Factor = f.parse('Collection_%s_Factor', skiprows = 4, nrows=15,  usecols='B:G')" %(CollectionType[i], CollectionType[i]))
    exec("df_Collection_%s_Factor['Waste type'] = WasteType" %(CollectionType[i]))
    exec("df_Collection_%s_Factor.set_index('Waste type', inplace=True)" %(CollectionType[i]))
    exec("df_Collection_%s_Factor.columns = BinType" %(CollectionType[i]))
del i
   ##### ENVRONMENTAL IMPACT
for i in range (len(CollectionType)):
    exec("df_Collection_%s_Env = f.parse('Collection_%s_EnvImpact', skiprows = 7, nrows=18,  usecols='C:D')" %(CollectionType[i], CollectionType[i]))
    exec("df_Collection_%s_Env.set_index('Impact Category', inplace=True)" %(CollectionType[i]))
    exec("df_Collection_%s_Env = df_Collection_%s_Env.reindex(columns=df_Collection_%s_Env.columns.tolist() + BinType)" %(CollectionType[i], CollectionType[i], CollectionType[i]))
    
    location = [7,31,55,0,79]
    for k in range(len(BinType)):
        if k != 3:
            exec("dummyDF = f.parse('Collection_%s_EnvImpact', skiprows = location[k], nrows=18,  usecols='E')" %(CollectionType[i]))
            dummyList = dummyDF['Impact per ton of collected waste'].values.tolist()
            exec("df_Collection_%s_Env.at[:,BinType[k]] = dummyList" %(CollectionType[i]))
        else:
            exec("df_Collection_%s_Env.at[:,BinType[k]] = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]" %(CollectionType[i]))



## LANDFILL DATA FROM LIFE CYCLE ASSESSMENT
   ##### CONSTRUCTION
#df_LandfillConstruction = f.parse('Landfill_EnvImpact', skiprows = 7, nrows=18,  usecols="C:E")
#df_LandfillConstruction.set_index('Impact Category', inplace=True)
   ##### LFG GENERATION
df_LandfillGasGeneration = f.parse('Landfill_InputData', skiprows = 5, nrows=26,  usecols="A:D")
for i in range (len(WasteType)):
    df_LandfillGasGeneration.at[i+10,'Unnamed: 0'] = df_LandfillGasGeneration.at[i+10,'Unnamed: 0'] + "_" + WasteType[i]
df_LandfillGasGeneration = df_LandfillGasGeneration.drop(labels=[7,8,9], axis=0)
del df_LandfillGasGeneration['Parameter'],df_LandfillGasGeneration['Units']
df_LandfillGasGeneration.set_index('Unnamed: 0', inplace=True)
   ##### LFG COLLECTION
df_LandfillGasCollectionEff = f.parse('Landfill_InputData', skiprows = 36, nrows=2,  usecols="B:F")
df_LandfillGasCollectionEff.set_index('Parameter', inplace=True)
df_LandfillGasCollectionEff = df_LandfillGasCollectionEff.T
df_LandfillGasCollectionEff.reset_index(inplace = True)
CollectionYears = df_LandfillGasCollectionEff['index'].values.tolist()
CollectionYears = KeepLastYearOnly(CollectionYears)
df_LandfillGasCollectionEff['CollectionYears'] = CollectionYears
del df_LandfillGasCollectionEff['index']
del CollectionYears
   ##### CH4, CO2, AND NMOC (EMISSIONS) CALCULATED WITH LANDGEM (Note: They still need to be divided per Lifetime)
Q_LandfillEmissions, Q_LandfillEmissions_Collected, Q_LandfillEmissions_Uncollected_Oxidized, Q_LandfillEmissions_ss = LandGEM_TonEmissionsPerTonWaste(df_LandfillGasGeneration, WasteType, df_LandfillGasCollectionEff, df_General_Factors)
   ##### LFG COMBUSTION IN ICE
df_LandfillGasCombustion_ICE = f.parse('Landfill_InputData', skiprows = 44, nrows=9,  usecols="A:D")
df_LandfillGasCombustion_ICE = df_LandfillGasCombustion_ICE.drop(labels=[1,2,3,4,5], axis=0)
df_LandfillGasCombustion_ICE.set_index('Unnamed: 0', inplace=True)
del df_LandfillGasCombustion_ICE['Parameter']
   ##### ENVIRONMENTAL IMPACT PER FLARING
df_LandfillGasCombustion_Flare_Env = f.parse('Landfill_EnvImpact', skiprows = 34, nrows=18,  usecols="C:E")
df_LandfillGasCombustion_Flare_Env.set_index('Impact Category', inplace=True)
   ##### ENVIRONMENTAL IMPACT PER COMBUSTION IN ICE
df_LandfillGasCombustion_ICE_Env = f.parse('Landfill_EnvImpact', skiprows = 58, nrows=18,  usecols="C:E")
df_LandfillGasCombustion_ICE_Env.set_index('Impact Category', inplace=True)
   ##### LEACHATE GENERATION AND COLLECTION
df_LandfillLeachate = f.parse('Landfill_InputData', skiprows = 59, nrows=2,  usecols="B:F")
df_LandfillLeachate.set_index('Parameter', inplace=True)
df_LandfillLeachate = df_LandfillLeachate.T
df_LandfillLeachate.reset_index(inplace = True)
CollectionYears = df_LandfillLeachate['index'].values.tolist()
CollectionYears = KeepLastYearOnly(CollectionYears)
df_LandfillLeachate['CollectionYears'] = CollectionYears
del df_LandfillLeachate['index']
del CollectionYears
df_LandfillLeachate['Leachate Collected [m3/ton of waste]'] = df_LandfillLeachate['Leachate generation [m3/ton of waste]'] * df_LandfillLeachate['Leachate collection efficiency [%]']



## MRF DATA
df_MRFeff = f.parse('MRF_InputData', skiprows = 6, nrows=16,  usecols="B:L")
df_MRFeff['Waste type'] = WasteType
df_MRFeff.set_index('Waste type', inplace=True)


## MT for biowaste DATA
df_MTb_eff = f.parse('MTb_InputData', skiprows = 7, nrows=15,  usecols="B:N")
df_MTb_eff['Waste type'] = WasteType
df_MTb_eff.set_index('Waste type', inplace=True)


## MT for mixed waste DATA
df_MTmw_eff = f.parse('MTb_InputData', skiprows = 7, nrows=15,  usecols="B:N")
df_MTmw_eff['Waste type'] = WasteType
df_MTmw_eff.set_index('Waste type', inplace=True)


## ENVRONMENTAL IMPACT for MRF, MTb, and MTmw
df_MRF_MT_Env = f.parse('MRF_EnvImpact', skiprows = 7, nrows=18,  usecols="C:D")
df_MRF_MT_Env.set_index('Impact Category', inplace=True)
df_MRF_MT_Env = df_MRF_MT_Env.reindex(columns=df_MRF_MT_Env.columns.tolist() + BinType)

location = [7,7,31,31,55]
for k in range(len(BinType)):
    if k == 0 or k == 3:
        dummyDF = f.parse('MT_EnvImpact', skiprows = location[k], nrows=18,  usecols="E")
        dummyList = dummyDF['Impact per ton of separated waste'].values.tolist()
        df_MRF_MT_Env.at[:,BinType[k]] = dummyList
    else:
        dummyDF = f.parse('MRF_EnvImpact', skiprows = location[k], nrows=18,  usecols="E")
        dummyList = dummyDF['Impact per ton of separated waste'].values.tolist()
        df_MRF_MT_Env.at[:,BinType[k]] = dummyList


## LEGISLATIVE RESTRICTIONS
df_Legislative = f.parse('Legislation_Industry', skiprows = 4, nrows=16,  usecols="B:D")
index = WasteType.copy()
index.append('TOTAL_Recyclables')
df_Legislative['Waste type'] = index
df_Legislative.set_index('Waste type', inplace=True)
del index


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""   
##  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
##  %%%%                                                        %%%%
##  %%%%          O P T I M I Z A T I O N   M O D E L           %%%%
##  %%%%                                                        %%%%
##  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
"""

import pyomo.environ as pyo
from pyomo.opt import SolverFactory


## Create the MODEL
model = pyo.ConcreteModel()


# Create SETS
model.WasteType = set(list(range(len(WasteType))))
model.CollectionType = set(list(range(len(CollectionType))))
model.BinType = set(list(range(len(BinType))))
model.ProcessType = set(list(range(len(ProcessType))))
model.CoproductType = set(list(range(len(CoproductType))))
#model.FluxType = set(list(range(len(FluxType))))


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

## Indexes along the code
   ##   i - WasteType, as per model.WasteType
   ##   j - CollectionType, as per model.CollectionType
   ##   k - BinType, as per model.BinType
   ##   l - ProcessType
   ##   n - CoproductType

#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------


"""                      
############     D E C L A R I N G    V A R I A B L E S     ###############
"""                      

## Variable to decide the percentage of each WasteType collected by CollectionType (Door2Door (D2D) or Curbside (Cs))
   # Rationale: We mimic a global "r" factor (instead of taking it per bin type or even waste type), because in Portugal, all bins in the collection D2D are picked-up together. Thus, it makes more sense to decide only the % of waste that will be collected D2D and drag all bins % accordingly.
   #            Before knowing this, I thought that it could be useful to control this "r" factor per waste type, because that could help us to identify the key waste type and implement marketing mechanisms to foster its prevention and recycling, instead of spending resources in trying to improve everything in a costly way.
   # Note : sum(r[j]) = 1 is a constraint
model.r = pyo.Var(model.CollectionType, domain=pyo.NonNegativeReals)

# Variable to calculate the increase in the installed capacity (size) for collection
model.sc = pyo.Var(model.CollectionType, model.BinType, domain=pyo.NonNegativeReals)

## Variables for the mass picked for collection j and bin type k
model.mp = pyo.Var(model.CollectionType, model.BinType, domain=pyo.NonNegativeReals)

## Qty. of coproducts from processes
model.x = pyo.Var(model.WasteType, model.CoproductType, domain=pyo.NonNegativeReals)
model.xs = pyo.Var(model.CoproductType, domain=pyo.NonNegativeReals) # When residues are NOT distiguished (single stream)
model.x_dummy = pyo.Var(model.WasteType, model.CoproductType, domain=pyo.NonNegativeReals) # To keep the detail of WasteType for Landfill (methane), LandfillRB (biogas), MRF
model.xs_dummy = pyo.Var(model.CoproductType, domain=pyo.NonNegativeReals) # To keep the detail of WasteType for Landfill (methane), LandfillRB (biogas), MRF

## Variables for the mass input to each technology/process
model.mi = pyo.Var(model.WasteType, model.ProcessType, domain=pyo.NonNegativeReals)
model.msi = pyo.Var(model.ProcessType, domain=pyo.NonNegativeReals) # When residues are NOT distiguished (single stream)
model.mi_dummy = pyo.Var(model.WasteType, model.ProcessType, domain=pyo.NonNegativeReals) # To keep the detail of WasteType for Landfill (methane), LandfillRB (biogas), MRF
model.msi_dummy = pyo.Var(model.ProcessType, domain=pyo.NonNegativeReals) # To keep the detail of WasteType for Landfill (methane), LandfillRB (biogas), MRF
model.mi_dummy_dummy = pyo.Var(model.WasteType, [1,2], domain=pyo.NonNegativeReals)
                         
# Variable to calculate the increase in the installed capacity (size) for processes
model.sp = pyo.Var(model.ProcessType, domain=pyo.NonNegativeReals)

## Variables for the mass output from each technology/process
model.mo = pyo.Var(model.WasteType, model.ProcessType, domain=pyo.NonNegativeReals) # Output1
model.mso = pyo.Var(model.ProcessType, domain=pyo.NonNegativeReals) # Output1 - When residues are NOT distiguished (single stream)

## Transitional variables for Landfill
  ## For the INITIAL mass input to Landfill
model.mi0 = pyo.Var(model.WasteType, [0], domain=pyo.NonNegativeReals)
  ## For the decision between Landfill and thermal treatments, from MRF and MT
model.mi_LvsTT = pyo.Var(model.WasteType, [0], domain=pyo.NonNegativeReals)

## Transitional variables for other processes
  ## For the syngas that goes from gasification to fermentation processes
model.moG_sg = pyo.Var([12], domain=pyo.NonNegativeReals)
  ## For the syngas that goes from pyrolysis to fermentation processes
model.moP_sg = pyo.Var([13], domain=pyo.NonNegativeReals)
  ## For the mass-in from MTbio to AD
model.mi_MTb = pyo.Var(model.WasteType, [8], domain=pyo.NonNegativeReals)
  ## Proportion of WM bin that goes to landfill
#model.d = pyo.Var([0], domain=pyo.NonNegativeReals)
#model.d = pyo.Var(within=pyo.NonNegativeReals, bounds=(0,1), initialize=0.5)

## Variables for calculate the emissions per technology/process (inclusing collection) [kg CO2 eq]
model.Ec = pyo.Var(model.CollectionType, domain=pyo.NonNegativeReals)
model.Ep = pyo.Var(model.ProcessType, domain=pyo.NonNegativeReals)


"""                      
############     D E C L A R I N G    P A R A M E T E R S     ###############
"""                      

## MRF and MT
MinPurity_Recycling = 0.1

##########

## GENERAL
#I = 0.05 # Interest rate
M = 1000 # Big constant for binary constraints


## LANDFILL
d = df_General_Factors.at['d','Value']   ## Proportion of waste from the MixedWaste bin that goes to landfill
#OperationTime = ReadValueFromExcel(filename, 'Landfill_InputData', 'E', 10)
ElectGenEff = 0.3
HeatGenEff = 0.65
HV_eq = ReadValueFromExcel(filename, 'Landfill_InputData', 'E', 26)
LFG_density = ReadValueFromExcel(filename, 'Landfill_InputData', 'E', 23)
#LFGg_ss = 0.8 # Landfill gas generation eff for single stream


## LANDFILL and AD
BMP_ss = 0.09 # Biochemical methane potential of the feedstock for single stream [ton of CH4/ton]



## COMPOSTING
CompostingRate = 0.6


## INCINERATION
eff_Incineration_Electricity = 0.4
#eff_Incineration_Heat = 0.6
eff_Incineration_Slag = 0.1


## GASIFICATION
eff_Gasification_Syngas = 0.6


## PYROLYSIS
eff_Pyrolysis_Bio_oil = 0.1
eff_Pyrolysis_Syngas = 0.8


## SHIFT REACTION
eff_Shift_Reaction_To_Hydrogen = 0.9


## FISCHER TROPSCH
eff_Fischer_Tropsch_To_Light_hydrocarbons = 0.9
eff_Fischer_Tropsch_To_Biodiesel = 0.8


## METHANOL SYNTHESIS
eff_Methanol_Synthesis = 0.9


## ALCOHOL SYNTHESIS
eff_Ethanol_Synthesis = 0.9


## FERMENTATION
eff_Ethanol_Fermentation = 0.9


del filename

#
"""
############     O B J E C T I V E    F U N C T I O N     ###############
"""

def z(model):
    return CollectionCAPEX(model) + CollectionOPEX(model) + ProcessesCAPEX(model) + ProcessesOPEX(model) - Revenues(model)

def CollectionCAPEX(model):
    return sum ( float(df_Collection_Cost_D2D.at['CAPEX [EUR/ton]', BinType[k]]) * model.sc[0,k] for k in model.BinType ) \
        + sum ( float(df_Collection_Cost_CS.at['CAPEX [EUR/ton]', BinType[k]]) * model.sc[1,k] for k in model.BinType ) \
    # sum ( CAPEX[j,k] * Increased installed capacity [j,k] )
    
def CollectionOPEX(model):
    return sum( float(df_Collection_Cost_D2D.at['OPEX [EUR/ton]', BinType[k]]) * model.mp[0,k] for k in model.BinType ) \
           + sum ( float(df_Collection_Cost_CS.at['OPEX [EUR/ton]', BinType[k]]) * model.mp[1,k] for k in model.BinType )
    # sum ( OPEX[j,k] * wasteGeneration[i] * r[j] * CollectionFactor[i,k] )
    
def ProcessesCAPEX(model):
    return sum ( float(df_Processes_Cost.at[ProcessType[l],'CAPEX [EUR/ton]']) \
                * model.sp[l] for l in model.ProcessType )
    # sum ( CAPEX[l] * ( sum ( mi[i,l] ) + msi[l] - Installed capacity [l] ) )

def ProcessesOPEX(model):
    return sum ( \
               float(df_Processes_Cost.at[ProcessType[l], 'OPEX [EUR/ton]']) \
                    * ( sum(model.mi[i,l] for i in model.WasteType) + model.msi[l] ) \
            for l in model.ProcessType )
    # sum ( OPEX[l] * ( sum ( mi[i,l] ) + msi[l] ) )
  
def Revenues(model):
    return sum ( \
                float(df_Revenues.at[CoproductType[n],'Commercial Price [EUR/unit]']) \
                     * ( sum(model.x[i,n] for i in model.WasteType) + model.xs[n]) \
                for n in model.CoproductType)
    # sum ( Revenue[n] * x[n] )

model.Solution = pyo.Objective(rule=z)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#####################       C O N S T R A I N T S       #####################
"""


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Waste Collection   #######
"""

# The different CollectionType must cover 100% of the collection for each WasteType (No. constraint = 1)
def Constraint1 (model):
    return sum( model.r[j] for j in model.CollectionType ) == 1
model.Constraint1 = pyo.Constraint(rule=Constraint1)


# Mass picked in collection j and bin type k (No. constraint = j*k )
def Constraint2a (model, k):
    return model.mp[0,k] == \
            sum( float(df_Generation.at[WasteType[i], 'mass [tons]']) * model.r[0] \
                * float(df_Collection_D2D_Factor.at[WasteType[i], BinType[k]]) for i in model.WasteType)
model.Constraint2a = pyo.Constraint(model.BinType, rule=Constraint2a)

def Constraint2b (model, k):
    return model.mp[1,k] == \
            sum( float(df_Generation.at[WasteType[i], 'mass [tons]']) * model.r[1] \
                * float(df_Collection_CS_Factor.at[WasteType[i], BinType[k]]) for i in model.WasteType)
model.Constraint2b = pyo.Constraint(model.BinType, rule=Constraint2b)


# Installed capacity increase for collection (No. constraint = j*k )
def Constraint3a (model, k):
    return model.sc[0,k] + \
            float(df_Collection_Cost_D2D.at['Existant installed capacity [ton]', BinType[k]]) >= \
            model.mp[0,k]
model.Constraint3a = pyo.Constraint(model.BinType, rule=Constraint3a)

def Constraint3b (model, k):
    return model.sc[1,k] + \
            float(df_Collection_Cost_CS.at['Existant installed capacity [ton]', BinType[k]]) >= \
            model.mp[1,k]
model.Constraint3b = pyo.Constraint(model.BinType, rule=Constraint3b)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Waste Collection CO2 EMISSIONS   #######
"""

## kg CO2 emissions in D2D collection (No. constraint = 1)
def Constraint4a (model):
    return model.Ec[0] == model.r[0] \
            * sum( float(df_Generation.at[WasteType[i], 'mass [tons]'])  \
            * sum(
            float(df_Collection_D2D_Factor.at[WasteType[i], BinType[k]]) * float(df_Collection_D2D_Env.at['Global warming',BinType[k]])
            for k in model.BinType )
            for i in model.WasteType )
model.Constraint4a = pyo.Constraint(rule=Constraint4a)

## kg CO2 emissions in CS collection (No. constraint = 1)
def Constraint4b (model):
    return model.Ec[1] == model.r[1] \
            * sum( float(df_Generation.at[WasteType[i], 'mass [tons]'])  \
            * sum(
            float(df_Collection_CS_Factor.at[WasteType[i], BinType[k]]) * float(df_Collection_CS_Env.at['Global warming',BinType[k]])
            for k in model.BinType )
            for i in model.WasteType )
model.Constraint4b = pyo.Constraint(rule=Constraint4b)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Landfill (producing methane for flare)   #######
"""

#LFGg = AveLandfillGasGeneration(df_LandfillGasGeneration, WasteType)


# Initial input mass to landfill (No. constraint = i)
def Constraint200 (model, i):
    return model.mi0[i,0] == d * float(df_Generation.at[WasteType[i],'mass [tons]']) \
                        * ( model.r[0] * float(df_Collection_D2D_Factor.at[WasteType[i],'Mixed_Waste']) \
                           + model.r[1] * float(df_Collection_CS_Factor.at[WasteType[i],'Mixed_Waste']) \
                        )
model.Constraint200 = pyo.Constraint(model.WasteType, rule=Constraint200)


# Total input mass (vector) to landfill (No. constraint = i)
def Constraint201 (model, i):
    return model.mi[i,0] == model.mi0[i,0] + model.mi_LvsTT[i,0]
model.Constraint201 = pyo.Constraint(model.WasteType, rule=Constraint201)


# Total input mass (single stream) to landfill (No. constraint = 1)
def Constraint202 (model):
    return model.msi[0] == model.mso[8] + model.mso[10] \
                        + model.mso[11] + model.mso[12] + model.mso[13] \
                        + model.mso[14] + model.mso[15] + model.mso[16] + model.mso[17] + model.mso[18] + model.mso[19]
model.Constraint202 = pyo.Constraint(rule=Constraint202)


# Installed capacity increase of Landfill (No. constraint = 1)
def Constraint203 (model):
    return model.sp[0] + df_Processes_Cost.at[ProcessType[0], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,0] for i in model.WasteType) + model.msi[0]
model.Constraint203 = pyo.Constraint(rule=Constraint203)


# Production of Methane [tons] for flaring and Input mass to LandfillRB
  # Production of Methane [tons] for flaring (No. constraint = 1)
def Constraint204a (model):
    return model.xs[0] == sum(model.x_dummy[i,0] for i in model.WasteType) + model.xs_dummy[0]
model.Constraint204a = pyo.Constraint(rule=Constraint204a)

  # Input mass to LandfillRB (No. constraint = 1)
def Constraint204b (model):
    return model.msi[1] == sum(model.mi_dummy[i,1] for i in model.WasteType) + model.msi_dummy[1]
model.Constraint204b = pyo.Constraint(rule=Constraint204b)

  # For EMISSIONS calulation in LandfillRB: Value of x_dummy[i,0] and mi_dummy[i,1] (No. constraint = i)
def Constraint204c (model, i):
    return model.x_dummy[i,0] + model.mi_dummy[i,1] == \
                ( model.mi[i,0] * float( Q_LandfillEmissions_Collected.at[WasteType[i], 'Q_CH4']) \
                 ) / df_Processes_Cost.at['Landfill','Lifetime_Years']
model.Constraint204c = pyo.Constraint(model.WasteType, rule=Constraint204c)

  # For EMISSIONS calulation in LandfillRB: Value of xs_dummy[0] and msi_dummy[1] (No. constraint = 1)
def Constraint204d (model):
    return model.xs_dummy[0] + model.msi_dummy[1] == \
                ( model.msi[0] * float( Q_LandfillEmissions_ss.at['Collected', 'Q_CH4']) \
                 ) / df_Processes_Cost.at['Landfill','Lifetime_Years']
model.Constraint204d = pyo.Constraint(rule=Constraint204d)

## Production of Methane [tons] for flaring and Input mass to LandfillRB (No. constraint = 1)
#def Constraint204 (model):
#    return model.xs[0] + model.msi[1] == \
#            ( sum(model.mi[i,0] * float( Q_LandfillEmissions_Collected.at[WasteType[i], 'Q_CH4']) for i in model.WasteType) \
#              + model.msi[0] * float( Q_LandfillEmissions_ss.at['Collected', 'Q_CH4']) \
#            ) / df_Processes_Cost.at['Landfill','Lifetime_Years']
#model.Constraint204 = pyo.Constraint(rule=Constraint204)


# Production of leachate from the landfill (No. constraint = 1)
def Constraint205 (model):
    return model.xs[1] == sum (model.mi[i,0] * AverageLeachateCollected(df_LandfillLeachate,df_LandfillGasGeneration) for i in model.WasteType) \
                            + model.msi[0] * AverageLeachateCollected(df_LandfillLeachate,df_LandfillGasGeneration)
model.Constraint205 = pyo.Constraint(rule=Constraint205)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   LandfillRB (producing biogas)   #######
"""

# Installed capacity increase of LandfillRB (No. constraint = 1)
def Constraint210 (model):
    return model.sp[1] + df_Processes_Cost.at[ProcessType[1], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,1] for i in model.WasteType) + model.msi[1]
model.Constraint210 = pyo.Constraint(rule=Constraint210)


# Production of biogas [ton] from landfillRB and Input mass to LandfillRE (No. constraint = i+1)
  # Production of biogas [tons] for landfillRB (No. constraint = 1)
def Constraint211a (model):
    return model.xs[2] == sum(model.x_dummy[i,2] for i in model.WasteType) + model.xs_dummy[2]
model.Constraint211a = pyo.Constraint(rule=Constraint211a)

  # Input mass to LandfillRE (No. constraint = 1)
def Constraint211b (model):
    return model.msi[2] == sum(model.mi_dummy[i,2] for i in model.WasteType) + model.msi_dummy[2]
model.Constraint211b = pyo.Constraint(rule=Constraint211b)

  # For EMISSIONS calulation in LandfillRE: Value of x_dummy[i,2] and mi_dummy[i,2] (No. constraint = i)
def Constraint211c (model, i):
    return model.x_dummy[i,2] + model.mi_dummy[i,2] == \
                ( model.mi_dummy[i,1] * float( Q_LandfillEmissions_Collected.at[WasteType[i], 'Q_CH4']) \
                 ) / df_Processes_Cost.at['Landfill','Lifetime_Years']
model.Constraint211c = pyo.Constraint(model.WasteType, rule=Constraint211c)

  # For EMISSIONS calulation in LandfillRE: Value of xs_dummy[2] and msi_dummy[2] (No. constraint = 1)
def Constraint211d (model):
    return model.xs_dummy[2] + model.msi_dummy[2] == \
                ( model.msi[1] * float( Q_LandfillEmissions_ss.at['Collected', 'Q_CH4']) \
                 ) / df_Processes_Cost.at['Landfill','Lifetime_Years']
model.Constraint211d = pyo.Constraint(rule=Constraint211d)

## Production of biogas [ton] from landfillRB and Input mass to landfillRE (No. constraint = i+1)
#def Constraint211 (model, i):
#    return model.xs[2] + model.msi[2] == model.msi[1] * df_LandfillGasCombustion_ICE.at['ePur','Value']
#model.Constraint211 = pyo.Constraint(model.WasteType, rule=Constraint211)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   LandfillRE (producing electricity)   #######
"""

# Installed capacity increase of LandfillRE (No. constraint = 1)
def Constraint220 (model):
    return model.sp[2] + df_Processes_Cost.at[ProcessType[2], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,2] for i in model.WasteType) + model.msi[2]
model.Constraint220 = pyo.Constraint(rule=Constraint220)


# Production of electricity [kWh] from landfillRE (No. constraint = 1)
def Constraint221 (model):
    return model.xs[3] == \
            model.msi[2] * df_LandfillGasCombustion_ICE.at['LFG_HV','Value'] * df_LandfillGasCombustion_ICE.at['eEff','Value']
model.Constraint221 = pyo.Constraint(rule=Constraint221)


# Production of heat [kWh] from landfillRE (No. constraint = 1)
def Constraint222 (model):
    return model.xs[4] == \
            model.msi[2] * df_LandfillGasCombustion_ICE.at['LFG_HV','Value'] * df_LandfillGasCombustion_ICE.at['hEff','Value']
model.Constraint222 = pyo.Constraint(rule=Constraint222)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Landfill CO2 EMISSIONS   #######
"""

## Equivalent input mass, per WasteType, for LandfillRB and LandfillRE (No. constraint = 2*i)
def Constraint230a (model, i):
    if float(Q_LandfillEmissions_Collected.at[WasteType[i],'Q_CH4']) > 0:
        return model.mi_dummy_dummy[i,1] == model.mi_dummy[i,1] / float(Q_LandfillEmissions_Collected.at[WasteType[i],'Q_CH4'])
    else:
        return model.mi_dummy_dummy[i,1] == 0
model.Constraint230a = pyo.Constraint(model.WasteType, rule=Constraint230a)

def Constraint230b (model, i):
    if float(Q_LandfillEmissions_Collected.at[WasteType[i],'Q_CH4']) > 0:
        return model.mi_dummy_dummy[i,2] == model.mi_dummy[i,2] \
            / ( float(Q_LandfillEmissions_Collected.at[WasteType[i],'Q_CH4']) * float(df_LandfillGasCombustion_ICE.at['ePur','Value']) )
    else:
        return model.mi_dummy_dummy[i,2] == 0
model.Constraint230b = pyo.Constraint(model.WasteType, rule=Constraint230b)



## kg CO2 emissions in Landfill (No. constraint = 1)
   ## 1) Due to flaring, per waste type i (extracting mass entering to LandfillRB and LandfillRE)
   ##    1a) CO2 emissions due CO2 collected, uncollected (adding CH4 oxidation),
   ##    1b) CO2 equivalent emissions due CH4 uncollected (substracting CH4 oxidation)
   ##        NOTE: collected CH4 is the one used for flaring, biogas production, or ICE combustion
   ## 2) Due to flaring, per single stream (extracting mass entering to LandfillRB and LandfillRE)
   ##    2a) CO2 emissions due CO2 collected, uncollected (adding CH4 oxidation),
   ##    2b) CO2 equivalent emissions due CH4 collected, uncollected (substracting CH4 oxidation)
   ## 3) CO2 emissions due to flare burning
def Constraint231 (model):
    return model.Ep[0] == \
            sum( ( model.mi[i,0] - model.mi_dummy_dummy[i,1] - model.mi_dummy_dummy[i,2] ) \
                * float( ( (Q_LandfillEmissions_Collected.at[WasteType[i],'Q_CO2'] + Q_LandfillEmissions_Uncollected_Oxidized.at[WasteType[i],'Q_CO2']) / df_Processes_Cost.at['Landfill','Lifetime_Years'] ) \
                         + (df_General_Factors.at['CH4_CO2eq','Value'] * (Q_LandfillEmissions_Uncollected_Oxidized.at[WasteType[i],'Q_CH4'])/ df_Processes_Cost.at['Landfill','Lifetime_Years'] ) \
                        ) for i in model.WasteType ) \
            + ( model.msi[0] - (model.msi_dummy[1]/float( Q_LandfillEmissions_ss.at['Collected', 'Q_CH4'])) - ( model.msi_dummy[2] / ( float(Q_LandfillEmissions_ss.at['Collected', 'Q_CH4']) * float(df_LandfillGasCombustion_ICE.at['ePur','Value']) ) ) ) \
                * float( ( (Q_LandfillEmissions_ss.at['Collected','Q_CO2'] + Q_LandfillEmissions_ss.at['Uncollected_Oxidized','Q_CO2']) / df_Processes_Cost.at['Landfill','Lifetime_Years'] ) \
                         + ( df_General_Factors.at['CH4_CO2eq','Value'] * Q_LandfillEmissions_ss.at['Uncollected_Oxidized','Q_CH4'] / df_Processes_Cost.at['Landfill','Lifetime_Years'] ) \
                        ) \
            + ( model.xs[0] * float(df_LandfillGasCombustion_Flare_Env.at['Global warming','Impact per ton of landfilled waste']) )
model.Constraint231 = pyo.Constraint(rule=Constraint231)


## kg CO2 emissions in LandfillRB (No. constraint = 1)
   ## 1) Due to biogas production, per waste type i (extracting mass entering to LandfillRB and LandfillRE)
   ##    1a) CO2 emissions due CO2 collected, uncollected (adding CH4 oxidation),
   ##    1b) CO2 equivalent emissions due CH4 collected, uncollected (substracting CH4 oxidation)
   ## 2) Due to biogas production, per single stream (extracting mass entering to LandfillRB and LandfillRE)
   ##    2a) CO2 emissions due CO2 collected, uncollected (adding CH4 oxidation),
   ##    2b) CO2 equivalent emissions due CH4 collected, uncollected (substracting CH4 oxidation)
def Constraint232 (model):
    return model.Ep[1] == \
            sum( ( model.mi_dummy_dummy[i,1] - model.mi_dummy_dummy[i,2] ) \
                * float( (Q_LandfillEmissions_Collected.at[WasteType[i],'Q_CO2'] + Q_LandfillEmissions_Uncollected_Oxidized.at[WasteType[i],'Q_CO2']) / df_Processes_Cost.at['Landfill','Lifetime_Years'] \
                         + (df_General_Factors.at['CH4_CO2eq','Value'] * Q_LandfillEmissions_Uncollected_Oxidized.at[WasteType[i],'Q_CH4'] / df_Processes_Cost.at['Landfill','Lifetime_Years'] ) \
                        ) for i in model.WasteType ) \
            + ( ( model.msi_dummy[1]/float( Q_LandfillEmissions_ss.at['Collected', 'Q_CH4']) ) - ( model.msi_dummy[2] / float(Q_LandfillEmissions_ss.at['Collected', 'Q_CH4']) / df_LandfillGasCombustion_ICE.at['ePur','Value'] ) ) \
                * float( ( (Q_LandfillEmissions_ss.at['Collected','Q_CO2'] + Q_LandfillEmissions_ss.at['Uncollected_Oxidized','Q_CO2']) / df_Processes_Cost.at['Landfill','Lifetime_Years'] ) \
                         + ( df_General_Factors.at['CH4_CO2eq','Value'] * Q_LandfillEmissions_ss.at['Uncollected_Oxidized','Q_CH4'] / df_Processes_Cost.at['Landfill','Lifetime_Years'] ) \
                        )
model.Constraint232 = pyo.Constraint(rule=Constraint232)


## kg CO2 emissions in LandfillRE (No. constraint = 1)
   ## 1) Due to biogas production, per waste type i (extracting mass entering to LandfillRB and LandfillRE)
   ##    1a) CO2 emissions due CO2 collected, uncollected (adding CH4 oxidation),
   ##    1b) CO2 equivalent emissions due CH4 collected, uncollected (substracting CH4 oxidation)
   ## 2) Due to biogas production, per single stream (extracting mass entering to LandfillRB and LandfillRE)
   ##    2a) CO2 emissions due CO2 collected, uncollected (adding CH4 oxidation),
   ##    2b) CO2 equivalent emissions due CH4 collected, uncollected (substracting CH4 oxidation)
   ## 3) CO2 emissions due to burning in ICE
def Constraint233 (model):
    return model.Ep[2] == \
            sum( model.mi_dummy_dummy[i,2] \
                * float( ( (Q_LandfillEmissions_Collected.at[WasteType[i],'Q_CO2'] + Q_LandfillEmissions_Uncollected_Oxidized.at[WasteType[i],'Q_CO2']) / df_Processes_Cost.at['Landfill','Lifetime_Years'] ) \
                         + ( df_General_Factors.at['CH4_CO2eq','Value'] * Q_LandfillEmissions_Uncollected_Oxidized.at[WasteType[i],'Q_CH4'] / df_Processes_Cost.at['Landfill','Lifetime_Years'] ) \
                        ) for i in model.WasteType ) \
            + ( model.msi_dummy[2] / float( Q_LandfillEmissions_ss.at['Collected', 'Q_CH4']) / df_LandfillGasCombustion_ICE.at['ePur','Value'] ) \
                * float(  ( (Q_LandfillEmissions_ss.at['Collected','Q_CO2'] + Q_LandfillEmissions_ss.at['Uncollected_Oxidized','Q_CO2']) / df_Processes_Cost.at['Landfill','Lifetime_Years'] ) \
                         + ( df_General_Factors.at['CH4_CO2eq','Value'] * Q_LandfillEmissions_ss.at['Uncollected_Oxidized','Q_CH4'] / df_Processes_Cost.at['Landfill','Lifetime_Years'] ) \
                        ) \
            + model.xs[2] * float(df_LandfillGasCombustion_ICE_Env.at['Global warming','Impact per ton of landfilled waste'])
model.Constraint233 = pyo.Constraint(rule=Constraint233)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   MRFpc   #######
"""

# Input mass to MRFpc (No. constraint = i)
def Constraint10 (model, i):
    return model.mi[i,3] == float(df_Generation.at[WasteType[i],'mass [tons]']) \
                        * ( model.r[0] * float(df_Collection_D2D_Factor.at[WasteType[i],'Paper_Cardboard']) \
                           + model.r[1] * float(df_Collection_CS_Factor.at[WasteType[i],'Paper_Cardboard']) \
                          )
model.Constraint10 = pyo.Constraint(model.WasteType, rule=Constraint10)


# Installed capacity increase of MRFpc (No. constraint = 1)
def Constraint11 (model):
    return model.sp[3] + df_Processes_Cost.at[ProcessType[3], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,3] for i in model.WasteType) + model.msi[3]
model.Constraint11 = pyo.Constraint(rule=Constraint11)

#----------

# Production of Recycled Paper from MRFpc (No. constraint = i)
def Constraint12 (model, i):
    return model.x[i,5] == model.mi[i,3] * df_MRFeff.at[WasteType[i],'MRFpc_Paper']
model.Constraint12 = pyo.Constraint(model.WasteType, rule=Constraint12)


## Purity required for the Recycled Paper [No. of constraints = 1]
#def Constraint13 (model):
#    return model.x[0,5] >= MinPurity_Recycling * sum(model.x[i,5] for i in model.WasteType)
##    return model.x[0,5] >= df_Legislative.at['Paper', 'Target for PURITY rate'] * sum(model.x[i,5] for i in model.WasteType)
#model.Constraint13 = pyo.Constraint(rule=Constraint13)


## Recycling rate target for Recycled Paper [No. of constraints = 1]
#def Constraint14 (model):
#    if df_Generation.at[WasteType[0],'mass [tons]'] > 0:
#        return ( sum(model.x[i,5] for i in model.WasteType) + sum(model.x[i,15] for i in model.WasteType) ) / float (df_Generation.at[WasteType[0],'mass [tons]']) >= df_Legislative.at['Paper', 'Target for RECYCLING rate']
#    else:
#        return pyo.Constraint.Skip
#model.Constraint14 = pyo.Constraint(rule=Constraint14)

#----------

# Production of Recycled Cardboard from MRFpc (No. constraint = i)
def Constraint15 (model, i):
    return model.x[i,6] == model.mi[i,3] * df_MRFeff.at[WasteType[i],'MRFpc_Cardboard']
model.Constraint15 = pyo.Constraint(model.WasteType, rule=Constraint15)


## Purity required for the Recycled Cardboard [No. of constraints = 1]
#def Constraint16 (model):
#    return model.x[1,6] >= MinPurity_Recycling * sum(model.x[i,6] for i in model.WasteType)
#model.Constraint16 = pyo.Constraint(rule=Constraint16)


## Recycling rate target for Recycled Cardboard [No. of constraints = 1]
#def Constraint17 (model):
#    if df_Generation.at[WasteType[1],'mass [tons]'] > 0:
#        return ( sum(model.x[i,6] for i in model.WasteType) + sum(model.x[i,16] for i in model.WasteType) ) / float (df_Generation.at[WasteType[1],'mass [tons]']) >= MinRecyclingRate
#    else:
#        return pyo.Constraint.Skip
#model.Constraint17 = pyo.Constraint(rule=Constraint17)

#----------

# Outflow of MRFpc (No. of constraints = i)
def Constraint18 (model, i):
    return model.mo[i,3] == model.mi[i,3] - model.x[i,5] - model.x[i,6]
model.Constraint18 = pyo.Constraint(model.WasteType, rule=Constraint18)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   MRFpm   #######
"""

# Input mass to MRFpm (No. constraint = i)
def Constraint20 (model, i):
    return model.mi[i,4] == float(df_Generation.at[WasteType[i],'mass [tons]']) \
                        * ( model.r[0] * float(df_Collection_D2D_Factor.at[WasteType[i],'Plastic_Metal']) \
                           + model.r[1] * float(df_Collection_CS_Factor.at[WasteType[i],'Plastic_Metal']) \
                        )
model.Constraint20 = pyo.Constraint(model.WasteType, rule=Constraint20)


# Installed capacity increase of MRFpm (No. constraint = 1)
def Constraint21 (model):
    return model.sp[4] + df_Processes_Cost.at[ProcessType[4], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,4] for i in model.WasteType) + model.msi[4]
model.Constraint21 = pyo.Constraint(rule=Constraint21)


# Binary variable to activate the CAPEX of MRFpm (No. constraint = 2)
#def Constraint21a (model): # 1st part
#    return sum ( model.mi[i,4] for i in model.WasteType) >= 0.0000000001 - M*( 1 - model.bp[4] )
#model.Constraint21a = pyo.Constraint(rule=Constraint21a)
#def Constraint21b(model): # 2nd part
#    return sum ( model.mi[i,4] for i in model.WasteType) <= 0 + M * model.bp[4]
#model.Constraint21b = pyo.Constraint(rule=Constraint21b)

#----------

# Production of Recycled Composites from MRFpm (No. constraint = i)
def Constraint22 (model, i):
    return model.x[i,7] == model.mi[i,4] * df_MRFeff.at[WasteType[i],'MRFpm_Composites']
model.Constraint22 = pyo.Constraint(model.WasteType, rule=Constraint22)


## Purity required for the Recycled Composites [No. of constraints = 1]
#def Constraint23 (model):
#    return model.x[2,7] >= MinPurity_Recycling * sum(model.x[i,7] for i in model.WasteType)
#model.Constraint23 = pyo.Constraint(rule=Constraint23)


## Recycling rate target for Recycled Composites [No. of constraints = 1]
#def Constraint24 (model):
#    if df_Generation.at[WasteType[2],'mass [tons]'] > 0:
#        return ( sum(model.x[i,7] for i in model.WasteType) + sum(model.x[i,17] for i in model.WasteType) ) / float (df_Generation.at[WasteType[2],'mass [tons]']) >= MinRecyclingRate
#    else:
#        return pyo.Constraint.Skip
#model.Constraint24 = pyo.Constraint(rule=Constraint24)

#----------

# Production of Recycled PE from MRFpm (No. constraint = i)
def Constraint25 (model, i):
    return model.x[i,8] == model.mi[i,4] * df_MRFeff.at[WasteType[i],'MRFpm_PE']
model.Constraint25 = pyo.Constraint(model.WasteType, rule=Constraint25)


## Purity required for the Recycled PE [No. of constraints = 1]
#def Constraint26 (model):
#    return model.x[3,8] >= MinPurity_Recycling * sum(model.x[i,8] for i in model.WasteType)
#model.Constraint26 = pyo.Constraint(rule=Constraint26)


## Recycling rate target for Recycled PE [No. of constraints = 1]
#def Constraint27 (model):
#    if df_Generation.at[WasteType[3],'mass [tons]'] > 0:
#        return ( sum(model.x[i,8] for i in model.WasteType) + sum(model.x[i,18] for i in model.WasteType) ) / float (df_Generation.at[WasteType[3],'mass [tons]']) >= MinRecyclingRate
#    else:
#        return pyo.Constraint.Skip
#model.Constraint27 = pyo.Constraint(rule=Constraint27)

#----------

# Production of Recycled PET from MRFpm (No. constraint = i)
def Constraint28 (model, i):
    return model.x[i,9] == model.mi[i,4] * df_MRFeff.at[WasteType[i],'MRFpm_PET']
model.Constraint28 = pyo.Constraint(model.WasteType, rule=Constraint28)


## Purity required for the Recycled PET [No. of constraints = 1]
#def Constraint29 (model):
#    return model.x[4,9] >= MinPurity_Recycling * sum(model.x[i,9] for i in model.WasteType)
#model.Constraint29 = pyo.Constraint(rule=Constraint29)


## Recycling rate target for Recycled PET [No. of constraints = 1]
#def Constraint30 (model):
#    if df_Generation.at[WasteType[4],'mass [tons]'] > 0:
#        return ( sum(model.x[i,9] for i in model.WasteType) + sum(model.x[i,19] for i in model.WasteType) ) / float (df_Generation.at[WasteType[4],'mass [tons]']) >= MinRecyclingRate
#    else:
#        return pyo.Constraint.Skip
#model.Constraint30 = pyo.Constraint(rule=Constraint30)

#----------

# Production of Recycled HDPE from MRFpm (No. constraint = i)
def Constraint31 (model, i):
    return model.x[i,10] == model.mi[i,4] * df_MRFeff.at[WasteType[i],'MRFpm_HDPE']
model.Constraint31 = pyo.Constraint(model.WasteType, rule=Constraint31)


## Purity required for the Recycled HDPE [No. of constraints = 1]
#def Constraint32 (model):
#    return model.x[5,10] >= MinPurity_Recycling * sum(model.x[i,10] for i in model.WasteType)
#model.Constraint32 = pyo.Constraint(rule=Constraint32)


## Recycling rate target for Recycled HDPE [No. of constraints = 1]
#def Constraint33 (model):
#    if df_Generation.at[WasteType[5],'mass [tons]'] > 0:
#        return ( sum(model.x[i,10] for i in model.WasteType) + sum(model.x[i,20] for i in model.WasteType) ) / float (df_Generation.at[WasteType[5],'mass [tons]']) >= MinRecyclingRate
#    else:
#        return pyo.Constraint.Skip
#model.Constraint33 = pyo.Constraint(rule=Constraint33)

#----------

# Production of Recycled MixedPlastics from MRFpm (No. constraint = i)
def Constraint34 (model, i):
    return model.x[i,11] == model.mi[i,4] * df_MRFeff.at[WasteType[i],'MRFpm_MixedPlastics']
model.Constraint34 = pyo.Constraint(model.WasteType, rule=Constraint34)


## Purity required for the Recycled MixedPlastics [No. of constraints = 1]
#def Constraint35 (model):
#    return model.x[6,11] >= MinPurity_Recycling * sum(model.x[i,11] for i in model.WasteType)
#model.Constraint35 = pyo.Constraint(rule=Constraint35)


## Recycling rate target for Recycled MixedPlastics [No. of constraints = 1]
#def Constraint36 (model):
#    if df_Generation.at[WasteType[6],'mass [tons]'] > 0:
#        return ( sum(model.x[i,11] for i in model.WasteType) + sum(model.x[i,21] for i in model.WasteType) ) / float (df_Generation.at[WasteType[6],'mass [tons]']) >= MinRecyclingRate
#    else:
#        return pyo.Constraint.Skip
#model.Constraint36 = pyo.Constraint(rule=Constraint36)

#----------

# Production of Recycled Ferrous metals from MRFpm (No. constraint = i)
def Constraint37 (model, i):
    return model.x[i,12] == model.mi[i,4] * df_MRFeff.at[WasteType[i],'MRFpm_Fmetals']
model.Constraint37 = pyo.Constraint(model.WasteType, rule=Constraint37)


## Purity required for the Recycled Ferrous metals [No. of constraints = 1]
#def Constraint38 (model):
#    return model.x[7,12] >= MinPurity_Recycling * sum(model.x[i,12] for i in model.WasteType)
#model.Constraint38 = pyo.Constraint(rule=Constraint38)


## Recycling rate target for Recycled Ferrous metals [No. of constraints = 1]
#def Constraint39 (model):
#    if df_Generation.at[WasteType[7],'mass [tons]'] > 0:
#        return ( sum(model.x[i,12] for i in model.WasteType) + sum(model.x[i,22] for i in model.WasteType) ) / float (df_Generation.at[WasteType[7],'mass [tons]']) >= MinRecyclingRate
#    else:
#        return pyo.Constraint.Skip
#model.Constraint39 = pyo.Constraint(rule=Constraint39)

#----------

# Production of Recycled Non-ferrous metals from MRFpm (No. constraint = i)
def Constraint40 (model, i):
    return model.x[i,13] == model.mi[i,4] * df_MRFeff.at[WasteType[i],'MRFpm_NonFmetals']
model.Constraint40 = pyo.Constraint(model.WasteType, rule=Constraint40)


## Purity required for the Recycled Non-ferrous metals [No. of constraints = 1]
#def Constraint41 (model):
#    return model.x[8,13] >= MinPurity_Recycling * sum(model.x[i,13] for i in model.WasteType)
#model.Constraint41 = pyo.Constraint(rule=Constraint41)


## Recycling rate target for Recycled Non-ferrous metals [No. of constraints = 1]
#def Constraint42 (model):
#    if df_Generation.at[WasteType[8],'mass [tons]'] > 0:
#        return ( sum(model.x[i,13] for i in model.WasteType) + sum(model.x[i,23] for i in model.WasteType) ) / float (df_Generation.at[WasteType[8],'mass [tons]']) >= MinRecyclingRate
#    else:
#        return pyo.Constraint.Skip
#model.Constraint42 = pyo.Constraint(rule=Constraint42)

#----------

# Outflow of MRFpm (No. of constraints = i)
def Constraint43 (model, i):
    return model.mo[i,4] == model.mi[i,4] \
                        - model.x[i,7] \
                        - model.x[i,8] \
                        - model.x[i,9] \
                        - model.x[i,10] \
                        - model.x[i,11] \
                        - model.x[i,12] \
                        - model.x[i,13]
model.Constraint43 = pyo.Constraint(model.WasteType, rule=Constraint43)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   MRFg   #######
"""

# Input mass to MRFg (No. constraint = i)
def Constraint50 (model, i):
    return model.mi[i,5] == float(df_Generation.at[WasteType[i],'mass [tons]']) \
                        * ( model.r[0] * float(df_Collection_D2D_Factor.at[WasteType[i],'Glass']) \
                           + model.r[1] * float(df_Collection_CS_Factor.at[WasteType[i],'Glass']) \
                        )
model.Constraint50 = pyo.Constraint(model.WasteType, rule=Constraint50)


# Installed capacity increase of MRFg (No. constraint = 1)
def Constraint51 (model):
    return model.sp[5] + df_Processes_Cost.at[ProcessType[5], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,5] for i in model.WasteType) + model.msi[5]
model.Constraint51 = pyo.Constraint(rule=Constraint51)


# Binary variable to activate the CAPEX of MRFg (No. constraint = 2)
#def Constraint51a (model): # 1st part
#    return sum ( model.mi[i,5] for i in model.WasteType) >= 0.0000000001 - M*( 1 - model.bp[5] )
#model.Constraint51a = pyo.Constraint(rule=Constraint51a)
#def Constraint51b(model): # 2nd part
#    return sum ( model.mi[i,5] for i in model.WasteType) <= 0 + M * model.bp[5]
#model.Constraint51b = pyo.Constraint(rule=Constraint51b)

#----------

# Production of Recycled Glass from MRFg (No. constraint = i)
def Constraint52 (model, i):
    return model.x[i,14] == model.mi[i,5] * df_MRFeff.at[WasteType[i],'MRFg_Glass']
model.Constraint52 = pyo.Constraint(model.WasteType, rule=Constraint52)


## Purity required for the Recycled Glass [No. of constraints = 1]
#def Constraint53 (model):
#    return model.x[9,14] >= MinPurity_Recycling * sum(model.x[i,14] for i in model.WasteType)
#model.Constraint53 = pyo.Constraint(rule=Constraint53)


## Recycling rate target for Recycled Glass [No. of constraints = 1]
#def Constraint54 (model):
#    if df_Generation.at[WasteType[9],'mass [tons]'] > 0:
#        return ( sum(model.x[i,14] for i in model.WasteType) + sum(model.x[i,24] for i in model.WasteType) ) / float (df_Generation.at[WasteType[9],'mass [tons]']) >= MinRecyclingRate
#    else:
#        return pyo.Constraint.Skip
#model.Constraint54 = pyo.Constraint(rule=Constraint54)

#----------

# Outflow of MRFg (No. of constraints = i)
def Constraint55 (model, i):
    return model.mo[i,5] == model.mi[i,5] - model.x[i,14]
model.Constraint55 = pyo.Constraint(model.WasteType, rule=Constraint55)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Mechanical Treatment for Biowaste and MixedWaste   #######
"""
## NOTE: Unlike MRF, the constraints in MT (64 to 73) do not allow impurities in the recycling streams.


# Input mass to MTb (No. constraint = i)
def Constraint60 (model, i):
    return model.mi[i,6] == float(df_Generation.at[WasteType[i],'mass [tons]']) \
                        * ( model.r[0] * float(df_Collection_D2D_Factor.at[WasteType[i],'Biowaste']) \
                           + model.r[1] * float(df_Collection_CS_Factor.at[WasteType[i],'Biowaste']) \
                        )
model.Constraint60 = pyo.Constraint(model.WasteType, rule=Constraint60)


# Installed capacity increase of MTb (No. constraint = 1)
def Constraint61 (model):
    return model.sp[6] + df_Processes_Cost.at[ProcessType[6], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,6] for i in model.WasteType) + model.msi[6]
model.Constraint61 = pyo.Constraint(rule=Constraint61)


#----------

# Input mass to MTmw (No. constraint = i) - See Constraint 200 for model.mi0[i,0]
def Constraint62 (model, i):
    return model.mi[i,7] == float(df_Generation.at[WasteType[i],'mass [tons]']) \
                        * ( model.r[0] * float(df_Collection_D2D_Factor.at[WasteType[i],'Mixed_Waste']) \
                           + model.r[1] * float(df_Collection_CS_Factor.at[WasteType[i],'Mixed_Waste']) \
                        ) \
                    - model.mi0[i,0]
model.Constraint62 = pyo.Constraint(model.WasteType, rule=Constraint62)


# Installed capacity increase of MTmw (No. constraint = 1)
def Constraint63 (model):
    return model.sp[7] + df_Processes_Cost.at[ProcessType[7], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,7] for i in model.WasteType) + model.msi[7]
model.Constraint63 = pyo.Constraint(rule=Constraint63)


#----------

# Production of Recycled Paper from MTb and MTmw [No. of constraints = i]
def Constraint64 (model, i):
    if i == 0:
        return model.x[i,15] == model.mi[i,6] * df_MTb_eff.at[WasteType[i],'MT_Paper'] \
                            + model.mi[i,7] * df_MTmw_eff.at[WasteType[i],'MT_Paper']
    else:
        return model.x[i,15] == 0
model.Constraint64 = pyo.Constraint(model.WasteType, rule=Constraint64)


#----------

# Production of Recycled Cardboard from MTb and MTmw [No. of constraints = i]
def Constraint65 (model, i):
    if i == 1:
        return model.x[i,16] == model.mi[i,6] * df_MTb_eff.at[WasteType[i],'MT_Cardboard'] \
                            + model.mi[i,7] * df_MTmw_eff.at[WasteType[i],'MT_Cardboard']
    else:
        return model.x[i,16] == 0
model.Constraint65 = pyo.Constraint(model.WasteType, rule=Constraint65)


#----------

# Production of Recycled Composites from MTb and MTmw [No. of constraints = i]
def Constraint66 (model, i):
    if i == 2:
        return model.x[i,17] == model.mi[i,6] * df_MTb_eff.at[WasteType[i],'MT_Composites'] \
                            + model.mi[i,7] * df_MTmw_eff.at[WasteType[i],'MT_Composites']
    else:
        return model.x[i,17] == 0
model.Constraint66 = pyo.Constraint(model.WasteType, rule=Constraint66)


#----------

# Production of Recycled PE from MTb and MTmw [No. of constraints = i]
def Constraint67 (model, i):
    if i == 3:
        return model.x[i,18] == model.mi[i,6] * df_MTb_eff.at[WasteType[i],'MT_PE'] \
                            + model.mi[i,7] * df_MTmw_eff.at[WasteType[i],'MT_PE']
    else:
        return model.x[i,18] == 0
model.Constraint67 = pyo.Constraint(model.WasteType, rule=Constraint67)


#----------

# Production of Recycled PET from MTb and MTmw [No. of constraints = i]
def Constraint68 (model, i):
    if i == 4:
        return model.x[i,19] == model.mi[i,6] * df_MTb_eff.at[WasteType[i],'MT_PET'] \
                            + model.mi[i,7] * df_MTmw_eff.at[WasteType[i],'MT_PET']
    else:
        return model.x[i,19] == 0
model.Constraint68 = pyo.Constraint(model.WasteType, rule=Constraint68)


#----------

# Production of Recycled HDPE from MTb and MTmw [No. of constraints = i]
def Constraint69 (model, i):
    if i == 5:
        return model.x[i,20] == model.mi[i,6] * df_MTb_eff.at[WasteType[i],'MT_HDPE'] \
                            + model.mi[i,7] * df_MTmw_eff.at[WasteType[i],'MT_HDPE']
    else:
        return model.x[i,20] == 0
model.Constraint69 = pyo.Constraint(model.WasteType, rule=Constraint69)


#----------

# Production of Recycled Mixed_Plastics from MTb and MTmw [No. of constraints = i]
def Constraint70 (model, i):
    if i == 6:
        return model.x[i,21] == model.mi[i,6] * df_MTb_eff.at[WasteType[i],'MT_MixedPlastics'] \
                            + model.mi[i,7] * df_MTmw_eff.at[WasteType[i],'MT_MixedPlastics']
    else:
        return model.x[i,21] == 0
model.Constraint70 = pyo.Constraint(model.WasteType, rule=Constraint70)


#----------

# Production of Recycled Ferrous_Metals from MTb and MTmw [No. of constraints = i]
def Constraint71 (model, i):
    if i == 7:
        return model.x[i,22] == model.mi[i,6] * df_MTb_eff.at[WasteType[i],'MT_Fmetals'] \
                            + model.mi[i,7] * df_MTmw_eff.at[WasteType[i],'MT_Fmetals']
    else:
        return model.x[i,22] == 0
model.Constraint71 = pyo.Constraint(model.WasteType, rule=Constraint71)


#----------

# Production of Recycled NonFerrous_Metals from MTb and MTmw [No. of constraints = i]
def Constraint72 (model, i):
    if i == 8:
        return model.x[i,23] == model.mi[i,6] * df_MTb_eff.at[WasteType[i],'MT_NonFmetals'] \
                            + model.mi[i,7] * df_MTmw_eff.at[WasteType[i],'MT_NonFmetals']
    else:
        return model.x[i,23] == 0
model.Constraint72 = pyo.Constraint(model.WasteType, rule=Constraint72)


#----------

# Production of Recycled Glass from MTb and MTmw [No. of constraints = i]
def Constraint73 (model, i):
    if i == 9:
        return model.x[i,24] == model.mi[i,6] * df_MTb_eff.at[WasteType[i],'MT_Glass'] \
                            + model.mi[i,7] * df_MTmw_eff.at[WasteType[i],'MT_Glass']
    else:
        return model.x[i,24] == 0
model.Constraint73 = pyo.Constraint(model.WasteType, rule=Constraint73)


## Purity required for the Recycled Glass [No. of constraints = 1] 
   ##IT DOES NOT MAKE SENSE TO PUT IT because the column "BaledForRecycling" does not consider pollutants
#def Constraint81 (model):
#    return model.x[9,24] >= MinPurity_Recycling * sum(model.x[i,24] for i in model.WasteType)
#model.Constraint81 = pyo.Constraint(rule=Constraint81)

#----------

# Production of feedstock for thermal treatments from MTb [No. of constraints = i]
def Constraint74 (model, i):
    return model.mo[i,6] == model.mi[i,6] * df_MTb_eff.at[WasteType[i],'MT_TTfeed']
model.Constraint74 = pyo.Constraint(model.WasteType, rule=Constraint74)


# Production of feedstock for thermal treatments from MTmw [No. of constraints = i]
def Constraint75 (model, i):
    return model.mo[i,7] == model.mi[i,7] * df_MTmw_eff.at[WasteType[i],'MT_TTfeed']
model.Constraint75 = pyo.Constraint(model.WasteType, rule=Constraint75)

#----------

# Production of feedstock for biological treatments from MTb [No. of constraints = i]
def Constraint76 (model, i):
    return model.x[i,25] == model.mi[i,6] * df_MTb_eff.at[WasteType[i],'MT_BTfeed']
model.Constraint76 = pyo.Constraint(model.WasteType, rule=Constraint76)

#----------

# Production of feedstock for biological treatments from MTmw [No. of constraints = i]
def Constraint77 (model, i):
    return model.x[i,26] == model.mi[i,7] * df_MTmw_eff.at[WasteType[i],'MT_BTfeed']
model.Constraint77 = pyo.Constraint(model.WasteType, rule=Constraint77)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   MRF and MT CO2 EMISSIONS   #######
"""

def Constraint401 (model):
    return model.Ep[3] == sum( model.mi[i,3] \
                            * float(df_MRF_MT_Env.at['Global warming','Paper_Cardboard']) \
                            for i in model.WasteType )
model.Constraint401 = pyo.Constraint(rule=Constraint401)


def Constraint402 (model):
    return model.Ep[4] == sum( model.mi[i,4] \
                            * float(df_MRF_MT_Env.at['Global warming','Plastic_Metal']) \
                            for i in model.WasteType )
model.Constraint402 = pyo.Constraint(rule=Constraint402)


def Constraint403 (model):
    return model.Ep[5] == sum( model.mi[i,5] \
                            * float(df_MRF_MT_Env.at['Global warming','Glass']) \
                            for i in model.WasteType )
model.Constraint403 = pyo.Constraint(rule=Constraint403)


def Constraint404 (model):
    return model.Ep[6] == sum( model.mi[i,6] \
                            * float(df_MRF_MT_Env.at['Global warming','Biowaste']) \
                            for i in model.WasteType )
model.Constraint404 = pyo.Constraint(rule=Constraint404)


def Constraint405 (model):
    return model.Ep[7] == sum( model.mi[i,7] \
                            * float(df_MRF_MT_Env.at['Global warming','Mixed_Waste']) \
                            for i in model.WasteType )
model.Constraint405 = pyo.Constraint(rule=Constraint405)



#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Anaerobic Digestion   #######
"""

# Input mass to AD, from MTb (that also go to Composting) and MTmw (No. constraint = i)
def Constraint80 (model, i):
    return model.mi[i,8] == model.mi_MTb[i,8] + model.x[i,26]
model.Constraint80 = pyo.Constraint(model.WasteType, rule=Constraint80)


# Installed capacity increase of AD (No. constraint = 1)
def Constraint81 (model):
    return model.sp[8] + df_Processes_Cost.at[ProcessType[8], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,8] for i in model.WasteType) + model.msi[8]
model.Constraint81 = pyo.Constraint(rule=Constraint81)


# Binary variable to activate the CAPEX of AD (No. constraint = 2)
#def Constraint81a (model): # 1st part
#    return sum ( model.mi[i,8] for i in model.WasteType) >= 0.0000000001 - M*( 1 - model.bp[8] )
#model.Constraint81a = pyo.Constraint(rule=Constraint81a)
#def Constraint81b(model): # 2nd part
#    return sum ( model.mi[i,8] for i in model.WasteType) <= 0 + M * model.bp[8]
#model.Constraint81b = pyo.Constraint(rule=Constraint81b)

#----------

# Production of Biogas (single stream) from AD [No. of constraints = 1]
def Constraint82 (model):
    return model.xs[27] - model.msi[9] == \
        sum (model.mi[i,8] * float(df_Generation.at[WasteType[i],'BMP[ton of CH4/ton]']) for i in model.WasteType)
model.Constraint82 = pyo.Constraint(rule=Constraint82)

#----------

# Production of Digestate (single stream) from AD to Composting [No. of constraints = 1]
def Constraint83 (model):
    return model.xs[28] == \
        sum (model.mi[i,8] * float(df_Generation.at[WasteType[i],'BDP[ton of Digestate/ton]']) for i in model.WasteType)
model.Constraint83 = pyo.Constraint(rule=Constraint83)

#----------

# Outflow of AD to Landfill (No. of constraints = 1)
def Constraint84 (model):
    return model.mso[8] == sum ( model.mi[i,8] * max(0, 1 - float(df_Generation.at[WasteType[i],'BMP[ton of CH4/ton]'] + df_Generation.at[WasteType[i],'BDP[ton of Digestate/ton]']) ) for i in model.WasteType)
model.Constraint84 = pyo.Constraint(rule=Constraint84)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Anaerobic Digestion Recovery Electricity   #######
"""

# Input mass to ADRE - see constraint 82


# Installed capacity increase of ADRE (No. constraint = 1)
def Constraint90 (model):
    return model.sp[9] + df_Processes_Cost.at[ProcessType[9], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,9] for i in model.WasteType) + model.msi[9]
model.Constraint90 = pyo.Constraint(rule=Constraint90)

#----------

# Production of Electricity (single stream) from ADRE [No. of constraints = 1]
def Constraint91 (model):
    return model.xs[29] == model.msi[9] * float(ElectGenEff)
model.Constraint91 = pyo.Constraint(rule=Constraint91)


# Production of Heat (single stream) from ADRE [No. of constraints = 1]
def Constraint92 (model):
    return model.xs[30] == model.msi[9] * float(HeatGenEff)
model.Constraint92 = pyo.Constraint(rule=Constraint92)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Composting   #######
"""

# Input mass from MTb to Composting and AD (waste vector) [No. of constraints = i]
def Constraint100 (model, i):
    return model.mi[i,10] == model.x[i,25] - model.mi_MTb[i,8]
model.Constraint100 = pyo.Constraint(model.WasteType, rule=Constraint100)


# Input mass from AD to Composting (single stream) [No. of constraints = 1]
def Constraint101 (model):
    return model.msi[10] == model.xs[28]
model.Constraint101 = pyo.Constraint(rule=Constraint101)


# Installed capacity increase of Composting (No. constraint = 1)
def Constraint102 (model):
    return model.sp[10] + df_Processes_Cost.at[ProcessType[10], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,10] for i in model.WasteType) + model.msi[10]
model.Constraint102 = pyo.Constraint(rule=Constraint102)

#----------

## Maximum C:N rate [No. of constraints = 1]
#def Constraint103 (model, i):
#    return sum(model.mi[i,10] * float(df_Generation.at[WasteType[i],'BOC[% of VS]']) for i in model.WasteType) <= 30 * sum(model.mi[i,10] * float(df_Generation.at[WasteType[i],'BON[% of VS]']) for i in model.WasteType)
#model.Constraint103 = pyo.Constraint(model.WasteType, rule=Constraint103)
#
#
## Minimum C:N rate [No. of constraints = 1]
#def Constraint104 (model, i):
#    return sum(model.mi[i,10] * float(df_Generation.at[WasteType[i],'BOC[% of VS]']) for i in model.WasteType) >= 20 * sum(model.mi[i,10] * float(df_Generation.at[WasteType[i],'BON[% of VS]']) for i in model.WasteType)
#model.Constraint104 = pyo.Constraint(model.WasteType, rule=Constraint104)

#----------

# Production of Compost from Composting [No. of constraints = 1]
def Constraint105 (model):
    return model.xs[31] == sum( model.mi[i,10] * CompostingRate for i in model.WasteType) \
                            + model.msi[10] * CompostingRate
model.Constraint105 = pyo.Constraint(rule=Constraint105)


# Outflow of Composting to Landfill [No. of constraints = 1]
def Constraint106 (model):
    return model.mso[10] == sum( model.mi[i,10] * (1 - CompostingRate) for i in model.WasteType) \
                            + model.msi[10] * (1 - CompostingRate)
model.Constraint106 = pyo.Constraint(rule=Constraint106)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Incineration   #######
"""

# Input mass to Incineration, Gasification, Pyrolysis, or Landfill (No. constraint = i)
def Constraint110 (model, i):
    return model.mi[i,11] + model.mi[i,12] + model.mi[i,13] + model.mi_LvsTT[i,0] == model.mo[i,3] + model.mo[i,4] + model.mo[i,5] + model.mo[i,6] + model.mo[i,7]
model.Constraint110 = pyo.Constraint(model.WasteType, rule=Constraint110)


# Installed capacity increase of Incineration (No. constraint = 1)
def Constraint111 (model):
    return model.sp[11] + df_Processes_Cost.at[ProcessType[11], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,11] for i in model.WasteType) + model.msi[11]
model.Constraint111 = pyo.Constraint(rule=Constraint111)

#----------

#def LowCaloricValue(HCV, moisture, ash, carbon):
#    Hawf = float (HCV) / float ( (1-ash) * 2445 ) # Ash and water free calorific value (Hawf)
#    Hinf = float(Hawf) * carbon - 2445 * moisture # Lower calorific value (Hinf)


# Production of Electricity from Incineration [No. of constraints = 1]
def Constraint112 (model):
    return model.xs[32] == sum( model.mi[i,11] * eff_Incineration_Electricity for i in model.WasteType)
model.Constraint112 = pyo.Constraint(rule=Constraint112)


# Production of Heat from Incineration [No. of constraints = 1]
def Constraint113 (model):
    return model.xs[33] == sum( model.mi[i,11] * (1 - eff_Incineration_Electricity - eff_Incineration_Slag) for i in model.WasteType)
model.Constraint113 = pyo.Constraint(rule=Constraint113)


# Outflow from Incineration to Landfill (Production of Slag) [No. of constraints = 1]
def Constraint114 (model):
    return model.mso[11] == sum( model.mi[i,11] * eff_Incineration_Slag for i in model.WasteType)
model.Constraint114 = pyo.Constraint(rule=Constraint114)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Gasification   #######
"""

# Input mass to Gasification - see constraint 110


# Installed capacity increase of Gasification (No. constraint = 1)
def Constraint120 (model):
    return model.sp[12] + df_Processes_Cost.at[ProcessType[12], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,12] for i in model.WasteType) + model.msi[12]
model.Constraint120 = pyo.Constraint(rule=Constraint120)

#----------

# Production of Syngas from Gasification [No. of constraints = 1]
def Constraint121 (model):
    return model.xs[34] == sum( model.mi[i,12] * eff_Gasification_Syngas for i in model.WasteType) \
                            - model.moG_sg[12]
model.Constraint121 = pyo.Constraint(rule=Constraint121)

#----------

# Outflow from Gasification to Landfill [No. of constraints = 1]
def Constraint122 (model):
    return model.mso[12] == sum( model.mi[i,12] * (1 - eff_Gasification_Syngas) for i in model.WasteType)
model.Constraint122 = pyo.Constraint(rule=Constraint122)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Pyrolysis   #######
"""

# Input mass to Pyrolysis - see constraint 110


# Installed capacity increase of Pyrolysis (No. constraint = 1)
def Constraint130 (model):
    return model.sp[13] + df_Processes_Cost.at[ProcessType[13], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,13] for i in model.WasteType) + model.msi[13]
model.Constraint130 = pyo.Constraint(rule=Constraint130)

#----------

# Production of Syngas from Pyrolysis [No. of constraints = 1]
def Constraint131 (model):
    return model.xs[35] == sum( model.mi[i,13] * eff_Pyrolysis_Syngas for i in model.WasteType) \
                            - model.moP_sg[13]
model.Constraint131 = pyo.Constraint(rule=Constraint131)


# Production of BioOil from Pyrolysis [No. of constraints = 1]
def Constraint132 (model):
    return model.xs[36] == sum( model.mi[i,13] * eff_Pyrolysis_Bio_oil for i in model.WasteType)
model.Constraint132 = pyo.Constraint(rule=Constraint132)

#----------

# Outflow from Pyrolysis to Landfill [No. of constraints = 1]
def Constraint133 (model):
    return model.mso[13] == sum( model.mi[i,13] * (1 - eff_Pyrolysis_Syngas - eff_Pyrolysis_Bio_oil) for i in model.WasteType)
model.Constraint133 = pyo.Constraint(rule=Constraint133)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Shift_Reaction   #######
"""

# Input mass to Shift Reaction + Fisher-Tropsch (for lh and b) + Methanol Syntesis + Alcohol Syntesis + Fermentation, from Gasification + Pyrolysis (No. constraint = 1)
def Constraint140 (model):
    return model.msi[14] + model.msi[15] + model.msi[16] + model.msi[17] + model.msi[18] + model.msi[19] == model.moG_sg[12] + model.moP_sg[13]
model.Constraint140 = pyo.Constraint(rule=Constraint140)


# Installed capacity increase of Shift_Reaction (No. constraint = 1)
def Constraint141 (model):
    return model.sp[14] + df_Processes_Cost.at[ProcessType[14], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,14] for i in model.WasteType) + model.msi[14]
model.Constraint141 = pyo.Constraint(rule=Constraint141)

#----------

# Production of Hydrogen from Shift_Reaction [No. of constraints = 1]
def Constraint142 (model):
    return model.xs[37] == model.msi[14] * eff_Shift_Reaction_To_Hydrogen
model.Constraint142 = pyo.Constraint(rule=Constraint142)

#----------

# Outflow from Shift_Reaction to Landfill [No. of constraints = 1]
def Constraint143 (model):
    return model.mso[14] == model.msi[14] * (1 - eff_Shift_Reaction_To_Hydrogen)
model.Constraint143 = pyo.Constraint(rule=Constraint143)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Fischer_Tropsch for LightHydrocarbons   #######
"""

# Input mass to Fisher-Tropsch for LightHydrocarbons - See constraint 140


# Installed capacity increase of Fisher-Tropsch for LightHydrocarbons (No. constraint = 1)
def Constraint150 (model):
    return model.sp[15] + df_Processes_Cost.at[ProcessType[15], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,15] for i in model.WasteType) + model.msi[15]
model.Constraint150 = pyo.Constraint(rule=Constraint150)

#----------

# Production of Light_hydrocarbons from Fischer_Tropsch [No. of constraints = 1]
def Constraint151 (model):
    return model.xs[38] == model.msi[15] * eff_Fischer_Tropsch_To_Light_hydrocarbons
model.Constraint151 = pyo.Constraint(rule=Constraint151)

#----------

# Outflow from Fischer_Tropsch for LightHydrocarbons to Landfill [No. of constraints = 1]
def Constraint152 (model):
    return model.mso[15] == model.msi[15] * (1 - eff_Fischer_Tropsch_To_Light_hydrocarbons)
model.Constraint152 = pyo.Constraint(rule=Constraint152)


#----------

"""
#######   Fischer_Tropsch for Biodiesel   #######
"""

# Input mass to Fisher-Tropsch for Biodiesel - See constraint 140


# Installed capacity increase of Fisher-Tropsch for Biodiesel (No. constraint = 1)
def Constraint153 (model):
    return model.sp[16] + df_Processes_Cost.at[ProcessType[16], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,16] for i in model.WasteType) + model.msi[16]
model.Constraint153 = pyo.Constraint(rule=Constraint153)

#----------

# Production of Biodiesel from Fischer_Tropsch [No. of constraints = 1]
def Constraint154 (model):
    return model.xs[39] == model.msi[16] * eff_Fischer_Tropsch_To_Biodiesel
model.Constraint154 = pyo.Constraint(rule=Constraint154)

#----------

# Outflow from Fischer_Tropsch for Biodiesel to Landfill [No. of constraints = 1]
def Constraint155 (model):
    return model.mso[16] == model.msi[16] * (1 - eff_Fischer_Tropsch_To_Biodiesel)
model.Constraint155 = pyo.Constraint(rule=Constraint155)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Methanol_Synthesis   #######
"""

# Input mass to Methanol_Synthesis - See constraint 140


# Installed capacity increase of Methanol_Synthesis (No. constraint = 1)
def Constraint160 (model):
    return model.sp[17] + df_Processes_Cost.at[ProcessType[17], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,17] for i in model.WasteType) + model.msi[17]
model.Constraint160 = pyo.Constraint(rule=Constraint160)


# Binary variable to activate the CAPEX of Methanol_Synthesis (No. constraint = 2)
#def Constraint160a (model): # 1st part
#    return model.msi[17] >= 0.0000000001 - M*( 1 - model.bp[17] )
#model.Constraint160a = pyo.Constraint(rule=Constraint160a)
#def Constraint160b(model): # 2nd part
#    return model.msi[17] <= 0 + M * model.bp[17]
#model.Constraint160b = pyo.Constraint(rule=Constraint160b)

#----------

# Production of Methanol from Methanol_Synthesis [No. of constraints = 1]
def Constraint161 (model):
    return model.xs[40] == model.msi[17] * eff_Methanol_Synthesis
model.Constraint161 = pyo.Constraint(rule=Constraint161)

#----------

# Outflow from Methanol_Synthesis to Landfill [No. of constraints = 1]
def Constraint162 (model):
    return model.mso[17] == model.msi[17] * (1 - eff_Methanol_Synthesis)
model.Constraint162 = pyo.Constraint(rule=Constraint162)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Alcohol_Synthesis   #######
"""

# Input mass to Alcohol_Synthesis - See constraint 140


# Installed capacity increase of Alcohol_Synthesis (No. constraint = 1)
def Constraint170 (model):
    return model.sp[18] + df_Processes_Cost.at[ProcessType[18], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,18] for i in model.WasteType) + model.msi[18]
model.Constraint170 = pyo.Constraint(rule=Constraint170)


# Binary variable to activate the CAPEX of Alcohol_Synthesis (No. constraint = 2)
#def Constraint170a (model): # 1st part
#    return model.msi[18] >= 0.0000000001 - M*( 1 - model.bp[18] )
#model.Constraint170a = pyo.Constraint(rule=Constraint170a)
#def Constraint170b(model): # 2nd part
#    return model.msi[18] <= 0 + M * model.bp[18]
#model.Constraint170b = pyo.Constraint(rule=Constraint170b)

#----------

# Production of Ethanol from Alcohol_Synthesis [No. of constraints = 1]
def Constraint171 (model):
    return model.xs[41] == model.msi[18] * eff_Ethanol_Synthesis
model.Constraint171 = pyo.Constraint(rule=Constraint171)

#----------

# Outflow from Alcohol_Synthesis to Landfill [No. of constraints = 1]
def Constraint172 (model):
    return model.mso[18] == model.msi[18] * (1 - eff_Ethanol_Synthesis)
model.Constraint172 = pyo.Constraint(rule=Constraint172)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Fermentation   #######
"""

# Input mass to Fermentation - See constraint 140


# Installed capacity increase of Fermentation (No. constraint = 1)
def Constraint180 (model):
    return model.sp[19] + df_Processes_Cost.at[ProcessType[19], 'Existant installed capacity [ton]'] >= \
            sum(model.mi[i,19] for i in model.WasteType) + model.msi[19]
model.Constraint180 = pyo.Constraint(rule=Constraint180)


# Binary variable to activate the CAPEX of Fermentation (No. constraint = 2)
#def Constraint180a (model): # 1st part
#    return model.msi[19] >= 0.0000000001 - M*( 1 - model.bp[19] )
#model.Constraint180a = pyo.Constraint(rule=Constraint180a)
#def Constraint180b(model): # 2nd part
#    return model.msi[19] <= 0 + M * model.bp[19]
#model.Constraint180b = pyo.Constraint(rule=Constraint180b)

#----------

# Production of Ethanol from Fermentation [No. of constraints = i]
def Constraint181 (model):
    return model.xs[42] == model.msi[19] * eff_Ethanol_Fermentation
model.Constraint181 = pyo.Constraint(rule=Constraint181)

#----------

# Outflow from Fermentation to Landfill [No. of constraints = i]
def Constraint182 (model):
    return model.mso[19] == model.msi[19] * (1 - eff_Ethanol_Fermentation)
model.Constraint182 = pyo.Constraint(rule=Constraint182)


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Temporal contraints to be able to run the model   #######
"""

# Not aplicable coproducts model.x[i,n]
def Constraint300 (model, i, n):
    if n == 5 or n == 6:
        return pyo.Constraint.Skip
    elif n == 7 or n == 8 or n == 9 or n == 10 or n == 11 or n == 12 or n == 13:
        return pyo.Constraint.Skip
    elif n == 14:
        return pyo.Constraint.Skip
    elif n == 15 or n == 16:
        return pyo.Constraint.Skip
    elif n == 17 or n == 18 or n == 19 or n == 20 or n == 21 or n == 22 or n == 23:
        return pyo.Constraint.Skip
    elif n == 24:
        return pyo.Constraint.Skip
    elif n == 25 or n == 26:
        return pyo.Constraint.Skip
    else:
        return model.x[i,n] == 0
model.Constraint300 = pyo.Constraint(model.WasteType, model.CoproductType, rule=Constraint300)


# Not aplicable coproducts model.x_dummy[i,n]
def Constraint300b (model, i, n):
##    return model.x_dummy[i,n] == 0
    if n == 0 or n == 2:
        return pyo.Constraint.Skip
    else:
        return model.x_dummy[i,n] == 0
model.Constraint300b = pyo.Constraint(model.WasteType, model.CoproductType, rule=Constraint300b)



# Not aplicable coproducts model.xs[n]
def Constraint301 (model, n):
#    return model.xs[n] == 0
    if n == 0 or n == 1:
        return pyo.Constraint.Skip
    elif n == 2:
        return pyo.Constraint.Skip
    elif n == 3 or n == 4:
        return pyo.Constraint.Skip
    elif n == 27 or n == 28:
        return pyo.Constraint.Skip
    elif n == 29 or n == 30:
        return pyo.Constraint.Skip
    elif n == 31:
        return pyo.Constraint.Skip
    elif n == 32 or n == 33:
        return pyo.Constraint.Skip
    elif n == 34:
        return pyo.Constraint.Skip
    elif n == 35 or n == 36:
        return pyo.Constraint.Skip
    elif n == 37 or n == 38 or n == 39 or n == 40 or n == 41 or n == 42:
        return pyo.Constraint.Skip
    else:
        return model.xs[n] == 0
model.Constraint301 = pyo.Constraint(model.CoproductType, rule=Constraint301)


# Not aplicable coproducts model.xs_dummy[n]
def Constraint301b (model, n):
##    return model.x_dummy[i,n] == 0
    if n == 0 or n == 2:
        return pyo.Constraint.Skip
    else:
        return model.xs_dummy[n] == 0
model.Constraint301b = pyo.Constraint(model.CoproductType, rule=Constraint301b)



# Not aplicable variable for input of the process model.mi[i,l]
def Constraint303 (model, i, l):
    if l == 0:
        return pyo.Constraint.Skip
    elif l == 3 or l == 4 or l == 5:
        return pyo.Constraint.Skip
    elif l == 6 or l == 7:
        return pyo.Constraint.Skip
    elif l == 8:
        return pyo.Constraint.Skip
    elif l == 10:
        return pyo.Constraint.Skip
    elif l == 11:
        return pyo.Constraint.Skip
    elif l == 12:
        return pyo.Constraint.Skip
    elif l == 13:
        return pyo.Constraint.Skip
    else:
        return model.mi[i,l] == 0
model.Constraint303 = pyo.Constraint(model.WasteType, model.ProcessType, rule=Constraint303)


# Not aplicable variable for input of the process model.mi_dummy[i,l]
def Constraint303b (model, i, l):
    if l == 1 or l == 2:
        return pyo.Constraint.Skip
    else:
        return model.mi_dummy[i,l] == 0
model.Constraint303b = pyo.Constraint(model.WasteType, model.ProcessType, rule=Constraint303b)



# Not aplicable variable for single output of the process model.msi[l]
def Constraint304 (model, l):
    if l == 0 or l == 1 or l == 2:
        return pyo.Constraint.Skip
    elif l == 9:
        return pyo.Constraint.Skip
    elif l == 10:
        return pyo.Constraint.Skip
    elif l == 14:
        return pyo.Constraint.Skip
    elif l == 15 or l == 16:
        return pyo.Constraint.Skip
    elif l == 17:
        return pyo.Constraint.Skip
    elif l == 18:
        return pyo.Constraint.Skip
    elif l == 19:
        return pyo.Constraint.Skip
    else:
        return model.msi[l] == 0
model.Constraint304 = pyo.Constraint(model.ProcessType, rule=Constraint304)


# Not aplicable variable for input of the process model.msi_dummy[l]
def Constraint304b (model, l):
    if l == 1 or l == 2:
        return pyo.Constraint.Skip
    else:
        return model.msi_dummy[l] == 0
model.Constraint304b = pyo.Constraint(model.ProcessType, rule=Constraint304b)



# Not aplicable variable for output of the process model.mo[i,l]
def Constraint305 (model, i, l):
    if l == 3 or l == 4 or l == 5:
        return pyo.Constraint.Skip
    if l == 6 or l == 7:
        return pyo.Constraint.Skip
    else:
        return model.mo[i,l] == 0
model.Constraint305 = pyo.Constraint(model.WasteType, model.ProcessType, rule=Constraint305)


# Not aplicable variable for single output of the process model.mso[l]
def Constraint306 (model, l):
    if l == 8:
        return pyo.Constraint.Skip
    elif l == 10:
        return pyo.Constraint.Skip
    elif l == 11:
        return pyo.Constraint.Skip
    elif l == 12:
        return pyo.Constraint.Skip
    elif l == 13:
        return pyo.Constraint.Skip
    elif l == 14:
        return pyo.Constraint.Skip
    elif l == 15 or l == 16:
        return pyo.Constraint.Skip
    elif l == 17:
        return pyo.Constraint.Skip
    elif l == 18:
        return pyo.Constraint.Skip
    elif l == 19:
        return pyo.Constraint.Skip
    else:
        return model.mso[l] == 0
model.Constraint306 = pyo.Constraint(model.ProcessType, rule=Constraint306)


# Not aplicable environmetal variables model.Ep[l]
def Constraint308 (model, l):
#    return model.Ep[l] == 0
    if l == 0 or l == 1 or l == 2:
        return pyo.Constraint.Skip
    elif l == 3 or l == 4 or l == 5 or l == 6 or l == 7:
        return pyo.Constraint.Skip
##    elif l == 10:
##        return pyo.Constraint.Skip
##    elif l == 11:
##        return pyo.Constraint.Skip
##    elif l == 12:
##        return pyo.Constraint.Skip
##    elif l == 13:
##        return pyo.Constraint.Skip
##    elif l == 14:
##        return pyo.Constraint.Skip
##    elif l == 15 or l == 16:
##        return pyo.Constraint.Skip
##    elif l == 17:
##        return pyo.Constraint.Skip
##    elif l == 18:
##        return pyo.Constraint.Skip
##    elif l == 19:
##        return pyo.Constraint.Skip
    else:
        return model.Ep[l] == 0
model.Constraint308 = pyo.Constraint(model.ProcessType, rule=Constraint308)


#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------


"""
###############     S O L V I N G    T H E    M O D E L     ###############
"""

opt = SolverFactory('cplex')
results = opt.solve(model, keepfiles=True, tee=True)
print(dir(results))
print(results.read)

results.write(num=1)
print(model.Solution.expr())


#%%-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
###############     S A V I N G    T H E    R E S U L T S    I N    E X C E L     ###############
"""

#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Recycling Rates    #######
"""

index = WasteType.copy()
index.append('TOTAL_Recyclables')

## Calculating the Recycling rates
RecyclingRate = []
massIn = 0
massOut = 0
for i in range (len(WasteType)):
    if i == 0 and df_Generation.at[WasteType[0],'mass [tons]'] > 0:
        RecyclingRate.append( ( sum(model.x[i,5].value for i in model.WasteType) + sum(model.x[i,15].value for i in model.WasteType) ) / float (df_Generation.at[WasteType[0],'mass [tons]']) )
        massIn = massIn + float (df_Generation.at[WasteType[0],'mass [tons]'])
        massOut = massOut + sum(model.x[i,5].value for i in model.WasteType) + sum(model.x[i,15].value for i in model.WasteType)

    elif i == 1 and df_Generation.at[WasteType[1],'mass [tons]'] > 0:
        RecyclingRate.append( ( sum(model.x[i,6].value for i in model.WasteType) + sum(model.x[i,16].value for i in model.WasteType) ) / float (df_Generation.at[WasteType[1],'mass [tons]']) )
        massIn = massIn + float (df_Generation.at[WasteType[1],'mass [tons]'])
        massOut = massOut + sum(model.x[i,6].value for i in model.WasteType) + sum(model.x[i,16].value for i in model.WasteType)

    elif i == 2 and df_Generation.at[WasteType[2],'mass [tons]'] > 0:
        RecyclingRate.append( ( sum(model.x[i,7].value for i in model.WasteType) + sum(model.x[i,17].value for i in model.WasteType) ) / float (df_Generation.at[WasteType[2],'mass [tons]']) )
        massIn = massIn + float (df_Generation.at[WasteType[2],'mass [tons]'])
        massOut = massOut + sum(model.x[i,7].value for i in model.WasteType) + sum(model.x[i,17].value for i in model.WasteType)

    elif i == 3 and df_Generation.at[WasteType[3],'mass [tons]'] > 0:
        RecyclingRate.append( ( sum(model.x[i,8].value for i in model.WasteType) + sum(model.x[i,18].value for i in model.WasteType) ) / float (df_Generation.at[WasteType[3],'mass [tons]']) )
        massIn = massIn + float (df_Generation.at[WasteType[3],'mass [tons]'])
        massOut = massOut + sum(model.x[i,8].value for i in model.WasteType) + sum(model.x[i,18].value for i in model.WasteType)

    elif i == 4 and df_Generation.at[WasteType[4],'mass [tons]'] > 0:
        RecyclingRate.append( ( sum(model.x[i,9].value for i in model.WasteType) + sum(model.x[i,19].value for i in model.WasteType) ) / float (df_Generation.at[WasteType[4],'mass [tons]']) )
        massIn = massIn + float (df_Generation.at[WasteType[4],'mass [tons]'])
        massOut = massOut + sum(model.x[i,9].value for i in model.WasteType) + sum(model.x[i,19].value for i in model.WasteType)

    elif i == 5 and df_Generation.at[WasteType[5],'mass [tons]'] > 0:
        RecyclingRate.append( ( sum(model.x[i,10].value for i in model.WasteType) + sum(model.x[i,20].value for i in model.WasteType) ) / float (df_Generation.at[WasteType[5],'mass [tons]']) )
        massIn = massIn + float (df_Generation.at[WasteType[5],'mass [tons]'])
        massOut = massOut + sum(model.x[i,10].value for i in model.WasteType) + sum(model.x[i,20].value for i in model.WasteType)

    elif i == 6 and df_Generation.at[WasteType[6],'mass [tons]'] > 0:
        RecyclingRate.append( ( sum(model.x[i,11].value for i in model.WasteType) + sum(model.x[i,21].value for i in model.WasteType) ) / float (df_Generation.at[WasteType[6],'mass [tons]']) )
        massIn = massIn + float (df_Generation.at[WasteType[6],'mass [tons]'])
        massOut = massOut + sum(model.x[i,11].value for i in model.WasteType) + sum(model.x[i,21].value for i in model.WasteType)

    elif i == 7 and df_Generation.at[WasteType[7],'mass [tons]'] > 0:
        RecyclingRate.append( ( sum(model.x[i,12].value for i in model.WasteType) + sum(model.x[i,22].value for i in model.WasteType) ) / float (df_Generation.at[WasteType[7],'mass [tons]']) )
        massIn = massIn + float (df_Generation.at[WasteType[7],'mass [tons]'])
        massOut = massOut + sum(model.x[i,12].value for i in model.WasteType) + sum(model.x[i,22].value for i in model.WasteType)

    elif i == 8 and df_Generation.at[WasteType[8],'mass [tons]'] > 0:
        RecyclingRate.append( ( sum(model.x[i,13].value for i in model.WasteType) + sum(model.x[i,23].value for i in model.WasteType) ) / float (df_Generation.at[WasteType[8],'mass [tons]']) )
        massIn = massIn + float (df_Generation.at[WasteType[8],'mass [tons]'])
        massOut = massOut + sum(model.x[i,13].value for i in model.WasteType) + sum(model.x[i,23].value for i in model.WasteType)

    elif i == 9 and df_Generation.at[WasteType[9],'mass [tons]'] > 0:
        RecyclingRate.append( ( sum(model.x[i,14].value for i in model.WasteType) + sum(model.x[i,24].value for i in model.WasteType) ) / float (df_Generation.at[WasteType[9],'mass [tons]']) )
        massIn = massIn + float (df_Generation.at[WasteType[9],'mass [tons]'])
        massOut = massOut + sum(model.x[i,14].value for i in model.WasteType) + sum(model.x[i,24].value for i in model.WasteType)

    elif i == 10 or i == 11 or i == 12 or i == 13 or i == 14:
        RecyclingRate.append('NA')

    else:
        RecyclingRate.append('NoWaste')

RecyclingRate.append(massOut/massIn)


data = {'Recycling rate [-]': RecyclingRate}

df_Rr = pd.DataFrame(data=data, index=index)
del i, data, RecyclingRate, index


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Purity Rates    #######
"""

index1 = save_CoproductName.copy()
index = [varx[1] for varx in enumerate(index1) if varx[0] in [5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24]]

index1_Processes = save_Process.copy()
indexProcesses = [varx[1] for varx in enumerate(index1_Processes) if varx[0] in [5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24]]


## Calculating the Purity required for each value-added product of recycling
PurityRate = []
for n in range (len(CoproductType)):
#    if n != 0 or n != 1 or n != 2 or n != 3 or n != 4 or n != 25 or n != 26 or n != 27 or n != 28 or n != 29 or n != 30 or n != 31 or n != 32 or n != 33 or n != 34 or n != 35 or n != 36 or n != 37 or n != 38 or n != 39 or n != 40 or n != 41  n != 42:
    if n == 5 and sum(model.x[i,5].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[0,5].value / sum(model.x[i,5].value for i in model.WasteType) )
    elif n == 6 and sum(model.x[i,6].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[1,6].value / sum(model.x[i,6].value for i in model.WasteType) )
    elif n == 7 and sum(model.x[i,7].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[2,7].value / sum(model.x[i,7].value for i in model.WasteType) )
    elif n == 8 and sum(model.x[i,8].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[3,8].value / sum(model.x[i,8].value for i in model.WasteType) )
    elif n == 9 and sum(model.x[i,9].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[4,9].value / sum(model.x[i,9].value for i in model.WasteType) )
    elif n == 10 and sum(model.x[i,10].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[5,10].value / sum(model.x[i,10].value for i in model.WasteType) )
    elif n == 11 and sum(model.x[i,11].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[6,11].value / sum(model.x[i,11].value for i in model.WasteType) )
    elif n == 12 and sum(model.x[i,12].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[7,12].value / sum(model.x[i,12].value for i in model.WasteType) )
    elif n == 13 and sum(model.x[i,13].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[8,13].value / sum(model.x[i,13].value for i in model.WasteType) )
    elif n == 14 and sum(model.x[i,14].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[9,14].value / sum(model.x[i,14].value for i in model.WasteType) )
    elif n == 15 and sum(model.x[i,15].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[0,15].value / sum(model.x[i,15].value for i in model.WasteType) )
    elif n == 16 and sum(model.x[i,16].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[1,16].value / sum(model.x[i,16].value for i in model.WasteType) )
    elif n == 17 and sum(model.x[i,17].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[2,17].value / sum(model.x[i,17].value for i in model.WasteType) )
    elif n == 18 and sum(model.x[i,18].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[3,18].value / sum(model.x[i,18].value for i in model.WasteType) )
    elif n == 19 and sum(model.x[i,19].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[4,19].value / sum(model.x[i,19].value for i in model.WasteType) )
    elif n == 20 and sum(model.x[i,20].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[5,20].value / sum(model.x[i,20].value for i in model.WasteType) )
    elif n == 21 and sum(model.x[i,21].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[6,21].value / sum(model.x[i,21].value for i in model.WasteType) )
    elif n == 22 and sum(model.x[i,22].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[7,22].value / sum(model.x[i,22].value for i in model.WasteType) )
    elif n == 23 and sum(model.x[i,23].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[8,23].value / sum(model.x[i,23].value for i in model.WasteType) )
    elif n == 24 and sum(model.x[i,24].value for i in model.WasteType) > 0:
        PurityRate.append( model.x[9,24].value / sum(model.x[i,24].value for i in model.WasteType) )
    elif n == 0 or n == 1 or n == 2 or n == 3 or n == 4 or n == 25 or n == 26 or n == 27 or n == 28 or n == 29 or n == 30 or n == 31 or n == 32 or n == 33 or n == 34 or n == 35 or n == 36 or n == 37 or n == 38 or n == 39 or n == 40 or n == 41 or n == 42:
        pass
    else:
        PurityRate.append('NoGeneration')


data = {'Purity rate [-]': PurityRate,
        'Process': indexProcesses}

df_Pr = pd.DataFrame(data=data, index=index)
del n, data, PurityRate, index1, index, index1_Processes, indexProcesses


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Collection   #######
"""

index = ['Door2Door', 'Curside']
Binary = []
Share = []
for j in range (len(CollectionType)):
    Share.append(model.r[j].value)
    if sum ( float( df_Generation.at[WasteType[i], 'mass [tons]'] ) for i in model.WasteType) > 0:
        if model.r[j].value > 0:
            Binary.append("Yes")
        else:
            Binary.append("No")
data1 = {'Active?': Binary,
        'Share': Share,
        }


for k in range (len(BinType)):
    exec("CurrentInstCap_%s = []" %(BinType[k]))
for j in range (len(CollectionType)):
    for k in range (len(BinType)):
        exec("CurrentInstCap_%s.append(df_Collection_Cost_%s.at['Existant installed capacity [ton]', BinType[k]])" %(BinType[k],CollectionType[j]))
data2 = {'Mixed Waste' : CurrentInstCap_Mixed_Waste,
         'Paper-Cardboard' : CurrentInstCap_Paper_Cardboard,
         'Plastic-Metal' : CurrentInstCap_Plastic_Metal,
         'Biowaste' : CurrentInstCap_Biowaste,
         'Glass' : CurrentInstCap_Glass,
         }
for k in range (len(BinType)):
    exec("del(CurrentInstCap_%s)" %(BinType[k]))


for k in range (len(BinType)):
    exec("ExtraInstCap_%s = []" %(BinType[k]))
for j in range (len(CollectionType)):
    for k in range (len(BinType)):
        exec("ExtraInstCap_%s.append(model.sc[j,k].value)" %(BinType[k]))
data3 = {'Mixed Waste' : ExtraInstCap_Mixed_Waste,
         'Paper-Cardboard' : ExtraInstCap_Paper_Cardboard,
         'Plastic-Metal' : ExtraInstCap_Plastic_Metal,
         'Biowaste' : ExtraInstCap_Biowaste,
         'Glass' : ExtraInstCap_Glass,
         }
for k in range (len(BinType)):
    exec("del(ExtraInstCap_%s)" %(BinType[k]))


for k in range (len(BinType)):
    exec("MassIn_%s = []" %(BinType[k]))
for j in range (len(CollectionType)):
    for k in range (len(BinType)):
        exec("MassIn_%s.append(model.mp[j,k].value)" %(BinType[k]))
data4 = {'Mixed Waste' : MassIn_Mixed_Waste,
         'Paper-Cardboard' : MassIn_Paper_Cardboard,
         'Plastic-Metal' : MassIn_Plastic_Metal,
         'Biowaste' : MassIn_Biowaste,
         'Glass' : MassIn_Glass,
         }
for k in range (len(BinType)):
    exec("del(MassIn_%s)" %(BinType[k]))


df_c_share = pd.DataFrame(data=data1, index=index)
df_c_Inst = pd.DataFrame(data=data2, index=index)
df_c_ExtraInst = pd.DataFrame(data=data3, index=index) 
df_c_MassIn = pd.DataFrame(data=data4, index=index) 

del j, k, data1, data2, data3, data4, Share, Binary, index


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Value-added products   #######
"""

CoproductProd_Total = []
for n in range (len(CoproductType)):
    CoproductProd_Total.append( sum(model.x[i,n].value for i in model.WasteType) + model.xs[n].value )
data = {'Value-added product': save_CoproductName,
        'Total Qty.': CoproductProd_Total,
        'Units': save_Units,
        'Process': save_Process}

df_x = pd.DataFrame(data=data, index=save_CoproductName)
del n, data, CoproductProd_Total
#del save_CoproductName, save_Process, save_Units


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   TechUse   #######
"""

Binary = []
CurrentInstCap = []
ExtraInstCap = []
MassIn = []
MassOut = []
for l in range (len(ProcessType)):
    CurrentInstCap.append(df_Processes_Cost.at[ProcessType[l], 'Existant installed capacity [ton]'])
    ExtraInstCap.append(model.sp[l].value)
    MassIn.append(sum(model.mi[i,l].value for i in model.WasteType) + model.msi[l].value)
    MassOut.append(-sum(model.mo[i,l].value for i in model.WasteType) - model.mso[l].value)
    if sum(model.mi[i,l].value for i in model.WasteType) + model.msi[l].value > 0:
        Binary.append("Yes")
    else:
        Binary.append("No")
data = {'Active?': Binary,
        'Current installed capacity [ton]':CurrentInstCap,
        'Extra required installed capacity [ton]':ExtraInstCap,
        'MassIn [ton]': MassIn,
        'MassOut [ton]': MassOut}

df_m = pd.DataFrame(data=data, index=ProcessType)
del l, data, MassIn, MassOut, Binary


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   TechUse_Detail   #######
"""

WasteTypeToSave = WasteType.copy()
WasteTypeToSave.append('SingleStream')


#data = []
for l in range (len(ProcessType)):
    exec("%s_MassIn = []" %(ProcessType[l]))
    exec("%s_MassOut = []" %(ProcessType[l]))
    for i in range (len(WasteType)):
        exec("%s_MassIn.append(model.mi[i,l].value)" %(ProcessType[l]))
        exec("%s_MassOut.append(-model.mo[i,l].value)" %(ProcessType[l]))
    exec("%s_MassIn.append(model.msi[l].value)" %(ProcessType[l]))
    exec("%s_MassOut.append(-model.mso[l].value)" %(ProcessType[l]))


data2 = {'Landfill_MassIn': Landfill_MassIn,
        'Landfill_MassOut': Landfill_MassOut,
        'LandfillRB_MassIn': LandfillRB_MassIn,
        'LandfillRB_MassOut': LandfillRB_MassOut,
        'LandfillRE_MassIn': LandfillRE_MassIn,
        'LandfillRE_MassOut': LandfillRE_MassOut,
        'MRFpc_MassIn': MRFpc_MassIn,
        'MRFpc_MassOut': MRFpc_MassOut,
        'MRFpm_MassIn': MRFpm_MassIn,
        'MRFpm_MassOut': MRFpm_MassOut,
        'MRFg_MassIn': MRFg_MassIn,
        'MRFg_MassOut': MRFg_MassOut,
        'MTb_MassIn': MTb_MassIn,
        'MTb_MassOut': MTb_MassOut,
        'MTmw_MassIn': MTmw_MassIn,
        'MTmw_MassOut': MTmw_MassOut,
        'AD_MassIn': AD_MassIn,
        'AD_MassOut': AD_MassOut,
        'ADRE_MassIn': ADRE_MassIn,
        'ADRE_MassOut': ADRE_MassOut,
        'Composting_MassIn': Composting_MassIn,
        'Composting_MassOut': Composting_MassOut,
        'Incineration_MassIn': Incineration_MassIn,
        'Incineration_MassOut': Incineration_MassOut,
        'Gasification_MassIn': Gasification_MassIn,
        'Gasification_MassOut': Gasification_MassOut,
        'Pyrolysis_MassIn': Pyrolysis_MassIn,
        'Pyrolysis_MassOut': Pyrolysis_MassOut,
        'SR_MassIn': SR_MassIn,
        'SR_MassOut': SR_MassOut,
        'FTh_MassIn': FTh_MassIn,
        'FTh_MassOut': FTh_MassOut,
        'FTb_MassIn': FTb_MassIn,
        'FTb_MassOut': FTb_MassOut,
        'MS_MassIn': MS_MassIn,
        'MS_MassOut': MS_MassOut,
        'AS_MassIn': AS_MassIn,
        'AS_MassOut': AS_MassOut,
        'Fermentation_MassIn': Fermentation_MassIn,
        'Fermentation_MassOut': Fermentation_MassOut}

df_md = pd.DataFrame(data=data2, index=WasteTypeToSave)
for l in range (len(ProcessType)):
    exec("del %s_MassIn" %(ProcessType[l]))
    exec("del %s_MassOut" %(ProcessType[l]))
del l, data2, WasteTypeToSave



#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Detailed Revenues   #######
"""

Revenues = []
for n in range (len(CoproductType)):
    Revenues.append(float(df_Revenues.at[CoproductType[n],'Commercial Price [EUR/unit]']) * ( sum(model.x[i,n].value for i in model.WasteType) + model.xs[n].value ) )

data = {'Revenues [EUR]': Revenues}
df_detrev = pd.DataFrame(data=data, index=CoproductType)
del n, data, Revenues


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Detailed Costs   #######
"""

AllTechs = ProcessType.copy()
AllTechs.insert(0, 'Door2Door Collection')
AllTechs.insert(1, 'Curside Collection')


CAPEX = []
CollCAPEX = []
ProcCAPEX = []
OPEX = []
CollOPEX = []
ProcOPEX = []


CAPEX.append( sum( -float(df_Collection_Cost_D2D.at['CAPEX [EUR/ton]', BinType[k]]) * model.sc[0,k].value for k in model.BinType) )
CAPEX.append( sum( - float(df_Collection_Cost_CS.at['CAPEX [EUR/ton]', BinType[k]]) * model.sc[1,k].value for k in model.BinType) )
CollCAPEX.append( sum( -float(df_Collection_Cost_D2D.at['CAPEX [EUR/ton]', BinType[k]]) * model.sc[0,k].value \
                       - float(df_Collection_Cost_CS.at['CAPEX [EUR/ton]', BinType[k]]) * model.sc[1,k].value \
                    for k in model.BinType) )
OPEX.append(sum( -float(df_Collection_Cost_D2D.at['OPEX [EUR/ton]', BinType[k]]) * model.mp[0,k].value for k in model.BinType ) )
OPEX.append(sum( - float(df_Collection_Cost_CS.at['OPEX [EUR/ton]', BinType[k]]) * model.mp[1,k].value for k in model.BinType ) )
CollOPEX.append(sum( -float(df_Collection_Cost_D2D.at['OPEX [EUR/ton]', BinType[k]]) * model.mp[0,k].value for k in model.BinType ) \
              + sum( - float(df_Collection_Cost_CS.at['OPEX [EUR/ton]', BinType[k]]) * model.mp[1,k].value for k in model.BinType ) )
for l in range (len(ProcessType)):
    CAPEX.append(-float(df_Processes_Cost.at[ProcessType[l],'CAPEX [EUR/ton]']) * model.sp[l].value )
    ProcCAPEX.append(-float(df_Processes_Cost.at[ProcessType[l],'CAPEX [EUR/ton]']) * model.sp[l].value )
    OPEX.append(-float(df_Processes_Cost.at[ProcessType[l], 'OPEX [EUR/ton]']) \
                    * ( sum(model.mi[i,l].value for i in model.WasteType) + model.msi[l].value ) )
    ProcOPEX.append(-float(df_Processes_Cost.at[ProcessType[l], 'OPEX [EUR/ton]']) \
                    * ( sum(model.mi[i,l].value for i in model.WasteType) + model.msi[l].value ) )


data = {'CAPEX [EUR]': CAPEX,
        'OPEX [EUR]': OPEX}
df_detcost = pd.DataFrame(data=data, index=AllTechs)
del l, data, CAPEX, OPEX, AllTechs


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Detailed Emissions   #######
"""

AllTechs = ProcessType.copy()
AllTechs.insert(0, 'Door2Door Collection')
AllTechs.insert(1, 'Curside Collection')


Emissions = []

for j in range (len(CollectionType)):
    Emissions.append(model.Ec[j].value)
for l in range (len(ProcessType)):
    Emissions.append(model.Ep[l].value)

data = {'Emissions [kg CO2 eq]': Emissions}
df_Emissions = pd.DataFrame(data=data, index=AllTechs)
del j, l, data, Emissions, AllTechs


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   Annual Costs   #######
"""

CollCAPEX = sum(CollCAPEX)
CollOPEX = sum(CollOPEX)

ProcCAPEX = sum(ProcCAPEX)
ProcOPEX = sum(ProcOPEX)

Rev = sum ( float(df_Revenues.at[CoproductType[n],'Commercial Price [EUR/unit]']) * ( sum(model.x[i,n].value for i in model.WasteType) + model.xs[n].value) for n in model.CoproductType)


if results['Problem'][0]['Lower bound'] == results['Problem'][0]['Upper bound']:
    BreakDown_1stLevel = [-results['Problem'][0]['Lower bound'], '', (CollCAPEX+CollOPEX), (ProcCAPEX+ProcOPEX), Rev]
    BreakDown_1stLevel_Emissions = [ sum(model.Ec[j].value for j in model.CollectionType) + sum(model.Ep[l].value for l in model.ProcessType), '', sum(model.Ec[j].value for j in model.CollectionType), sum(model.Ep[l].value for l in model.ProcessType), '']
    index=['Total', 'BREAK-DOWN', 'Collection', 'Processes', 'Revenues']
else:
    BreakDown_1stLevel = [-results['Problem'][0]['Lower bound'], -results['Problem'][0]['Upper bound'], '', -(CollCAPEX+CollOPEX), -(ProcCAPEX+ProcOPEX), Rev]
    BreakDown_1stLevel_Emissions = [ sum(model.Ec[j].value for j in model.CollectionType) + sum(model.Ep[l].value for l in model.ProcessType), '', '', sum(model.Ec[j].value for j in model.CollectionType), sum(model.Ep[l].value for l in model.ProcessType), '']
    index=['Lower bound','Upper bound', 'BREAK-DOWN', 'Collection', 'Processes', 'Revenues']


data = {'AnnualCost [EUR]': BreakDown_1stLevel,
        'AnnualEmissions [kg CO2 eq]': BreakDown_1stLevel_Emissions}
df_totals = pd.DataFrame(data, index=index)
del index, data


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------


#index = CollectionType + ProcessType
#Emissions = []
#for j in range (len(CollectionType)):
#    Emissions.append(model.Ec[j].value)
#for l in range (len(ProcessType)):
#    Emissions.append(model.Ep[l].value)
#
#data = {'Emissions [kg CO2 eq]': Emissions}
#df_E = pd.DataFrame(data, index=index)
#del j, l, index, data


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------


#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
#######   CREATING THE EXCEL FILE   #######
"""

# Creating the excel file

#initialze the excel writer
writer = pd.ExcelWriter('NEW_Results_v6_2_10_9.xlsx', engine='xlsxwriter')

#store your dataframes in a  dict, where the key is the sheet name you want
#frames = {'data_CoproductPrices': df_CoproductsPrices, 'result_CoproductProd': df_x, \
#          'data_TechCost': df_Techs.iloc[:, 0], 'result_TechUse': df_m, \
#          'Number_Trucks': df_trucks, 'Emissions': df_e, 'result_TotalCost': df_totals}
frames = {'Total': df_totals, \
          'DetailedEmissions': df_Emissions, \
          'DetailedCost': df_detcost, \
          'DetailedRevenues' : df_detrev, \
          'Collection': df_c_share, \
          'TechUse': df_m, \
          'TechUseDetail' : df_md, \
          'ValueAddedProducts': df_x, \
          'RecyclingRates': df_Rr, \
          'ReciclablesPurity': df_Pr}


#now loop thru and put each on a specific sheet
for sheet, frame in frames.items(): # .use .items for python 3.X
    if sheet == 'ValueAddedProducts':
        frame.to_excel(writer, sheet_name = sheet, index = False)
    else:
        frame.to_excel(writer, sheet_name = sheet)


df_c_Inst.to_excel(writer, sheet_name='Collection',startrow=5, startcol=3, index_label='Current installed capacity [ton]')
df_c_ExtraInst.to_excel(writer, sheet_name='Collection',startrow=9, startcol=3, index_label='Extra required installed capacity [ton]') 
df_c_MassIn.to_excel(writer, sheet_name='Collection',startrow=13, startcol=3, index_label='MassIn [ton]') 




# Get the xlsxwriter workbook objects
workbook  = writer.book

# Add some cell formats
format1 = workbook.add_format({'num_format': '#,##0.00', 
                               'font_color': '#5081BB',
                               'font_size': 12})
format1.set_bold()
format2 = workbook.add_format({'num_format': '#,##0.0',
                               'font_color': '#A6A6A6',
                               'italic': True ,
                               'font_size': 12})
format3 = workbook.add_format({'font_color': '#5081BB',
                               'font_size': 10})
format4 = workbook.add_format({'num_format': '#,##0.00',
                               'font_color': '#A6A6A6',
                               'font_size': 9})
format5 = workbook.add_format({'num_format': '0%'})
format6 = workbook.add_format({'bg_color': '#A9D08E',
                              'font_color': '#000000'})
format7 = workbook.add_format({'bg_color': '#FFC7CE',
                              'font_color': '#9C0006'})
format8 = workbook.add_format({'num_format': '#,##0.00', 
                               'font_color': '#009FE3',
                               'font_size': 12})
format8.set_bold()
format9 = workbook.add_format({'num_format': '#,##0.00', 
                               'font_color': '#C55A11',
                               'font_size': 12})
format9.set_bold()
format10 = workbook.add_format({'num_format': '#,##0.00', 
                               'font_color': '#5081BB',
                               'font_size': 13})
format10.set_bold()
format_example = workbook.add_format({'bg_color': '#0096A9',
                               'text_wrap':'true'})
format11 = workbook.add_format({'font_color': '#A6A6A6',
                               'italic': True ,
                               'font_size': 10})
format12 = workbook.add_format({'font_color': '#5081BB',
                               'font_size': 13})
format12.set_bold()



# Get the xlsxwriter worksheet objects and format
worksheet = writer.sheets['Total']
worksheet.set_column('A:A', 20) # Set the column width but not the format
worksheet.set_column('B:B', 20) # Set the column width and format
worksheet.set_column('C:C', 27) # Set the column width and format
#worksheet.set_column(2, 2, None, format1) # Set the format but not the column width
number_rows = len(df_totals.index) + 1
worksheet.conditional_format("$B$2:$C$%d" % (number_rows),
                             {"type": "cell",
                              "criteria": '>=',
                              "value": 0.00001,
                              "format": format10
                             }
)
worksheet.conditional_format("$B$2:$C$%d" % (number_rows),
                             {"type": "cell",
                              "criteria": '<=',
                              "value": -0.00001,
                              "format": format9
                             }
)
worksheet.hide_gridlines(2)



worksheet = writer.sheets['DetailedEmissions']
worksheet.set_column(0, 0, 20, format3) # Set the column width but not the format
worksheet.set_column(1, 1, 20, format2) # Set the column width and format
number_rows = len(df_Emissions.index) + 1
worksheet.conditional_format("$B$2:$C$%d" % (number_rows),
                             {"type": "cell",
                              "criteria": '>=',
                              "value": 0.00001,
                              "format": format10
                             }
)
worksheet.hide_gridlines(2)



worksheet = writer.sheets['DetailedCost']
worksheet.set_column(0, 0, 20, format3) # Set the column width but not the format
worksheet.set_column(1, 1, 20, format2) # Set the column width and format
worksheet.set_column(2, 2, 20, format2) # Set the column width and format
number_rows = len(df_detcost.index) + 1
worksheet.conditional_format("$B$2:$C$%d" % (number_rows),
                             {"type": "cell",
                              "criteria": '>=',
                              "value": 0.00001,
                              "format": format10
                             }
)
worksheet.conditional_format("$B$2:$C$%d" % (number_rows),
                             {"type": "cell",
                              "criteria": '<=',
                              "value": -0.00001,
                              "format": format9
                             }
)
worksheet.hide_gridlines(2)



worksheet = writer.sheets['DetailedRevenues']
worksheet.set_column(0, 0, 22, format3) # Set the column width but not the format
worksheet.set_column(1, 1, 15, format2) # Set the column width and format
number_rows = len(df_detrev.index) + 1
worksheet.conditional_format("$B$2:$B$%d" % (number_rows),
                             {"type": "cell",
                              "criteria": '>=',
                              "value": 0.001,
                              "format": format10
                             }
)
worksheet.conditional_format("$B$2:$B$%d" % (number_rows),
                             {"type": "cell",
                              "criteria": '<=',
                              "value": -0.00001,
                              "format": format9
                             }
)
worksheet.hide_gridlines(2)



worksheet = writer.sheets['Collection']
worksheet.set_column(0, 0, 10, format3) # Set the column width but not the format
worksheet.set_column(1, 1, 7, format2) # Set the column width and format
worksheet.set_column(2, 2, 8, format2) # Set the column width and format
worksheet.set_column(3, 3, 35, format2) # Set the column width but not the format
worksheet.set_column(4, 4, 13, format2) # Set the column width and format
worksheet.set_column(5, 5, 17, format2) # Set the column width and format
worksheet.set_column(6, 6, 14, format2) # Set the column width and format
worksheet.set_column(7, 7, 12, format2) # Set the column width and format
worksheet.set_column(8, 8, 12, format2) # Set the column width and format
number_rows = len(df_c_share.index) + 1
worksheet.conditional_format("$B$2:$B$%d" % (number_rows),
                             {"type": "text",
                              "criteria": "containing",
                              "value": "Yes",
                              "format": format10
                             }
)
worksheet.conditional_format("$C$2:$E$%d" % (number_rows),
                             {"type": "cell",
                              "criteria": '>=',
                              "value": 0.00001,
                              "format": format10
                             }
)
worksheet.hide_gridlines(2)



worksheet = writer.sheets['TechUse']
worksheet.set_column(0, 0, 20, format3) # Set the column width but not the format
worksheet.set_column(1, 1, 10, format2) # Set the column width and format
worksheet.set_column(2, 2, 29, format2) # Set the column width and format
worksheet.set_column('D:D', 35, format2) # Set the column width but not the format
#worksheet.set_column('D:D', 7, format4) # Set the column width but not the format
worksheet.set_column(4, 4, 14, format2) # Set the column width and format
worksheet.set_column(5, 5, 14, format2) # Set the column width and format
number_rows = len(df_m.index) + 1
worksheet.conditional_format("$C$2:$F$%d" % (number_rows),
                             {"type": "cell",
                              "criteria": '>=',
                              "value": 0.00001,
                              "format": format10
                             }
)
worksheet.conditional_format("$C$2:$F$%d" % (number_rows),
                             {"type": "cell",
                              "criteria": '<=',
                              "value": -0.00001,
                              "format": format9
                             }
)
worksheet.conditional_format("$B$2:$B$%d" % (number_rows),
                             {"type": "text",
                              "criteria": "containing",
                              "value": "Yes",
                              "format": format10
#                              "format": format6
                             }
)
worksheet.hide_gridlines(2)



worksheet = writer.sheets['TechUseDetail']
worksheet.set_column(0, 0, 28, format3) # Set the column width but not the format
worksheet.set_column(1, 40, 22, format2) # Set the column width and format
number_rows = len(df_md.index) + 1
worksheet.conditional_format("$B$2:$AO$%d" % (number_rows),
                             {"type": "cell",
                              "criteria": '>=',
                              "value": 0.00001,
                              "format": format10
                             }
)
worksheet.conditional_format("$B$2:$AO$%d" % (number_rows),
                             {"type": "cell",
                              "criteria": '<=',
                              "value": -0.00001,
                              "format": format9
                             }
)
worksheet.hide_gridlines(2)



worksheet = writer.sheets['ValueAddedProducts']
worksheet.set_column(0, 0, 32, format3) # Set the column width but not the format
worksheet.set_column(1, 1, 15, format2) # Set the column width and format
worksheet.set_column(2, 2, 5, format4) # Set the column width and format
worksheet.set_column(3, 3, 45, format3) # Set the column width but not the format
number_rows = len(df_x.index) + 1
worksheet.conditional_format("$B$2:$B$%d" % (number_rows),
                             {"type": "cell",
                              "criteria": '>=',
                              "value": 0.001,
                              "format": format10
                             }
)
worksheet.hide_gridlines(2)



worksheet = writer.sheets['RecyclingRates']
worksheet.set_column(0, 0, 28, format3) # Set the column width but not the format
worksheet.set_column(1, 1, 16, format2) # Set the column width and format
number_rows = len(df_Rr.index) + 1
worksheet.conditional_format("$B$2:$B$%d" % (number_rows),
                             {"type": "cell",
                              "criteria": '>=',
                              "value": 0.65,
                              "format": format10
                             }
)
worksheet.conditional_format("$B$2:$B$%d" % (number_rows),
                             {"type": "cell",
                              "criteria": '<',
                              "value": 0.65,
                              "format": format9
                             }
)
worksheet.hide_gridlines(2)



worksheet = writer.sheets['ReciclablesPurity']
worksheet.set_column(0, 0, 35, format3) # Set the column width but not the format
worksheet.set_column(1, 1, 17, format2) # Set the column width and format
worksheet.set_column(2, 2, 50, format3) # Set the column width but not the format
number_rows = len(df_Pr.index) + 1
worksheet.conditional_format("$B$2:$B$%d" % (number_rows),
                             {"type": "cell",
                              "criteria": '>=',
                              "value": 0.95,
                              "format": format10
                             }
)
worksheet.hide_gridlines(2)



# Close the Pandas Excel writer and output the Excel file.
workbook.close()
writer.save()


##SOURCE : https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.to_excel.html


##**
## Turn off the default header and skip one row to allow us to insert a
## user defined header.
#df.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False)
#
## Get the xlsxwriter workbook and worksheet objects.
#workbook  = writer.book
#worksheet = writer.sheets['Sheet1']
#
## Add a header format.
#header_format = workbook.add_format({
#    'bold': True,
#    'text_wrap': True,
#    'valign': 'top',
#    'fg_color': '#D7E4BC',
#    'border': 1})
#
## Write the column headers with the defined format.
#for col_num, value in enumerate(df.columns.values):
#    worksheet.write(0, col_num + 1, value, header_format)
## SOURCE : https://stackoverflow.com/questions/50401744/using-xlsxwriter-to-align-left-a-row
## SOURCE2 : https://xlsxwriter.readthedocs.io/working_with_pandas.html#formatting-of-the-dataframe-headers
##**


##**
#import pandas as pd
#writer = pd.ExcelWriter('simple.xlsx', engine='xlsxwriter')
#df=df.style.set_properties(**{'text-align': 'center'})
#df.to_excel(writer, sheet_name='Sheet1')
#writer.save()
## SOURCE : https://stackoverflow.com/questions/58297534/center-align-using-xlsxwriter-without-conditional-formatting
##**


##**
#import pandas as pd
#
#df = pd.DataFrame({"Name": ['A', 'B', 'C', 'D', 'E'], 
#"Status": ['SUCCESS', 'FAIL', 'SUCCESS', 'FAIL', 'FAIL']})
#
#number_rows = len(df.index) + 1
#
#writer = pd.ExcelWriter('Report_1.xlsx', engine='xlsxwriter')
#
#df.to_excel(writer, sheet_name='Sheet1', index=False)
#
#workbook  = writer.book
#worksheet = writer.sheets['Sheet1']
#
#format1 = workbook.add_format({'bg_color': '#FFC7CE',
#                              'font_color': '#9C0006'})
#
#worksheet.conditional_format("$A$1:$B$%d" % (number_rows),
#                             {"type": "formula",
#                              "criteria": '=INDIRECT("B"&ROW())="SUCCESS"',
#                              "format": format1
#                             }
#)
#
#workbook.close()
## SOURCE : https://stackoverflow.com/questions/48527243/format-entire-row-with-conditional-format-using-pandas-xlswriter-module
##**


##**
#workbook = xlsxwriter.Workbook('demo1.xlsx')
#worksheet = workbook.add_worksheet()
#format = workbook.add_format({ 'bg_color': '#5081BB','font_color': 
#'#FFFFFF','font_size': 12,'text_wrap':'true'})
#textWrap = workbook.add_format({'text_wrap':'true'})
#
#worksheet.set_column('B:B', 18, format)       //format formatting apply to entire B column
#worksheet.set_column('C:C', None, textWrap)   //textWrap formatting apply to entire C column
##**
#worksheet0.conditional_format('B:B', {
#    'type': 'cell',
#    'criteria': '<',
#    'value': 0, 'format': number_format
#        }
#    )
#
#worksheet0.conditional_format('J3:V22', {
#    'type': 'cell',
#    'criteria': '>=',
#    'value': 0, 'format': number_format
#    }
#)
## SOURCE : https://stackoverflow.com/questions/28966420/how-to-set-formatting-for-entire-row-or-column-in-xlsxwriter-python
##**


##**
#import pandas as pd
#writer = pd.ExcelWriter('demo.xlsx', engine='xlsxwriter')
#
#df = pd.DataFrame({'Name': ['E', 'F', 'G', 'H'],
#                   'Age': [100, 70, 40, 60]})
#
## Write data to an excel
#df.to_excel(writer, sheet_name="Sheet1", index=False)
#
## Get workbook
#workbook = writer.book
## Get Sheet1
#worksheet = writer.sheets['Sheet1']
#
#cell_format = workbook.add_format()
#cell_format.set_bold()
#cell_format.set_font_color('blue')
#
#worksheet.set_column('B:B', None, cell_format)
#
#writer.close()
## SOURCE : https://medium.com/codeptivesolutions/https-medium-com-nensi26-formatting-in-excel-sheet-using-xlsxwriter-part-1-2c2c547b2bea
##**

    
##**
#import pandas as pd
#
#df = pd.DataFrame({'Test_1':['Pass','Fail', 'Pending', 'Fail'],
#                   'expect':['d','f','g', 'h'],
#                   'Test_2':['Pass','Pending', 'Pass', 'Fail'],
#                  })
#
#fmt_dict = {
#    'Pass': 'background-color: green',
#    'Fail': 'background-color: red',
#    'Pending': 'background-color: yellow; border-style: solid; border-color: blue; color:red',
#}
#
#def fmt(data, fmt_dict):
#    return data.replace(fmt_dict)
#
#styled = df.style.apply(fmt, fmt_dict=fmt_dict, subset=['Test_1', 'Test_2' ])
#styled.to_excel('styled.xlsx', engine='openpyxl')
## SOURCE : https://stackoverflow.com/questions/44150078/python-using-pandas-to-format-excel-cell
##**


##**
#from xlsxwriter.workbook import Workbook
#
#workbook = Workbook('hello_world.xlsx')
#worksheet = workbook.add_worksheet()
#
#worksheet.write('A1', 'Hello world')
#worksheet.hide_gridlines(2)
#
#workbook.close()
## SOURCE : https://stackoverflow.com/questions/16342893/removing-gridlines-from-excel-using-python-openpyxl
##**




#%%
#%%
#%%
#%%
#%%
#%%
#%%
#%%
#%%
#%%
#%%