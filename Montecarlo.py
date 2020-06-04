import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from pandas import read_excel

Excel_Filepath = 'stock_return.xlsx'

AnnualisedReturn = [0.218232133, 0.28170943, 0.403572534, 0.116522438, 0.470356619, 0.158961835, 0.221796948, 0.310196222, 0.192966333, 0.196209769, 0.19937625, 0.328680283, 0.280773141]
AnnualisedDiavition = [0.28549703, 0.428536523, 0.40902669, 0.334740171, 0.550055992, 0.219055727, 0.242188706, 0.219002851, 0.1349524, 0.206914844, 0.370937867, 0.264619756, 0.510383668]
SharpAnnula = [0.207232133, 0.27480943, 0.403572534, 0.082522438, 0.453356619, 0.150461835, 0.219496948, 0.303996222, 0.175466333, 0.184509769, 0.19937625, 0.320480283, 0.280773141]


def Cov_Mat_Calculation(Return_excel_filepath, annual):
    stock_data = read_excel(Return_excel_filepath)
    cov_mat = stock_data.cov()
    return cov_mat * annual

def MontCarlo(Return_List, Cov_Mat_Annual, ReturnSharp, Setting_Iteration):
    IterCounter = 0
    Return_Sim = []
    Volativity_Sim = []
    Return_Sharp = []
    WeightGroupList = []
    while IterCounter < Setting_Iteration:
        #np.random.seed(12122)
        random = np.random.random(len(Return_List))
        random_weight = random/np.sum(random)
        WeightGroupList.append(random_weight)
        WeightedReturn = sum(random_weight * np.array(Return_List))
        Return_Sim.append(WeightedReturn)
        SharpWeighted = sum(random_weight * np.array(ReturnSharp))
        Return_Sharp.append(SharpWeighted)
        Random_Volatility= np.sqrt(np.dot(random_weight.T,np.dot(Cov_Mat_Annual,random_weight)))
        Volativity_Sim.append(Random_Volatility)
        IterCounter += 1
    return Return_Sim, Volativity_Sim, Return_Sharp, WeightGroupList

def SharpCalculation(ReturnList_for_Sharp, v_list, ConstantAnnual = 0.02):
    SharpForMonte = []
    for index, return_value in enumerate(ReturnList_for_Sharp):
        sharp = (return_value - ConstantAnnual)/v_list[index]
        SharpForMonte.append(sharp)
    return SharpForMonte

def MontCarloPlot(VolativityList, ReturnList, Sharpmaxindex, MinVolatindex, alpha = 0.5):
    plt.scatter(Volativity_List, Return_List, alpha=0.5)

    #plot max-sharp point
    x = VolativityList[Sharpmaxindex]
    y = ReturnList[Sharpmaxindex]
    plt.scatter(x, y, color='red')
    plt.text(np.round(x,4),np.round(y,4),(np.round(x,4),np.round(y,4)),ha='left',va='bottom',fontsize=15)

    #plot min-vola point
    m = VolativityList[int(MinVolatindex)]
    n = ReturnList[int(MinVolatindex)]
    plt.scatter (m,n , color='green')
    plt.text(np.round(m,4),np.round(n,4),(np.round(m,4),np.round(n,4)),ha='left',va='bottom',fontsize=15)

    plt.xlabel("Volativity")
    plt.ylabel("Returns")

    plt.show()


Cov_Mat_Ann = Cov_Mat_Calculation(Excel_Filepath,12)
Return_List, Volativity_List, Return_Sharp, WeightedGroupList = MontCarlo(AnnualisedReturn, Cov_Mat_Ann, SharpAnnula, 10000)
SharpList = SharpCalculation(Return_Sharp, Volativity_List, 0.02)
MaxSharpindex = SharpList.index(max(SharpList))
MinVolaindex = Volativity_List.index(min(Volativity_List))

print('Max-Sharp Investment Weight Should be')
print(WeightedGroupList[MaxSharpindex])

print('Min-Volativity Investment Weight Should be')
print(WeightedGroupList[MinVolaindex])

MontCarloPlot(Volativity_List,Return_List,MaxSharpindex,MinVolaindex, 0.5)