import openpyxl as op
import os
import itertools

#Input Data
Titles = ["AACT Capacity", "Package 1", "Package 2", "Package 3", "Package 4"]
PacakgeCost_Estimate = [0, 900, 700, 600, 800]
aact_Requirement = [0, 1800, 1400, 1200, 1600]
ContractorA = [2500, 790, 600, 510, 690]
ContractorB = [2800, 780, 610, 515, 730]
ContractorC = [3200, 800, 650, 530, 740]
ContractorD = [3600, 820, 640, 540, 710]

Contractors = [ContractorA, ContractorB, ContractorC, ContractorD]

#Creating Combinations
packs  = []
for package in range(1, len(Titles)):
    packs.append(Titles[package])

Cont_s = []
for contractor in Contractors:
    Cont_s.append(contractor[1:])

level_V = len(packs)

Req_List = []
List_Indexes = []
for index in range(0, len(Cont_s[0])):
    dynlist = []
    dynIndexes = []
    for A in Cont_s:
        dynlist.append(A[index])
        dynIndexes.append(index)
    Req_List.append(dynlist)
    List_Indexes.append(dynIndexes)

List_Indexes1 = ["A1", "B1", "C1", "D1"]
List_Indexes2 = ["A2", "B2", "C2", "D2"]
List_Indexes3 = ["A3", "B3", "C3", "D3"]
List_Indexes4 = ["A4", "B4", "C4", "D4"]

BidsCombinations = list(itertools.product(Req_List[0], Req_List[1], Req_List[2], Req_List[3]))
CombinationIndex = list(itertools.product(List_Indexes1, List_Indexes2, List_Indexes3, List_Indexes4))


def CheckCapacity(ValueList, IndexList):
    IndexList = list(IndexList)    
    AlphabetListUnique = []
    for index in IndexList:
        AlphabetListUnique.append(index[0])

    AlphabetListUnique = list(set(AlphabetListUnique))
    AlphabetListUnique.sort()

    #Packages Capacities defined by the winning Alphabets
    AACT_Capacities = []
    for data in AlphabetListUnique:
        if data == "A":
            AACT_Capacities.append(Contractors[0][0])
        if data == "B":
            AACT_Capacities.append(Contractors[1][0])
        if data == "C":
            AACT_Capacities.append(Contractors[2][0])
        if data == "D":
            AACT_Capacities.append(Contractors[3][0])
        
    
    #Checking whether the Requirement is under the capacity or not
    alphaNum_Dict = {"A":1, "B":2, "C":3, "D":4}
    capacityCheck = True
    CostIncurred = 0.0000
    for index,  i in enumerate(AlphabetListUnique):
        AACT_Requirement = 0.000
        for item in IndexList:
            Alphabet = item[0]
            if i == Alphabet:
                AACT_Requirement += aact_Requirement[int(item[1])]
        AACT_Capacity = AACT_Capacities[index]
        

        if AACT_Requirement>AACT_Capacity:
            capacityCheck = False
    for item in IndexList:
        Alphabet = item[0]
        Numerical = int(item[1])
        a = int(alphaNum_Dict[Alphabet]) - 1

        CostIncurred += int(Contractors[a][Numerical])
    return CostIncurred, capacityCheck


#Calling function Capccity and checking whether the condition persists
ReqOutputData = []
for index, IndexL in enumerate(CombinationIndex):
    Sum, Check = CheckCapacity(index, IndexL)
    if Check:
        ReqOutputData.append([IndexL, Sum])

Sorted_ReqOutputData = sorted(ReqOutputData, key=lambda x: x[1])

for value in Sorted_ReqOutputData:
    print(value)
