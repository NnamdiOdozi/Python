import xlwings as xw
import numpy as np
import pandas as pd
import math
#import seaborn as sns
import datetime
import matplotlib.pyplot as plt
#from pandas-datareader import data, wb
from scipy.interpolate import interp1d
#matplotlib inline


start = datetime.datetime(2006,1,1)
end = datetime.datetime(2016,1,1)

@xw.func
@xw.arg('X1', doc = 'This is the input')
@xw.arg('Y', np.array, ndim=1)
@xw.arg('X', np.array, ndim=1)

def Interpol(X1, Y, X):
    """Interpolates a table of values"""
    y_interp = interp1d(X, Y)
    return y_interp(X1)

@xw.func
@xw.arg('Y', np.array, ndim=1)
@xw.arg('X', np.array, ndim=1)

def Adding(Y, X):
    return (np.sum(Y) + np.sum(X))



#def  MortDecr(Age, MortTable, AgeRating, MortScalar, Term):

# Script to calculate the probability of surviving 10 years from age 27
Age = 27
Term = 10
AgeRating = 0
MortScalar = 1
#Reads in a mortality table from open Excel file.  
# I haven't bothered to add the filename as the plan is to convert this script to an imported Excel function and 
# then will not need the Excel filename since it will be called from the same file
df_MortTable =  xw.Range('A67_70_Months').options(pd.DataFrame,index=False, header=False).value
df_MortTable.columns = ['Age','qx']
df_MortTable.set_index('Age', inplace = True, append = False, drop=True) 
#df_MortTable.reindex(index = ['Age'], columns = ['qx']) #The reset.index method produces an extra index col which is not needed
Age = math.floor(Age * 12)
AgeRating = AgeRating *12
Term = math.ceil(Term *12) 
End = Age + Term
#print(df_MortTableValues)
df_MortTable['px'] = 1 - df_MortTable['qx']

df_MortTable['Cum_p'] = df_MortTable['px'].cumprod()
df_MortTable['MortDecr'] = df_MortTable['Cum_p']/df_MortTable['Cum_p'].loc[Age]
ProbSurviving = df_MortTable['MortDecr'].loc[End].item()

Tmp = df_MortTable['px'].loc[Age:Age+Term].shift(periods=1, fill_value=1).cumprod() 

df_MortTable['Benefit'] = np.zeros(len(df_MortTable))
df_MortTable.loc[Age:Age+Term, "Benefit"] = 1
pd.set_option("display.max_rows", 2000, "display.max_columns", None)
print(df_MortTable)
#print(MortDecr.loc[Age:End])
print(df_MortTable['MortDecr'].loc[End].item(), '\n')
print(df_MortTable['MortDecr'].at[End], '\n')



@xw.func
@xw.arg('Age')
@xw.arg('AgeRating')
@xw.arg('df_MortTableValues', pd.DataFrame, index=True, header=False )
@xw.arg('MortScalar')
@xw.arg('Term')

def  MortDecr(Age,AgeRating, df_MortTableValues, MortScalar, Term):

    """Python Function to calculate the probability of surviving 10 years from given age"""
    #Age = 27
    #Term = 10
    #AgeRating = 0
    #MortScalar = 1

    #Reads in a mortality table from open Excel file.  
    # I haven't bothered to add the filename as the plan is to convert this script to an imported Excel function and 
    # then will not need the Excel filename since it will be called from the same file
    #df_MortTableValues =  xw.Range(MortTable).options(pd.DataFrame, index=True, header=False).value 
    Age = math.floor(Age * 12)
    AgeRating = AgeRating *12
    Term = math.ceil(Term *12) 
    End = Age + Term
    #print(df_MortTableValues)
    df_p = 1 - df_MortTableValues

    df_Cum_p = df_p.cumprod()
    MortDecr = df_Cum_p/df_Cum_p.loc[Age]
    ProbSurviving = MortDecr.loc[End].item()

    #print(MortDecr.loc[Age:End])

    return MortDecr.loc[End].item()


# X1 = 1/12 
# X = [0,1,2,3]
# Y = [0.0953, 0.0953, 0.1197, 0.1223]
# Y1 = Interpol(X1, Y, X)
# Print(Y1)