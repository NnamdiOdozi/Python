import xlwings as xw
import numpy as np

@xw.sub
def SayHello():
    wb = xw.Book.caller()
    wb.sheets[0].range("A1").value = "Hello world"


@xw.func
@xw.arg('X1', np.array, ndim=2)
@xw.arg('X2', np.array, ndim=2)
@xw.arg('Y1', np.array, ndim=2)

def CalcAnnuity(X1, X2, Y1)
    
    return CalcAnnuity