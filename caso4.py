
"""
Autores:
    Victor Rivera Gijón
    Ismael Garcia Mayorga
    Juan Manuel Alcalá Palomo
    Manuel Domínguez Montero
"""  
#Caso 2 IO -> Problema de transporte.
# -*- coding: utf-8 -*-
from ortools.linear_solver import pywraplp
from IOfunctionsExcel import *

name='Datos.xlsx'
excel_doc=openpyxl.load_workbook(name,data_only=True)
sheet=excel_doc['Hoja1']

tipos = [ 0 , 1, 2 ]
demanda_minima = Read_Excel_to_List(sheet, 'c4', 'c8')
horas = Read_Excel_to_List(sheet, 'e4', 'e8')
franjas = [ 0, 1, 2, 3 , 4 ]
min_prod = Read_Excel_to_List(sheet, 'c13', 'c15')
max_prod = Read_Excel_to_List(sheet, 'd13', 'd15')
coste_min_hora = Read_Excel_to_List(sheet, 'e13', 'e15')
coste_mw = Read_Excel_to_List(sheet, 'f13', 'f15')
coste_arranque = Read_Excel_to_List(sheet, 'g13', 'g15')
numero_gen = Read_Excel_to_List(sheet, 'c19', 'c21')



def Problema():
    solver=pywraplp.Solver.CreateSolver('CLP')
    
    """ Variables cerrojo """
    A = {}
    for i in tipos:
        A[i] = {}
        for n in franjas:
            A[i][n] = solver.IntVar(0,1,'Generadores de tipo i encendidos en la franja n')
            
    """Variables de cuanto genera cada generador """
    G = {}
    cantidad_generadores = {}
    for i in tipos:
        G[i] = {}
        b = numero_gen[i]
        cantidad_generadores[i] = range(b)
        for j in range(b):
            G[i][j] = {}
            for n in franjas:
                G[i][j][n] = solver.IntVar(0, max_prod[i],'Generacion en Mw del generador de tipo i numero j en la franja horaria n')
                solver.Add( G[i][j][n] >= min_prod[i] * A[i][n]) #restriccion de produccion mínima
    Gm = {}  #Esta variable probablemente sea innecesaria
    for i in tipos:
        Gm[i] = {}
        b = numero_gen[i]
        for j in range(b):
            G[i][j] = max_prod[i]
            
    """ Restricciones """        
    
    for n in franjas:
        for i in tipos:
            b = numero_gen[i]
            solver.Add( demanda_minima[n] == solver.Sum( G[i][j][n] * A[i][n] for j in range(b)))
            solver.Add( demanda_minima[n] * 1.15 <= solver.Sum( Gm[i][j] * A[i][n] for j in range(b))) #aseguramos q la demanda maxima siempre pueda cumplir el 15%
            
    
    """ Costes """  
      
    for n in franjas:
        for i in tipos:
            coste_extra[n][i] = horas[n] * coste_mw[i]
            coste_min[n][i] = horas[n] * coste_min[i]
    
    solver.Minimize( solver.Sum( solver.Sum( (coste_min[j] + solver.Sum(coste_extra[n][i] * (G[i][j][n] - min_prod[i]) for j in cantidad_generadores[i]))for i in tipos) +
                                solver.Sum(solver.Sum((1 - A[i][n-1]) * A[i][n] * coste_arranque[i] for j in cantidad_generadores[i]) for i in tipos) for n in franjas))
    
    
    
Problema()
