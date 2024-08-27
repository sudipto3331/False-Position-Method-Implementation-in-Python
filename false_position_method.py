#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Mar 19 10:03:56 2023

@author: sudipto3331
"""
# -*- coding: utf-8 -*-
"""

"""
import math
import numpy as np
from xlwt import Workbook
import xlwt
from xlrd import open_workbook
from xlutils.copy import copy

def fnc(x):
    return (667.38/x)*(1-math.exp(-0.146843*x))-40

def bisection(fxl, fxu, err, ite):
    x_l=np.zeros([ite])
    x_u=np.zeros([ite])
    x_c=np.zeros([ite])
    
    f_xl=np.zeros([ite])
    f_xu=np.zeros([ite])
    f_xc=np.zeros([ite])
    
    rel_err=np.zeros([ite])
    itern=np.zeros([ite])
    x_l[0]=xl
    x_u[0]=xu
    
    f_xl[0]=fxl
    f_xu[0]=fxu 

    for i in range(ite):
        #storing the values of iteration
        itern[i]=i+1
        #Bisection Formula
        x_c[i]=(x_l[i]+x_u[i])/2
        
        f_xl[i]=fnc(x_l[i])
        f_xu[i]=fnc(x_u[i])
        f_xc[i]=fnc(x_c[i])
        #calculating error    
        if i>0:
            rel_err[i]=((x_c[i]-x_c[i-1])/x_c[i])*100
        #terminate if error criteria meets
        if all ([i>0, abs(rel_err[i])<err]):
            temp = i;
            break 
        elif f_xc[i]==0:
            temp = i;
            break
   
        if i==ite-1:
            temp = i;
            break
        #replacement of the new estimate
        if all ([f_xc[i]>0, f_xl[i]>0]):
            x_l[i+1]=x_c[i]
            x_u[i+1]=x_u[i]
        elif all ([f_xc[i]>0, f_xu[i]>0]):
            x_u[i+1]=x_c[i]
            x_l[i+1]=x_l[i]
        elif all ([f_xc[i]<0, f_xl[i]<0]):
            x_l[i+1]=x_c[i]
            x_u[i+1]=x_u[i]
        elif all ([f_xc[i]<0, f_xu[i]<0]):
            x_u[i+1]=x_c[i]
            x_l[i+1]=x_l[i]
    

    wb = Workbook()

    sheet1 = wb.add_sheet('Sheet 1')
    num_of_iter=i

    sheet1.write(0,3,'Bisection')
    sheet1.write(0,4,'Method')


    sheet1.write(1,0,'Number of iteration')
    sheet1.write(1,1,'x_l')
    sheet1.write(1,2,'x_u')
    sheet1.write(1,3,'f(x_l)')
    sheet1.write(1,4,'f(x_u)')
    sheet1.write(1,5,'x_c')
    sheet1.write(1,6,'f(x_c)')
    sheet1.write(1,7,'Relative error')
    

    for n in range(num_of_iter+1):
        
        sheet1.write(n+2,0,itern[n])
        sheet1.write(n+2,1,x_l[n])
        sheet1.write(n+2,2,x_u[n])
        sheet1.write(n+2,3,f_xl[n])
        sheet1.write(n+2,4,f_xu[n])
        sheet1.write(n+2,5,x_c[n])
        sheet1.write(n+2,6,f_xc[n])
        sheet1.write(n+2,7,rel_err[n])
    
    sheet1.write(n+4,2,'The')
    sheet1.write(n+4,3,'root')
    sheet1.write(n+4,4,'is')
    sheet1.write(n+4,5,x_c[i])
    wb.save('LAB2.xls')
        
        
    return temp;


def fasleposition(fxl, fxu, err, ite, index):
    x_l=np.zeros([ite])
    x_u=np.zeros([ite])
    x_c=np.zeros([ite])
    
    f_xl=np.zeros([ite])
    f_xu=np.zeros([ite])
    f_xc=np.zeros([ite])
    
    rel_err=np.zeros([ite])
    itern=np.zeros([ite])
    
    x_l[0]=xl
    x_u[0]=xu
    
    f_xl[0]=fxl
    f_xu[0]=fxu 
    #begin iteration   
    for i in range(ite):
        #storing the values of iteration
        itern[i]=i+1
        #Bisection Formula
        
        
        f_xl[i]=fnc(x_l[i])
        f_xu[i]=fnc(x_u[i])
        x_c[i]=((x_u[i]*f_xl[i])-(x_l[i]*f_xu[i]))/(f_xl[i]-f_xu[i])
        f_xc[i]=fnc(x_c[i])
        
        
        #calculating error    
        if i>0:
            rel_err[i]=((x_c[i]-x_c[i-1])/x_c[i])*100
        #terminate if error criteria meets
        if all ([i>0, abs(rel_err[i])<err]):
            temp = i;
            break 
        elif f_xc[i]==0:
            temp = i;
            break
   
        if i==ite-1:
            temp = i;
            break
        #replacement of the new estimate
        if all ([f_xc[i]>0, f_xl[i]>0]):
            x_l[i+1]=x_c[i]
            x_u[i+1]=x_u[i]
        elif all ([f_xc[i]>0, f_xu[i]>0]):
            x_u[i+1]=x_c[i]
            x_l[i+1]=x_l[i]
        elif all ([f_xc[i]<0, f_xl[i]<0]):
            x_l[i+1]=x_c[i]
            x_u[i+1]=x_u[i]
        elif all ([f_xc[i]<0, f_xu[i]<0]):
            x_u[i+1]=x_c[i]
            x_l[i+1]=x_l[i]
    

    rb = open_workbook("LAB2.xls")
    wb = copy(rb)
    sheet1 = wb.get_sheet(0)
    
    num_of_iter=i
    
    
    sheet1.write(index,3,'False')
    sheet1.write(index,4,'Position')
    sheet1.write(index,5,'Method')

    sheet1.write(index+1,0,'Number of iteration')
    sheet1.write(index+1,1,'x_l')
    sheet1.write(index+1,2,'x_u')
    sheet1.write(index+1,3,'f(x_l)')
    sheet1.write(index+1,4,'f(x_u)')
    sheet1.write(index+1,5,'x_c')
    sheet1.write(index+1,6,'f(x_c)')
    sheet1.write(index+1,7,'Relative error')
    

    for n in range(num_of_iter+1):
        
        sheet1.write(index+n+2,0,itern[n])
        sheet1.write(index+n+2,1,x_l[n])
        sheet1.write(index+n+2,2,x_u[n])
        sheet1.write(index+n+2,3,f_xl[n])
        sheet1.write(index+n+2,4,f_xu[n])
        sheet1.write(index+n+2,5,x_c[n])
        sheet1.write(index+n+2,6,f_xc[n])
        sheet1.write(index+n+2,7,rel_err[n])
    
    sheet1.write(index+n+4,2,'The')
    sheet1.write(index+n+4,3,'root')
    sheet1.write(index+n+4,4,'is')
    sheet1.write(index+n+4,5,x_c[i])
    wb.save('LAB2.xls')
    return temp;
    

xl=np.float(input ('Enter 1st initial value: '))   
xu=float(input ('Enter 2nd initial value: ')) 
  

fxl=fnc(xl)
fxu=fnc(xu)

if fxl*fxu>0:
    print('Wrong initial input')
elif fxl*fxu<0:
    err=float(input('Enter desired percentage relative error: '))
    ite=int(input('Enter number of iterations: '))
    index = bisection(fxl, fxu, err, ite)
    index = index+6
    fasleposition(fxl, fxu, err, ite, index)
    
print("Task Successfull");

  
rb = open_workbook("LAB2.xls")
wb = copy(rb)

sheet1 = wb.get_sheet(0)

sheet1.merge(40, 40, 3, 4)
sheet1.write(40,3,'Bisection Method',xlwt.easyxf("font: bold 1,height 250; align: horiz center"))


wb.save('LAB2.xls')
