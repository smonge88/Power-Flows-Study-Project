def flujodecarga():
    import math as mt
    import numpy as np
    import pandas as pd
    from numpy import matrix
    datos=pd.read_excel('Datos de las líneas.xlsx')#lee los datos de las barras#
    barras=pd.read_excel('Datos de las barras.xlsx')
    n1=matrix(datos.iloc[:,0])#importa datos de barra i y los hace matriz#
    n2=matrix(datos.iloc[:,1])#importa datos de barra j y los hace matriz#
    r=matrix(datos.iloc[:,2])#importa datos de resistencia y los hace matriz#
    x=matrix(datos.iloc[:,3])#importa datos de reactancia y los hace matriz#
    b=matrix(datos.iloc[:,4])#importa datos de suceptancia y los hace matriz#
    tr=matrix(datos.iloc[:,5])#importa datos de relación de transformador y los hace matriz#
    com=complex(0,1)#numero complejo#
    b=b*com#define b como numero complejo 
    shunt=matrix(barras.iloc[:,10])
    vbase=matrix(barras.iloc[:,11])#importa datos de relación de transformador y los hace matriz#
    
   
    print("Ingrese la potencia base en MVA")
    sbase=int(input())
    print("Ingrese el porcentaje de error")
    perror=float(input())
    print("Ingrese el el voltaje base en la zona de alto voltaje en kV")
    hv=float(input())
    print("Ingrese el el voltaje base en la zona de medio voltaje en kV")
    mv=float(input())
    print("Ingrese el el voltaje base en la zona de bajo voltaje en kV")
    lv=float(input())
    s=shunt.shape
    
    for i in range(s[1]):
        if vbase[0,i] in ['HV']:
            vbase[0,i]=hv
        elif vbase[0,i] in ['MV']:
            vbase[0,i]=mv
        elif vbase[0,i] in ['LV']:
            vbase[0,i]=lv
    
    #####################################################
    ########### creacion de la ybarra####################
    #####################################################

    s=n1.shape#define el numero de filas para cada columna#
    z=np.zeros((1,s[1]),dtype=np.complex)#realiza matriz de 0s complejos para crear la matriz de impedancia#
    
    for i in range(s[1]):#crea la matriz de impedancias#
        if tr[0,i]==0:
            tr[0,i]=1 
        z[0,i-1]=complex(r[0,i-1],x[0,i-1])#se agrega la reactancia y resistencia de cada fila#
    n=0#variable contadora de numero de barras#
    for i in range (s[1]):#busca el numero de barras totales#
        if n1[0,i-1]>n:#recorre columna barra i#
            n=n1[0,i-1]#cambia el valor n de numero de barras si encuentra una barra mayor#
        if n2[0,i-1]>n:#recorre columna barra j#
            n=n2[0,i-1]#cambia el valor n de numero de barras si encuentra una barra mayor#
   
    ybus=np.zeros((n,n),dtype=np.complex)#crea matriz nxn con el numero de barras#
    for i in range (s[1]):#coomienza a crear y bus trasladandose por las filas llenando fuera de la diagonal#
                ybus[n1[0,i]-1,n2[0,i]-1]=-tr[0,i]*(1/(z[0,i]))#agrega admitancia de conexión fuera de diagonal#
                ybus[n2[0,i]-1,n1[0,i]-1]=-tr[0,i]*(1/(z[0,i]))#agrega admitancia de conexión fuera de diagonal a su reciproco#
    
    for i in range (n):#recorre toda la diagonal#
        for j in range(s[1]):#recorre todas las filas de lasmatrices#
            if n1[0,j]==i+1:#si la barra corresponde a la diagonal#
                
                ybus[i,i]=ybus[i,i]+((1/(z[0,j]))+(b[0,j])/2)#agrega los miembros conectados y la admitancia en paralelo#
            if n2[0,j]==i+1:#si la barra corresponde a la diagonal#
               ybus[i,i]=ybus[i,i]+((1/(z[0,j]))*(tr[0,j]**2)+(b[0,j])/2)#agrega los miembros conectados y la admitancia en paralelo#
        ybus[i,i]=ybus[i,i]+complex(0,shunt[0,i])#agrega compensacion shunt
    
    ybus=ybus*(100/sbase)
            
    excelybus=pd.DataFrame(ybus)#convierte la matriz en un dataframe  
    excelybus.to_excel("Matriz Ybus.xlsx",sheet_name="Ybus", index_label="Barra", startrow=0, startcol=0)#escribe los datos de ybarra en un archivo excel
    #####################################################
    ########### Gauss Seidel ############################
    #####################################################
    
    
    b1=matrix(barras.iloc[:,0])#lee el numero de barra
    b2=matrix(barras.iloc[:,1])#lee el tipo de barra
    s=b1.shape#determina el tamaño de filas (numero de barras)
    tipo=np.zeros((s[1],12),dtype=float)#define una matriz de 0 tipo flotante  con 0 destinada para guardar los datos de flujo de potencia
   
    for i in range(s[1]):#ciclo destinado a agregar las variables para cada barra 
        tipo[i,0]=b1[0,i]#asigna el numero de barra 
        if b2[0,i] in ['Referencia ']:#solo contempla barras de tipo de referencia
            tipo[i,1]=1#se asigna un valor de 1 a las barras de referencia 
            tipo[i,2]=barras.iloc[i,2]#lee el valor de voltaje de la barra y lo agrega a la matriz 
            tipo[i,3]=barras.iloc[i,3]#lee el valor de angulo de voltaje de la barra y lo agrega a la matriz 
            tipo[i,4]=barras.iloc[i,4]#lee el valor de carga activa de la barra y lo agrega a la matriz 
            tipo[i,5]=barras.iloc[i,5]#lee el valor de carga reactiva de la barra y lo agrega a la matriz 
            tipo[i,8]=barras.iloc[i,8]#lee el valor de qmax de generacióny lo agrega a la matriz 
            tipo[i,9]=barras.iloc[i,9]#lee el valor de qmin de generación de la barra y lo agrega a la matriz 
            
            
        if b2[0,i] in ['Control (PV)']:#solo contempla barras de tipo de control
            tipo[i,1]=2#se asigna un valor de 2 a las barras de control 
            tipo[i,2]=barras.iloc[i,2]#lee el valor de voltaje de la barra y lo agrega a la matriz 
            tipo[i,3]=0#se asigna un valor semilla de 0 en angulo para iniciar la iteración
            tipo[i,4]=barras.iloc[i,4]#lee el valor de carga activa de la barra y lo agrega a la matriz 
            tipo[i,5]=barras.iloc[i,5]#lee el valor de carga reactiva de la barra y lo agrega a la matriz 
            if barras.iloc[i,6] in ['?']:#si el valor de generación en MW es desconocido se asigna un 0 y se agrega a la matriz 
                tipo[i,6]=0#se agrega el0
            else:
                tipo[i,6]=barras.iloc[i,6]#sino se agrega el valor de generacion
            tipo[i,8]=barras.iloc[i,8]#lee el valor de qmax de generación y lo agrega a la matriz 
            tipo[i,9]=barras.iloc[i,9]#lee el valor de qmin de generación de la barra y lo agrega a la matriz 
      
            
        if b2[0,i] in ['Carga (PQ)']:#solo contempla barras de tipo de control
            tipo[i,1]=3#se asigna un valor de 3 a las barras de carga 
            tipo[i,2]=1#se asigna un valor semilla de 1 en V para iniciar la iteración
            tipo[i,3]=0#se asigna un valor semilla de 0 en angulo para iniciar la iteración
            tipo[i,4]=barras.iloc[i,4]#lee el valor de carga activa de la barra y lo agrega a la matriz 
            tipo[i,5]=barras.iloc[i,5]#lee el valor de carga reactiva de la barra y lo agrega a la matriz 
            tipo[i,6]=barras.iloc[i,6]#lee el valor de generaccion de MW de la barra y lo agrega a la matriz 
            tipo[i,7]=barras.iloc[i,7]#lee el valor de generaccion de MVAR de la barra y lo agrega a la matriz 
            tipo[i,8]=barras.iloc[i,8]#lee el valor de qmax de generacióny lo agrega a la matriz 
            tipo[i,9]=barras.iloc[i,9]#lee el valor de qmin de generación de la barra y lo agrega a la matriz 
            
    v=np.zeros((s[1],1),dtype=np.complex)#define una matriz de 0 complejos para el voltaje de cada barra 
    scomp=np.zeros((s[1],1),dtype=np.complex)#define una matriz de 0 complejos para la potencia en las iteraciones de cada barra 
    for i in range(s[1]):#toma los datos de voltaje y angulo para transformarlo en conordenadas polares
        v[i,0]=complex(tipo[i,2]*mt.cos(mt.radians(tipo[i,3])), tipo[i,2]*mt.sin(mt.radians(tipo[i,3])))#transforma el voltaje de numero polar a complejo rectangular 
        scomp[i,0]=((complex((-tipo[i,4]+tipo[i,6]),-(-tipo[i,5]+tipo[i,7])))/sbase)#toma los valores de potencia reactiva y activa y los muestra en forma compleja y los transforma en PU
    
    error=100#porcentaje de error por defecto para iniciar el ciclo 
    itera=0#contador de iteraciones 
    
    while error>perror:#realiza el ciclo buscando un cierto porcentaje de error 
        itera=itera+1#determina el numero de iteraciones 
        error=0#variable para guardar el  porcentaje de error 
        for i in range(s[1]):#inicia la iteración de flujos de potencia
            ang=0
            if tipo[i,1]==2:#barra PV, Cálculo de Q
                suma=0#variable para guardar las Y*V
                contador=ybus.shape#Contador para recorrer las columnas de la ybarra
                for y in range(contador[0]):#recorre las columnas de la y barra
                    suma=suma+ybus[i,y]*v[y,0]#multiplica y suma los YV
                q=v[i,0].conjugate()*suma#se calcula una potencia para la iteración de acuerdo con la formula 
                q=-q.imag#se toma la parte imaginaria de la potencia para esa iteración
                
                if (q+(tipo[i,5]/sbase))>(tipo[i,8]/sbase):#cambia la potencia reactiva en caso de que  sea más grande que la potencia máxima 
                    q=(-tipo[i,5]+tipo[i,8])/sbase#cambia a la potencia reactiva máxima 
            
                elif (q+(tipo[i,5]/sbase))<(tipo[i,9]/sbase):#cambia la potencia reactiva en caso de que  sea más pequeña que la potencia mínima  
                    q=(-tipo[i,5]+tipo[i,9])/sbase#cambia a la potencia reactiva mínima  
                
                scomp[i,0]=complex(scomp[i,0].real,-q)#se sustituye q por el q nuevo 
                vs=(1/ybus[i,i])*((scomp[i,0]/v[i,0].conjugate())-(suma-(ybus[i,i]*v[i,0])))#se calcula el voltaje
                v[i,0]=abs(v[i,0])*(vs/abs(vs))#se corrije para obtener la magnitud requerida 
                ang=mt.degrees(mt.atan((v[i,0].imag)/(v[i,0].real)))
                tipo[i,7]=q*sbase+tipo[i,5]
                
            elif tipo[i,1]==3:#si la barra es PQ se calcula el voltaje de la barra 
                suma=0#variable para guardar las Y*V
                vant=abs(v[i,0])#voltaje para calcular el porcentaje de error
                contador=ybus.shape#Contador para recorrer las columnas de la ybarra
                
                for y in range(contador[0]):#recorre las columnas de la y barra
                    suma=suma+ybus[i,y]*v[y,0]#multiplica y suma los YV
                
                v[i,0]=(1/ybus[i,i])*((scomp[i,0]/v[i,0].conjugate())-(suma-(ybus[i,i]*v[i,0])))#se calcula el voltaje nuevo
                ang=mt.degrees(mt.atan((v[i,0].imag)/(v[i,0].real)))
                
                errornuevo=(abs(vant-abs(v[i,0]))/abs(v[i,0]))*100#se calcula el porcentaje de error nuevo
                if errornuevo>error: #busca definir el procentaje de error más alto entre las barras 
                    error=errornuevo#si es más alto se cambia
            
            
                    
            
            tipo[i,2]=abs(v[i,0])
            tipo[i,3]=ang
            
    
    suma=0#variable para guardar las Y*V
    contador=ybus.shape#Contador para recorrer las columnas de la ybarra
    for y in range(contador[0]):#recorre las columnas de la y barra
        suma=suma+ybus[0,y]*v[y,0]#multiplica y suma los YV
    
    scomp[0,0]=v[0,0]*suma+complex(tipo[0,4]/sbase,-tipo[0,5]/sbase)
    
    tipo[0,6]=scomp[0,0].real*sbase
    tipo[0,7]=-scomp[0,0].imag*sbase
    print("Se realizaron",itera, "iteraciones")

    tipo[:,11]=tipo[:,9]
    tipo[:,10]=tipo[:,8]
    tipo[:,9]=tipo[:,7]
    tipo[:,8]=tipo[:,6]        
    tipo[:,7]=tipo[:,5]
    tipo[:,6]=tipo[:,4]
    tipo[:,5]=tipo[:,3]
    
    
    s=vbase.shape
    
    
    for i in range(s[1]):
        tipo[i,3]=vbase[0,i]
        tipo[i,4]=tipo[i,2]*vbase[0,i]

    exceltipo=pd.DataFrame(tipo)#convierte la matriz en un dataframe
    exceltipo.iloc[:,1]=barras.iloc[:,1]
    exceltipo.to_excel("Perfiles de tensión y potencia de las barras.xlsx",sheet_name="buses",header=["Barra","Tipo", "Voltaje (pu)","Voltaje base (kV)","Voltaje (kV)","Angulo","Carga (MW)","Carga (MVar)","Generación (MW)","Generación(MVar)","Q max gen (MVAR)","Q min gen (MVAR)"], index=False)#escribe los datos de ybarra en un archivo excel
    
    
    
    ###################################
    #######flujos de potencia##########
    ###################################
    s=n1.shape#define el numero de filas para cada columna#
    flujos=np.zeros((2*s[1],6),dtype=np.float)
    

    for i in range (s[1]):
        flujos[2*i,0]=(n1[0,i])
        flujos[2*i,1]=(n2[0,i])
        flujos[2*i+1,0]=(n2[0,i])
        flujos[2*i+1,1]=(n1[0,i])
    
    for i in range (s[1]):
        b1=n1[0,i]
        b2=n2[0,i]
        fp=(v[b1-1,0]*(v[b1-1,0].conjugate()-v[b2-1,0].conjugate())*((1/z[0,i])*(100/sbase)).conjugate())+(v[b1-1,0]*v[b1-1,0].conjugate()*((b[0,i].conjugate())/2)*(100/sbase))
        flujos[2*i,2]=fp.real
        flujos[2*i,3]=fp.imag
        flujos[2*i,4]=fp.real*sbase
        flujos[2*i,5]=fp.imag*sbase
        fp=(v[b2-1,0]*(v[b2-1,0].conjugate()-v[b1-1,0].conjugate())*((1/z[0,i])*(100/sbase)).conjugate())+(v[b2-1,0]*v[b2-1,0].conjugate()*((b[0,i].conjugate())/2)*(100/sbase))
        flujos[2*i+1,2]=fp.real
        flujos[2*i+1,3]=fp.imag
        flujos[2*i+1,4]=fp.real*sbase
        flujos[2*i+1,5]=fp.imag*sbase
        
        
     
    
    
    
    
    excelflujos=pd.DataFrame(flujos)
    excelflujos.to_excel("Flujos de potencia.xlsx",sheet_name="Flujos",header=["Desde barra","Hasta Barra", "Potencia real (pu)","Potencia reactiva(pu)","Potencia real (MW)","Potencia reactiva(MVar)"], index=False)#escribe los datos de ybarra en un archivo excel
    
   
                    
    