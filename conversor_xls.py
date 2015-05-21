# -*- encoding: utf-8 -*-
from Tkinter import *
from tkMessageBox import *
from FileDialog import *
import xlrd

def buscar_fila(cadena, libro):
    i = -1
    for f in libro.col(1):
        i +=1
        if cadena in unicode(f):
            return i
    return False

def test():
    """Simple test program."""
    root = Tk()
    root.withdraw()
    
    fd1 = LoadFileDialog(root)
    loadfile1 = fd1.go(key="DESGLOSE NOMINA")
    loadfile_name1 = loadfile1.split("/")[-1][:-4]
    empresa_cod =loadfile_name1[-3:] 
    libro1 = xlrd.open_workbook(loadfile1)
    libro1_hoja1 = libro1.sheet_by_index(0)
    
    fd2 = LoadFileDialog(root)
    loadfile2 = fd2.go(key="LISTADO DNI Y CENTRO DE TRABAJO")
    libro2 = xlrd.open_workbook(loadfile2)
    libro2_hoja1 = libro2.sheet_by_index(0)
    
    new_file = loadfile_name1 + "_convertido.txt"    
    
    with open(new_file, "w") as f:
        columna = -1
        for empleado in libro1_hoja1.row(3):
            empleado = unicode(empleado.value)
            columna += 1
            empl_cod = unicode((empleado)[:3])
            empl_name = unicode(empleado[8:].split(":")[0][:-5])
                        
            if empl_cod == empresa_cod:
                fila = -1
                for nomempl in libro2_hoja1.col(0):
                    fila += 1
                    aux = unicode(libro2_hoja1.cell_value(rowx=fila, colx=0))
                    if aux == empl_name:
                        dni = unicode(libro2_hoja1.cell_value(rowx=fila, colx=4)).strip()
                        dni = " " + dni
                        centro_trabajo = unicode(libro2_hoja1.cell_value(rowx=fila, colx=3))[-3:]
                        if empresa_cod == "119":
                            if centro_trabajo == "001":
                                cent_trab = "01"
                            else:
                                cent_trab = "00"
                        elif empresa_cod == "121" or empresa_cod == "120":
                            cent_trab = "00"
                        elif empresa_cod == "122":
                            if centro_trabajo == "005":
                                cent_trab = "01"
                            elif centro_trabajo == "003":
                                cent_trab = "02"
                            elif centro_trabajo == "004":
                                cent_trab = "04"
                            elif centro_trabajo == "001":
                                cent_trab = "05"
                            else:
                                cent_trab = "00"
                        else:
                            cent_trab = "00"
                        break
                
                cabecera = unicode(libro1_hoja1.cell_value(rowx=1, colx=1))
                if cabecera.find("ENERO") >= 0:
                    mes = "01"
                elif cabecera.find("FEBRERO") >= 0:
                    mes = "02"
                elif cabecera.find("MARZO") >= 0:
                    mes = "03"
                elif cabecera.find("ABRIL") >= 0:
                    mes = "04"
                elif cabecera.find("MAYO") >= 0:
                    mes = "05"
                elif cabecera.find("JUNIO") >= 0:
                    mes = "06"
                elif cabecera.find("JULIO") >= 0:
                    mes = "07"
                elif cabecera.find("AGOSTO") >= 0:
                    mes = "08"
                elif cabecera.find("SEPTIEMBRE") >= 0:
                    mes = "09"
                elif cabecera.find("OCTUBRE") >= 0:
                    mes = "10"
                elif cabecera.find("NOVIEMBRE") >= 0:
                    mes = "11"
                elif cabecera.find("DICIEMBRE") >= 0:
                    mes = "12"
                else:
                    mes = "00"
                    
                if empresa_cod == "119":
                    cod1 = "06060"
                elif empresa_cod == "120":
                    cod1 = "16326"
                elif empresa_cod == "121":
                    cod1 = "06059"
                elif empresa_cod == "122":
                    cod1 = "02113"
                else:
                    cod1 = "00000"
                f.write("\""+cod1+"\"")
                f.write("\t")
                
                f.write("\""+cent_trab+"\"")
                f.write("\t")
                
                f.write("\"2015"+mes+"N\"")
                f.write("\t")
                
                f.write("\""+dni+"\"")
                f.write("\t")
                
                f.write("\""+empl_name.encode('utf8')+"\"")
                f.write("\t")
                
                fila_base_irpf = buscar_fila("9908 BASE I.R.P.F.", libro1_hoja1)
                base_irpf = unicode(libro1_hoja1.cell_value(rowx=fila_base_irpf, colx=columna))
                if base_irpf == "":
                    base_irpf = "0.00"
                fila_incentivos = buscar_fila("3 INCENTIVOS 0001", libro1_hoja1)
                incentivos = unicode(libro1_hoja1.cell_value(rowx=fila_incentivos, colx=columna))
                if incentivos == "":
                    incentivos = "0.00"
                base = float(base_irpf) - float(incentivos)
                f.write(unicode(base))
                f.write("\t")
                
                fila_cuota_irpf = buscar_fila("9103 RETENCION I.R.P.F.", libro1_hoja1)
                cuota_irpf = unicode(libro1_hoja1.cell_value(rowx=fila_cuota_irpf, colx=columna))
                if cuota_irpf=="":
                    cuota_irpf = "0.00"
                f.write(cuota_irpf)
                f.write("\t")
                
                fila_base_irpf_esp = buscar_fila("37 S. ESPECIE 0013", libro1_hoja1)
                base_irpf_esp = unicode(libro1_hoja1.cell_value(rowx=fila_base_irpf_esp, colx=columna))
                if base_irpf_esp=="":
                    base_irpf_esp = "0.00"
                f.write(base_irpf_esp)
                f.write("\t")
                
                f.write("0.00")
                f.write("\t")
                
                fila_dietas = buscar_fila("9 DIETAS 0042", libro1_hoja1)
                dietas = unicode(libro1_hoja1.cell_value(rowx=fila_dietas, colx=columna))
                if dietas=="":
                    dietas = "0.00"
                f.write(dietas)
                f.write("\t")
                
                fila_prorratas = buscar_fila("9905 PRORRATA P. EXT.", libro1_hoja1)
                prorratas = unicode(libro1_hoja1.cell_value(rowx=fila_prorratas, colx=columna))
                if prorratas=="":
                    prorratas = "0.00"
                fila_prorratas2 = buscar_fila("1000 Parte P.P.Extras 0004", libro1_hoja1)
                prorratas2 = unicode(libro1_hoja1.cell_value(rowx=fila_prorratas2, colx=columna))
                if prorratas2=="":
                    prorratas2 = "0.00"
                prorr = float(prorratas) + float(prorratas2)
                f.write(unicode(prorr))
                f.write("\t")
                
                fila_total_desc = buscar_fila("TOTAL DESCUENTOS", libro1_hoja1)
                total_desc = unicode(libro1_hoja1.cell_value(rowx=fila_total_desc, colx=columna))
                if total_desc=="":
                    total_desc = "0.00"
                fila_reten_irpf = buscar_fila("9103 RETENCION I.R.P.F.", libro1_hoja1)
                reten_irpf = unicode(libro1_hoja1.cell_value(rowx=fila_reten_irpf, colx=columna))
                if reten_irpf=="":
                    reten_irpf = "0.00"  
                fila_anticipos = buscar_fila("9104 ANTICIPOS Y ESPECIES", libro1_hoja1)
                anticipos = unicode(libro1_hoja1.cell_value(rowx=fila_anticipos, colx=columna))
                if anticipos=="":
                    anticipos = "0.00"  
                ss_trab = float(total_desc) - float(reten_irpf) - float(anticipos)
                f.write(unicode(ss_trab))
                f.write("\t")
                
                fila_ss_emp = buscar_fila("9601 COSTE S.S. EMP", libro1_hoja1)
                ss_emp = unicode(libro1_hoja1.cell_value(rowx=fila_ss_emp, colx=columna))
                if ss_emp=="":
                    ss_emp = "0.00"
                f.write(ss_emp)
                f.write("\t")
                
                fila_liquido = buscar_fila("LIQUIDO", libro1_hoja1)
                liquido = unicode(libro1_hoja1.cell_value(rowx=fila_liquido, colx=columna))
                if liquido=="":
                    liquido = "0.00"
                f.write(liquido)
                f.write("\t")
                
                fila_kilometraje = buscar_fila("29 KILOMETRAJE 0050", libro1_hoja1)
                if not fila_kilometraje:
                    kilometraje = "0.00"
                else:
                    kilometraje = unicode(libro1_hoja1.cell_value(rowx=fila_kilometraje, colx=columna))
                f.write(kilometraje)
                f.write("\t")
                
                incentivos = unicode(libro1_hoja1.cell_value(rowx=fila_incentivos, colx=columna))
                if incentivos=="":
                    incentivos = "0.00"
                f.write(incentivos)
                f.write("\t")
                
                f.write("\n")
    f.close()


if __name__ == '__main__':
    test()