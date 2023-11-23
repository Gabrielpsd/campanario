
from  PyPDF2 import PdfFileWriter
import fdb
import configparser
import traceback
import datetime
from openpyxl.workbook import Workbook
from openpyxl.styles import PatternFill, Alignment

ConfigName = "CONFIG.ini"

def creatReport(idClifor,cursor):
    safras = getSafras(cursor)
    return extractData(idClifor,safras,cursor) 
    
def extractData(idClifor,safras,cursor):
    command = """
    select distinct
    safra.descricao,
    arm_especie.descricao,
    arm_salclifor_fs.saldo_seco,
    arm_salclifor_fs.id_safra 
    from arm_salclifor_fs
    left join safra
        on (safra.id_safra =  arm_salclifor_fs.id_safra)
    left join clifor
        on (clifor.id = arm_salclifor_fs.id_clifor)
    left join arm_especie  
        on (arm_especie.id_especie = arm_salclifor_fs.id_especie)
    where arm_salclifor_fs.id_clifor = %d and arm_salclifor_fs.id_safra = %d
    order by arm_salclifor_fs.data desc
    """ 
    resultado = []
    
    for safra in safras:
        cursor.execute(command%(idClifor,safra[0]))
        
        for data in cursor:
            resultado.append(data)      
              
    return resultado

def creatExcel(Clifor,datas,cursor):
    fields = [Clifor,"Safra","Insumo", "Saldo"]
    safras = getSafras(cursor)
    date = datetime.datetime.now().strftime("%X").replace(':','_')
    row = 0
    column = 0
    wb = Workbook()
    
    work = wb.active
    work.title = Clifor
    work.auto_filter.ref = "B4"
    
    # the initial position of the name of the clifor will be in the 
    # B2 e C2 merged cell
    
    row = column = 2
    
    work.merge_cells(start_row=row,start_column=2,end_row=row,end_column =column +1)
    work.cell(row=row,column=column,value=Clifor)
    
    row = row + 3
    
    # each "safra" will be a tuple, bein the first positional argument the ID of the safra 
    # and the second posiotion argument the description of the safra
    for safra in safras:
        
        work.merge_cells(start_row=row,start_column=2,end_row=row,end_column =column +1)
        work.cell(row=row,column=column,value=safra[1])
        row = row + 1
        
        for data in datas:
            if (data[3] == safra[0]):
                #dadoInscrever = (data[1],int(data[2]))
                work.cell(row=row,column=column,value=data[1])
                work.cell(row=row,column=column+1,value=f"{data[2]} Kg")
                row = row + 1
                
        row = row + 1 
    savePath = getSavePath()
    wb.save(f"{savePath}\\saldo_{date}.xlsx")
        

def getSafras(cursor):
    command = """
        select id_safra , descricao from safra
    """   
    try:
        cursor.execute(command)
    except:
        raise Exception("Erro ao executar comando GetSafra")
    else: 
        safras = []
        for safra in cursor:
            safras.append(safra)
        return safras
 
# Read the INI file and get the path to the database   
def getPath():
    config = configparser.ConfigParser()
    config.sections() 
    config.read(ConfigName)
    
    DataBasePath = config["DATAPTH"]["DatabasePath"]
    return DataBasePath

def getSavePath():
    config = configparser.ConfigParser()
    config.sections() 
    config.read(ConfigName)
    
    DataSavePath = config["DATAPTH"]["SavePath"]
    return DataSavePath

def searchClifor(id_clifor,con):
    command = """
    select nome from clifor
    where id ="""
    
    con.execute(f"{command}{id_clifor}")
    
    # the return of fetchall is a list of tuples so we must get the position  
    name = con.fetchall()
    
    # returned a empty name we filter the error
    try:
        name = name[0][0]
    except IndexError:
       raise IndexError("Id do Clifor não encontrado")
    except Exception:
        raise Exception("Erro nao catalogado")
    else:
        return name
    
def main():
    dataBasePath = getPath()
    connection = fdb.connect(dsn= f"{dataBasePath}",password="masterkey",user="SYSDBA")
    cursor = connection.cursor()
    
    print(dataBasePath)
    print("Saldo Total de um produtor")
    
    while(1):
        
        id_clifor = int(input("Digite o ID do clifor:"))
        
        try:
            cliforName = searchClifor(id_clifor,cursor)
        except IndexError as value:
            print(traceback.format_exception_only(type(value),value)[0])
            continue
        except Exception as error:
            print(traceback.format_exception_only(type(error),error))
           
        else:
            print(f"Criar relatorio para o {cliforName}")
            opcao = int(input("1- Sim    2- Não  3- sair \n Digite a opcao: "))
            
            if opcao == 1:
                report = creatReport(id_clifor, cursor)
                print(report)
                creatExcel(cliforName,report,cursor)
                connection.close()
                break
            elif opcao == 3:
                connection.close()
                exit()
        
if "__main__" == __name__:
    main()
