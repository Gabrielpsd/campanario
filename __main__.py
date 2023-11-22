
from  PyPDF2 import PdfFileWriter
import fdb
import configparser
import traceback

ConfigName = "CONFIG.ini"

def creatReport(idClifor,cursor):
    safras = getSafras(cursor)
    extractData(idClifor,safras,cursor) 
    
def extractData(idClifor,safras,cursor):
    command = """
    select distinct
    safra.descricao,
    clifor.nome,
    arm_especie.descricao,
    arm_salclifor_fs.saldo_seco
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
    
    print("--------------- resultado das buscas ------------------ ")
    print(resultado)
       
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
                creatReport(id_clifor, cursor)
                break
            elif opcao == 3:
                exit()
        
if "__main__" == __name__:
    main()
