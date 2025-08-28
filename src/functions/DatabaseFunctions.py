import sqlite3
import os
import pandas as pd


class Database:
    def __init__(self):
        self.baseDirectory = os.getcwd()
        self.dataBaseDirectory = self.baseDirectory + "\\data\\database\\"
        self.CreateTable()


    def CreateTable(self):
        conn = sqlite3.connect(f"{self.dataBaseDirectory}compilado_recobranca.db")
        self.cursor = conn.cursor()

        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS compilado_robo_recobranca (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                DATA_EXECUCAO DATETIME DEFAULT CURRENT_TIMESTAMP,
                STATUS TEXT,
                CODCOLIGADA INTEGER,
                CODFILIAL INTEGER,
                IDTURMADISC INTEGER,
                CODTIPOCURSO INTEGER,
                CODPERLET TEXT,
                UNIDADE TEXT,
                CURSO TEXT,
                TURNO TEXT,
                CODTURMA TEXT,
                CODDISC TEXT,
                DISCIPLINA TEXT,
                INSTRUTOR TEXT,
                CPF TEXT,
                EMAIL TEXT,
                ETAPA TEXT,
                AULA INTEGER,
                FREQUENCIALIBERADA TEXT,
                DATA TEXT,
                HORARIO TEXT,
                CONTEUDOPREVISTO TEXT,
                CONTEUDOREALIZADO TEXT,
                QTD INTEGER)
                        ''')
        
        conn.commit()
        conn.close()
    
    def UploadDFToTable(self, df):
        conn = sqlite3.connect(f"{self.dataBaseDirectory}compilado_recobranca.db")
        columns_to_drop = ['CPF', 'SUPIMED', 'SUPIMED_EMAIL', 'SUPIMED_DTINICIAL', 'SUPIMED_DTFINAL', 
        'RESP_PED', 'RESP_PED_EMAIL', 'RESP_PED_M', 'RESP_PED_V', 'RESP_PED_N', 
        'RESP_PED_I']
        df = df.drop(columns=columns_to_drop)
        df.to_sql("compilado_robo_recobranca", conn, if_exists="append", index=False)
        conn.commit()
        conn.close()

    def ExportToExcel(self, pathToSave):
        conn = sqlite3.connect(f"{self.dataBaseDirectory}compilado_recobranca.db")
        df = pd.read_sql_query("SELECT * FROM compilado_robo_recobranca", conn)
        df['DATA_EXECUCAO'] = pd.to_datetime(df['DATA_EXECUCAO']).dt.strftime('%d/%m/%Y')

        try:
            os.remove(pathToSave)
        except:
            pass
    
        df.to_excel(pathToSave, index=False)
        conn.close()