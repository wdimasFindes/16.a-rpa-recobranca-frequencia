from src.functions.RmAPI import RmAPI
from src.functions.ExcelFunctions import ExcelFunctions
from src.functions.MailFunctions import MailFunctions
from src.functions.SlackFunctions import SlackNotifier
from src.functions.SharepointFunctions import Sharepoint
from src.functions.Logger import LogGenerator
from src.functions.DatabaseFunctions import Database
import os
import sys
import pandas as pd
from datetime import datetime, timedelta
from dotenv import load_dotenv

load_dotenv()
class main:
    def __init__(self, slack_notifier, logger):
        self.baseDirectory = os.getcwd()
        self.dataDirectory = self.baseDirectory + "\\data\\tempFiles\\"
        self.slack = slack_notifier
        self.logger = logger

        #Limpa todos os arquivos da tempFiles
        try:
            tempFiles = os.listdir(self.dataDirectory)
            for file in tempFiles:
                    caminho_completo = os.path.join(self.dataDirectory, file)
                    if os.path.isfile(caminho_completo):
                        os.remove(caminho_completo)
        except:
            pass


    def main(self):
        rmApi = RmAPI(self.logger)

        dtInicio = (datetime.now() - timedelta(days=10)).strftime("%Y%m%d")
        dtFim = (datetime.now() - timedelta(days=4)).strftime("%Y%m%d")

        self.logger.info("Pegando dados da API do RM...")

        status, dfRM = rmApi.GetConsultaSQL(dtInicio, dtFim)

        if "Error" in status:
            self.logger.info("Erro ao pegar dados da API.")
            self.slack.post_message("Erro ao pegar dados da API.")
            sys.exit()

        excel = ExcelFunctions()
        sharepoint = Sharepoint()
        self.logger.info("Iniciando etapa de excel/emails.")
        try:
            self.logger.info("Lendo arquivo RM.")
            dfRM["STATUS"] = "PENDENTE"
            dfRM.rename(columns={'PROFESSOR': 'INSTRUTOR'}, inplace=True)

            dfInstrutores = excel.GetInstrutores(dfRM)

        except Exception as e:
             self.logger.info(f"Erro ao ler dataframe: {e}")
             self.slack.post_message(f"Error in get DataFrame: {e}")
             sys.exit()


        self.logger.info("Gerando token e-mail.")
        mail = MailFunctions()
        token = mail.GenerateToken()
        
        self.logger.info("Começando Loop instrutores")

        for index, row in dfInstrutores.iterrows():
            try:
                nomeInstrutor = row['INSTRUTOR']
                self.logger.info(f"Começando instrutor: {nomeInstrutor}")

                dfFiltradoLoop = dfRM.query(f'INSTRUTOR == "{nomeInstrutor}"')
                dfFiltradoLoop_all_columns = dfRM.query(f'INSTRUTOR == "{nomeInstrutor}"')
                colunas_email = ['UNIDADE', 'INSTRUTOR', 'EMAIL', 'CODPERLET', 'CODTURMA', 'DISCIPLINA', 'DATA', 'TURNO', 'AULA', 'FREQUENCIALIBERADA', 'CONTEUDOREALIZADO', 'CONTEUDOPREVISTO'] 
                dfFiltradoLoop = dfFiltradoLoop.filter(items=colunas_email)


                periodoLetivo = dfFiltradoLoop["CODPERLET"].iloc[0]
                turno = dfFiltradoLoop["TURNO"].iloc[0]

                self.logger.info("Lendo Email do supervisor e E-mail do instrutor")

                emailSupervisor = dfFiltradoLoop_all_columns["SUPIMED_EMAIL"].iloc[0]
                emailInstrutor = dfFiltradoLoop_all_columns["EMAIL"].iloc[0]

                
                self.logger.info("Lendo Email do orientador")

                emailOrientador = dfFiltradoLoop_all_columns["RESP_PED_EMAIL"].iloc[0]
                # print(f'emailInstrutor: {emailInstrutor}')
                # print(f'cc1: emailSupervisor: {emailSupervisor}')
                # print(f'cc2: emailOrientador: {emailOrientador}')
                
                dicData = [dfFiltradoLoop.to_dict(), dfFiltradoLoop.to_dict('index')]
                
                self.logger.info("Criando tabela HTML")
                htmlTable = excel.CreateHTMLTable(dicData)
                topEmail = """Prezado(a) Instrutor(a),
                                <br><br>
                                Conforme tabela abaixo, consta(m) pendência(s) de lançamento(s). 
                                Por gentileza, verificar a(s) coluna(s) preenchida(s) em amarelo. Favor efetuar a correção da(s) pendência(s).<br><br>"""
                body = topEmail + htmlTable

                # emailInstrutor = 'jhonata.alves@findes.org.br'
                # emailSupervisor ='jhonata.alves@findes.org.br'
                # emailOrientador = 'jhonata.alves@findes.org.br'

                self.logger.info("Enviando e-mail.")
                mail.SendMail(token, emailInstrutor, emailSupervisor, emailOrientador, body, nomeInstrutor)

                dfRM.loc[dfRM['INSTRUTOR'] == nomeInstrutor, 'STATUS'] = 'CONCLUÍDO'
                self.logger.info(f"""E-mail enviado com sucesso!
                                Instrutor: {nomeInstrutor},
                                Email instrutor: {emailInstrutor},
                                Email supervisor: {emailSupervisor},
                                Periodo Letivo: {periodoLetivo},
                                Turno: {turno}
                                 """)

            except Exception as e:
                self.logger.info(f"Erro no instrutor: {nomeInstrutor}. Motivo: {e}")
                self.slack.post_message(f"Erro no instrutor: {nomeInstrutor}, pulando e indo para o proximo.  Error: {e}")
                dfRM.loc[dfRM['INSTRUTOR'] == nomeInstrutor, 'STATUS'] = 'ERRO'
                continue
        

        database = Database()
        database.UploadDFToTable(dfRM)
        database.ExportToExcel(f"{self.dataDirectory}\\compilado_recobranca.xlsx")
        
        sharepoint.UploadFile(f"{self.dataDirectory}\\compilado_recobranca.xlsx")
        self.slack.post_message(f"Finalizado com sucesso, arquivo no Sharepoint.")



if __name__ == "__main__":
    slack_notifier = SlackNotifier(os.getenv("ENDPOINT_SLACK"), os.getenv("CHANNEL_SLACK"), os.getenv("NAME_ALERT"))
    log_instance = LogGenerator()
    logger = log_instance.setup_logger()
    
    start = main(slack_notifier, logger)
    start.main()