import base64
from PIL import Image
import io
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
DEV_MODE = os.getenv("DEV_MODE", "False").lower() == "true"

class main:
    def __init__(self, slack_notifier, logger):
        self.baseDirectory = os.getcwd()
        self.dataDirectory = self.baseDirectory + "\\data\\tempFiles\\"
        self.slack = slack_notifier
        self.logger = logger
        # Carregar o caminho da imagem do .env
        self.image_path = os.getenv('IMAGE_PATH')
        self.image_resized_base64 =None  # Obtém o caminho da imagem da variável do .env

        if self.image_path:
            print(f"Imagem será carregada de: {self.image_path}")
            self.logger.info("Imagem de Assinatura encontrada e carregada")
        else:
            print("Erro: Caminho da imagem não encontrado no .env")

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

        if DEV_MODE:
            dfRM = dfRM.assign(
                SUPIMED_EMAIL="wescley@findes.org.br",
                RESP_PED_EMAIL="wescley@findes.org.br"
                #SUPIMED_EMAIL="jbravo@findes.org.br",
                #RESP_PED_EMAIL="jangelica@findes.org.br"
            )

        if "Error" in status:
            self.logger.info("Erro ao pegar dados da API.")
            self.slack.post_message("Erro ao pegar dados da API.")
            sys.exit()

        excel = ExcelFunctions()
        sharepoint = Sharepoint(self.logger)
        self.logger.info("Iniciando etapa de excel/emails.")
        try:
            self.logger.info("Lendo arquivo RM.")
            dfRM["STATUS"] = "PENDENTE"
            #dfRM.rename(columns={'PROFESSOR': 'INSTRUTOR'}, inplace=True)

            #Filtrando apenas para a FILIAL 4 - piloto
            dfFiltradoFilial = dfRM.query('CODFILIAL == 4')
            dfRM = dfFiltradoFilial  

            dfInstrutores = excel.GetInstrutores(dfRM)

        except Exception as e:
             self.logger.info(f"Erro ao ler dataframe: {e}")
             self.slack.post_message(f"Error in get DataFrame: {e}")
             sys.exit()


        self.logger.info("Gerando token e-mail.")
        mail = MailFunctions()
        token = mail.GenerateToken()
        
        self.logger.info("Começando Loop instrutores")

        #for index, row in dfInstrutores.iterrows():
        for idx, (index, row) in enumerate(dfInstrutores.iterrows()):
            if idx >= 1:  # Limita o loop a 3 execuções
                break
            try:
                nomeInstrutor = row['PROFESSOR']
                self.logger.info(f"Começando instrutor: {nomeInstrutor}")

                dfFiltradoLoop = dfRM.query(f'PROFESSOR == "{nomeInstrutor}"')
                dfFiltradoLoop_all_columns = dfRM.query(f'PROFESSOR == "{nomeInstrutor}"')
                colunas_email = ['UNIDADE', 'PROFESSOR', 'EMAIL', 'CODPERLET', 'CODTURMA', 'DISCIPLINA', 'DATA', 'TURNO', 'AULA', 'FREQUENCIALIBERADA', 'CONTEUDOREALIZADO', 'CONTEUDOPREVISTO'] 
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
                topEmail = """Prezado(a) Professor(a),
                                <br><br>
                                Conforme tabela abaixo, consta(m) pendência(s) de lançamento(s). 
                                Por gentileza, verificar a(s) coluna(s) preenchida(s) em amarelo. Favor efetuar a correção da(s) pendência(s).<br><br>"""
                
                body = topEmail + htmlTable

                # Carregar e redimensionar a imagem
                if self.image_path and os.path.exists(self.image_path):
                    with Image.open(self.image_path) as img:
                        # Verificar se a imagem é JPG ou PNG
                        if img.format not in ['JPEG', 'PNG']:
                            raise ValueError("O arquivo de imagem não é um JPG ou PNG válido.")
                        
                        # Redimensionar a imagem para 740x220
                        img_resized = img.resize((500, 200), Image.Resampling.LANCZOS)

                        # Salvar a imagem redimensionada em um buffer de memória (em formato JPG ou PNG)
                        img_byte_arr = io.BytesIO()
                        if img.format == 'JPEG':
                            img_resized.save(img_byte_arr, format="JPEG")  # Salvar como JPG
                        else:
                            img_resized.save(img_byte_arr, format="PNG")  # Salvar como PNG
                        img_byte_arr = img_byte_arr.getvalue()
                        
                        # Codificar a imagem em base64
                        img_base64 = base64.b64encode(img_byte_arr).decode('utf-8')
                        
                        # Atualizar a variável de imagem codificada
                        self.image_resized_base64 = img_base64                   


                    assinaturaEmail = f"""<p><strong>Atenciosamente,</strong></p>
                    <p><strong>Pedagogico Bot</strong></p>
                    <p>Equipe de Validação</p>
                    
                    <!-- Inserir a assinatura com a imagem fixa à esquerda -->
                    <div style="text-align: left;">
                    <p><strong>Classificação: Interno</strong></p> <!-- Informar que o e-mail é interno -->

                    <img src="data:image/{img.format.lower()};base64,{self.image_resized_base64}" alt="Assinatura" style="display: block;" />
                    </div>
                    <!-- Adiciona uma linha de espaço (escape) -->
                    <p> </p>
                    <!-- Espaço maior com margem -->
                    <p style="margin-top: 20px;"> </p>

                    <!-- Centralizar Visão e Missão -->
                    <div style="text-align: left;">
                    <p><strong>Visão:</strong> “Ser referência como Departamento Econômico entre as Federações de Indústria até 2030”</p>
                    <p><strong>Missão:</strong> “Fortalecer o desenvolvimento da indústria do Estado do Espírito Santo por meio de pesquisas, estudos e análises de dados”</p>
                    </div>
                    </body>
                    </html>
                    """

                    body = body + assinaturaEmail
                    

                emailInstrutor = 'wescleydecarvalho@hotmail.com'
                emailSupervisor ='wescley@findes.org.br'
                emailOrientador = 'wescley@findes.org.br'

                self.logger.info("Enviando e-mail.")
                mail.SendMail(token, emailInstrutor, emailSupervisor, emailOrientador, body, nomeInstrutor)

                dfRM.loc[dfRM['PROFESSOR'] == nomeInstrutor, 'STATUS'] = 'CONCLUÍDO'
                self.logger.info(f"""E-mail enviado com sucesso!
                                Professor: {nomeInstrutor},
                                Email professor: {emailInstrutor},
                                Email supervisor: {emailSupervisor},
                                Periodo Letivo: {periodoLetivo},
                                Turno: {turno}
                                 """)

            except Exception as e:
                self.logger.info(f"Erro no instrutor: {nomeInstrutor}. Motivo: {e}")
                self.slack.post_message(f"Erro no instrutor: {nomeInstrutor}, pulando e indo para o proximo.  Error: {e}")
                dfRM.loc[dfRM['PROFESSOR'] == nomeInstrutor, 'STATUS'] = 'ERRO'
                continue
        

        database = Database()
        database.UploadDFToTable(dfRM)
        database.ExportToExcel(f"{self.dataDirectory}compilado_recobranca.xlsx")
        
        sharepoint.UploadFile(f"{self.dataDirectory}compilado_recobranca.xlsx")
        self.slack.post_message(f"Finalizado com sucesso, arquivo no Sharepoint.")



if __name__ == "__main__":
    slack_notifier = SlackNotifier(os.getenv("ENDPOINT_SLACK"), os.getenv("CHANNEL_SLACK"), os.getenv("NAME_ALERT"))
    log_instance = LogGenerator()
    logger = log_instance.setup_logger()
    
    start = main(slack_notifier, logger)
    start.main()