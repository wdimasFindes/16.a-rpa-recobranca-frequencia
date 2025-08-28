# import os
# from office365.sharepoint.client_context import ClientContext
# from office365.runtime.auth.user_credential import UserCredential
# from datetime import datetime
# from dotenv import load_dotenv

# load_dotenv()
# class Sharepoint:
#     def __init__(self):
#         self.SHAREPOINT_EMAIL = os.getenv('SHAREPOINT_EMAIL')
#         self.SHAREPOINT_PASSWORD = os.getenv('SHAREPOINT_PASSWORD')
#         self.SHAREPOINT_URL_SITE = os.getenv('SHAREPOINT_URL_SITE')
#         self.SHAREPOINT_SITE_NAME = os.getenv('SHAREPOINT_SITE_NAME')
#         self.SHAREPOINT_DOC_LIBRARY = os.getenv('SHAREPOINT_DOC_LIBRARY')
#         self.DOWNLOAD_PATH = os.getenv('DOWNLOAD_PATH')

#     def ConnectSharepoint(self):
#         return ClientContext(self.SHAREPOINT_URL_SITE).with_credentials(UserCredential(self.SHAREPOINT_EMAIL, self.SHAREPOINT_PASSWORD))

    
#     def DownloadTabelaAuxiliar(self, nomeUnidade):
#         ctx = self.ConnectSharepoint()

#         sharepointFilePath = f"{self.SHAREPOINT_DOC_LIBRARY}{datetime.now().year}/{nomeUnidade}/Tabela Auxiliar/Tabela Auxiliar.xlsx"
#         fileSavePath = self.DOWNLOAD_PATH + f"Tabela Auxiliar - {nomeUnidade}.xlsx"

#         print(fileSavePath)
#         try:

#             with open(fileSavePath, "wb") as local_file:
#                     ctx.web.get_file_by_server_relative_url(sharepointFilePath).download(local_file).execute_query()

#             return "Success"
#         except Exception as e:
#             print('>>>>>>>>>>>>>>>>>> ', e)
#             return f"Error: {str(e)}"


#     def UploadFile(self, filePath):
#         ctx = self.ConnectSharepoint()
#         folderPath = f"{self.SHAREPOINT_DOC_LIBRARY}{datetime.now().year}/Compilado Geral"
#         try:
#             folder = ctx.web.get_folder_by_server_relative_url(folderPath).get().execute_query()

#             with open(filePath, "rb") as content_file:
#                 file_content = content_file.read()
            
#             folder.upload_file(os.path.basename(filePath), file_content).execute_query()

#             return "Success"
#         except Exception as e:
#             return f"Error: {str(e)}"

#     def DeleteCompiladoGeral(self):
#         ctx = self.ConnectSharepoint()

#         try:
#             fileUrl = f"{self.SHAREPOINT_DOC_LIBRARY}{datetime.now().year}/Compilado Geral/Compilado.xlsx"
#             file = ctx.web.get_file_by_server_relative_url(fileUrl)
#             file.delete_object().execute_query()

#             return "Success"
#         except Exception as e:
#             return f"Error: {str(e)}"

import os
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

class Sharepoint:
    def __init__(self, logger):

        self.logger = logger
        # Configurações de autenticação moderna
        self.CLIENT_ID = os.getenv('ID_CLIENT')
        self.CLIENT_SECRET = os.getenv('SECRET_TD')
        self.TENANT_ID = os.getenv('TENANT_ID')
        
        # Configurações do SharePoint
        self.SHAREPOINT_URL_SITE = os.getenv('SHAREPOINT_URL_SITE')
        self.SHAREPOINT_SITE_NAME = os.getenv('SHAREPOINT_SITE_NAME')
        self.SHAREPOINT_DOC_LIBRARY = os.getenv('SHAREPOINT_DOC_LIBRARY')
        self.DOWNLOAD_PATH = os.getenv('DOWNLOAD_PATH')
        
        # Validação das variáveis de ambiente
        self._validate_config()

    def _validate_config(self):
        """Valida se todas as variáveis necessárias estão configuradas"""
        required_vars = {
            'ID_CLIENT': self.CLIENT_ID,
            'SECRET_TD': self.CLIENT_SECRET,
            'TENANT_ID': self.TENANT_ID,
            'SHAREPOINT_URL_SITE': self.SHAREPOINT_URL_SITE,
            'SHAREPOINT_SITE_NAME': self.SHAREPOINT_SITE_NAME,
            'SHAREPOINT_DOC_LIBRARY': self.SHAREPOINT_DOC_LIBRARY,
            'DOWNLOAD_PATH': self.DOWNLOAD_PATH
        }
        
        missing = [name for name, value in required_vars.items() if not value]
        if missing:
            raise ValueError(f"Variáveis de ambiente faltando: {', '.join(missing)}")

    def ConnectSharepoint(self):
        """Conecta ao SharePoint usando autenticação moderna"""
        try:
            credentials = ClientCredential(self.CLIENT_ID, self.CLIENT_SECRET)
            ctx = ClientContext(self.SHAREPOINT_URL_SITE).with_credentials(credentials)
            
            # Teste rápido para validar a conexão
            web = ctx.web
            ctx.load(web)
            ctx.execute_query()
            
            return ctx
        except Exception as e:
            print(f"Erro na conexão com o SharePoint: {str(e)}")
            raise

    def DownloadTabelaAuxiliar(self, nomeUnidade):
        """Download de arquivo com tratamento de erros melhorado"""
        try:
            ctx = self.ConnectSharepoint()
            
            # Construção do caminho no SharePoint
            sharepointFilePath = f"{self.SHAREPOINT_DOC_LIBRARY}{datetime.now().year}/{nomeUnidade}/Tabela Auxiliar/Tabela Auxiliar.xlsx"
            fileSavePath = os.path.join(self.DOWNLOAD_PATH, f"Tabela Auxiliar - {nomeUnidade}.xlsx")
            
            print(f"Baixando arquivo para: {fileSavePath}")
            
            # Download do arquivo
            with open(fileSavePath, "wb") as local_file:
                ctx.web.get_file_by_server_relative_url(sharepointFilePath).download(local_file).execute_query()
            
            if not os.path.exists(fileSavePath):
                raise Exception("Arquivo não foi baixado corretamente")
                
            return "Success"
            
        except Exception as e:
            error_msg = f"Erro ao baixar Tabela Auxiliar: {str(e)}"
            print(error_msg)
            return f"Error: {error_msg}"

    def UploadFile(self, filePath):
        """Upload de arquivo com verificação de sucesso"""
        try:
            self.logger.info(f"Iniciando upload do arquivo {filePath}")
            if not os.path.exists(filePath):
                raise Exception("Arquivo local não encontrado")
            
            ctx = self.ConnectSharepoint()
            folderPath = f"{self.SHAREPOINT_DOC_LIBRARY}{datetime.now().year}/Compilado Recobranca"
            
            #Verifica se a pasta existe
            self.logger.info(f"Verificando se a pasta {folderPath} existe.")
            if self.PastaExiste(folderPath):
                # Obtém a referência da pasta
                folder = ctx.web.get_folder_by_server_relative_url(folderPath)
            else:
                self.logger.info(f"Pasta não existe. Criando diretório no sharepoint")                
                try:
                    folder = self.CriarPasta(folderPath)            
                    self.logger.info("Pasta criada com sucesso")
                except Exception as e:
                    self.logger.error(f"Não foi possivel criar pasta. Erro {str(e)}")

                
            
            ctx.load(folder)
            ctx.execute_query()
            
            # Faz o upload
            with open(filePath, "rb") as content_file:
                file_content = content_file.read()
            
            folder.upload_file(os.path.basename(filePath), file_content).execute_query()
            return "Success"
            
        except Exception as e:
            error_msg = f"Erro no upload do arquivo: {str(e)}"
            print(error_msg)
            return f"Error: {error_msg}"

    def DeleteCompiladoGeral(self):
        """Exclusão de arquivo com confirmação"""
        try:
            ctx = self.ConnectSharepoint()
            fileUrl = f"{self.SHAREPOINT_DOC_LIBRARY}{datetime.now().year}/Compilado Recobranca/compilado_recobranca.xlsx"
            
            # Verifica se o arquivo existe antes de deletar
            file_to_delete = ctx.web.get_file_by_server_relative_url(fileUrl)
            ctx.load(file_to_delete)
            ctx.execute_query()
            
            if not file_to_delete.exists:
                return "Success"  # Já não existe
                
            # Deleta o arquivo
            file_to_delete.delete_object()
            ctx.execute_query()
            
            return "Success"
            
        except Exception as e:
            error_msg = f"Erro ao deletar arquivo: {str(e)}"
            print(error_msg)
            return f"Error: {error_msg}"
        
    def PastaExiste(self, caminho_relativo):
        ctx = self.ConnectSharepoint()
        try:
            ctx.web.get_folder_by_server_relative_url(caminho_relativo).get().execute_query()
            return True
        except Exception as e:
            if "FILE NOT FOUND" in str(e).upper() or "DOES NOT EXIST" in str(e).upper() or "NÃO ENCONTRADO" in str(e).upper():
                return False
            else:
                raise e  # Se for outro erro, lança para tratar fora
            
    def CriarPasta(self, caminho_relativo):
        ctx = self.ConnectSharepoint()
        try:
            folder = ctx.web.folders.add(caminho_relativo).execute_query()
            return folder
        except Exception as e:
            raise Exception(f"Houve um erro ao criar a pasta: {str(e)}")

    def ListarPastas(self, caminho_relativo):
        ctx = self.ConnectSharepoint()
        try:
            folder = ctx.web.get_folder_by_server_relative_url(caminho_relativo).get().execute_query()
            subpastas = folder.folders.get().execute_query()
            
            for subpasta in subpastas:
                print("📁", subpasta.properties["Name"])
        except Exception as e:
            print("Erro ao listar pastas:", e)