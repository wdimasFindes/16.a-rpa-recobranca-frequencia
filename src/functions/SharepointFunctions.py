import os
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()
class Sharepoint:
    def __init__(self):
        self.SHAREPOINT_EMAIL = os.getenv('SHAREPOINT_EMAIL')
        self.SHAREPOINT_PASSWORD = os.getenv('SHAREPOINT_PASSWORD')
        self.SHAREPOINT_URL_SITE = os.getenv('SHAREPOINT_URL_SITE')
        self.SHAREPOINT_SITE_NAME = os.getenv('SHAREPOINT_SITE_NAME')
        self.SHAREPOINT_DOC_LIBRARY = os.getenv('SHAREPOINT_DOC_LIBRARY')
        self.DOWNLOAD_PATH = os.getenv('DOWNLOAD_PATH')

    def ConnectSharepoint(self):
        return ClientContext(self.SHAREPOINT_URL_SITE).with_credentials(UserCredential(self.SHAREPOINT_EMAIL, self.SHAREPOINT_PASSWORD))

    
    def DownloadTabelaAuxiliar(self, nomeUnidade):
        ctx = self.ConnectSharepoint()

        sharepointFilePath = f"{self.SHAREPOINT_DOC_LIBRARY}{datetime.now().year}/{nomeUnidade}/Tabela Auxiliar/Tabela Auxiliar.xlsx"
        fileSavePath = self.DOWNLOAD_PATH + f"Tabela Auxiliar - {nomeUnidade}.xlsx"
        try:

            with open(fileSavePath, "wb") as local_file:
                    ctx.web.get_file_by_server_relative_url(sharepointFilePath).download(local_file).execute_query()

            return "Success"
        except Exception as e:
            return f"Error: {str(e)}"

    def UploadFile(self, filePath):
        ctx = self.ConnectSharepoint()
        folderPath = f"{self.SHAREPOINT_DOC_LIBRARY}{datetime.now().year}/Compilado Recobranca"
        try:
            folder = ctx.web.get_folder_by_server_relative_url(folderPath).get().execute_query()

            with open(filePath, "rb") as content_file:
                file_content = content_file.read()
            
            folder.upload_file(os.path.basename(filePath), file_content).execute_query()

            return "Success"
        except Exception as e:
            return f"Error: {str(e)}"

    def DeleteCompiladoGeral(self):
        ctx = self.ConnectSharepoint()

        try:
            fileUrl = f"{self.SHAREPOINT_DOC_LIBRARY}{datetime.now().year}/Compilado Recobranca/compilado_recobranca.xlsx"
            file = ctx.web.get_file_by_server_relative_url(fileUrl)
            file.delete_object().execute_query()

            return "Success"
        except Exception as e:
            return f"Error: {str(e)}"