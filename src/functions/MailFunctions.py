from O365 import Account, FileSystemTokenBackend
from datetime import datetime, timedelta
import os
import requests
from dotenv import load_dotenv
import math


load_dotenv()


class MailFunctions:
    def GenerateToken(self):
        try:
            CREDENTIALS = (os.getenv('ID_CLIENT'), os.getenv('SECRET_TD'))

            token_backend = FileSystemTokenBackend(token_path=os.getenv('LOCAL_TOKEN_PATH'), token_filename=os.getenv('TOKEN_FILENAME'))
            account = Account(CREDENTIALS, token_backend=token_backend)

            if not account.is_authenticated:
                account.authenticate(scopes=['basic', 'calendar_all', 'onedrive_all', 'message_all'])

            expires_at = account.connection.token_backend.get_token()['expires_at']

            expires_at = datetime.fromtimestamp(expires_at)

            if expires_at - timedelta(minutes=5) <= datetime.now():
                account.connection.refresh_token()
                
            token_acess = account.connection.token_backend.get_token()['access_token']
            
            return token_acess
        except Exception as e:
            print(f'Erro na geração de Token de acesso {e}')


    def SendMail(self, token_acess, user_email, cc1, cc2, message, nomeInstrutor):
        headers = {
            'Authorization': 'Bearer ' + token_acess
        }

        # Inicializa a lista de destinatários em cópia
        cc_recipients = []

        # Função para verificar se o valor é um "nan" válido
        def is_valid_email(email):
            # Verifica se o valor é None, string vazia ou "nan" como string
            if email is None or email == '' or (isinstance(email, str) and email.lower() == "nan"):
                return False
            # Verifica se é NaN numérico
            if isinstance(email, float) and math.isnan(email):
                return False
            return True

        # Verifica se cc1 é válido
        if is_valid_email(cc1):
            cc_recipients.append({
                'emailAddress': {
                    'address': cc1
                }
            })

        # Verifica se cc2 é válido
        if is_valid_email(cc2):
            cc_recipients.append({
                'emailAddress': {
                    'address': cc2
                }
            })

        # Construa o corpo do e-mail
        request_body = {
            'message': {
                'toRecipients': [
                    {
                        'emailAddress': {
                            'address': user_email
                        }
                    }
                ],
                'ccRecipients': cc_recipients,  # Aqui estamos usando a lista filtrada de CCs
                'from': {
                    "emailAddress": {
                        "address": "pdgbot@findes.org.br"
                    }
                },
                'subject': f'Recobrança - Verificação de Pendências',
                'importance': 'normal',
                'body': {
                    'contentType': 'HTML',
                    'content': message
                },
            }
        }

        GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0'
        endpoint = GRAPH_ENDPOINT + '/me/sendMail'
        try:
            # Envia a requisição POST para o Graph API
            response = requests.post(endpoint, headers=headers, json=request_body)

            # Verifica a resposta da requisição
            if response.status_code == 202:
                return "success"
            else:
                # Levanta um erro caso a resposta não seja 202
                raise Exception(f"Erro ao enviar e-mail. Código de status: {response.status_code}, Razão: {response.reason}")

        except Exception as e:
            # Registra o erro no log
            self.logger.info(f"Erro no instrutor: {nomeInstrutor}. Motivo: {e}")
            return f"Error: {e}"

    
    def  CreateAndSendMail(self, emailInstrutor, emailSupervisor, emailOrientador, body):
        token = self.GenerateToken()
        try:
            result = self.SendMail(token, emailInstrutor, emailSupervisor, emailOrientador, body)
            if "Error" in result:
                return f"Erro ao enviar e-mail: {result}"
            return "success"
        except Exception as e:
            return f"Error: {e}"