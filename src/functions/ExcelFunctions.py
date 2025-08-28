import pandas as pd

class ExcelFunctions:

    def GetInstrutores(self, df):
        # Lista de domínios permitidos
        dominios_permitidos = ['@sesi-es.org.br', '@senai-es.org.br', '@findes.org.br', '@docente.senai.br']

        # Remove as linhas
        df = df.dropna(subset=['EMAIL'])
        
        # Filtrando os e-mails que terminam com os domínios permitidos
        df_filtrado = df[df['EMAIL'].str.endswith(tuple(dominios_permitidos))]
        
        # Selecionando as colunas necessárias e removendo duplicatas
        self.dfInstrutores = df_filtrado[["INSTRUTOR", "EMAIL"]].drop_duplicates()
        
        return self.dfInstrutores

    # def GetInstrutores(self, df):
    #     self.dfInstrutores = df[["INSTRUTOR", "EMAIL"]].drop_duplicates()
    #     return self.dfInstrutores

    def CreateHTMLTable(self, dicData):

        return_str = '<table style="border-collapse: collapse; border: 1px solid #333333;"><tr>'

        for key in dicData[0].keys():
            if key == "FREQUENCIALIBERADA":
                return_str = return_str + '<th class="header" style="background-color: #333333; color: #FFFFFF;">' + "FREQUENCIA LIBERADA" + '</th>'
            elif key == "CONTEUDOREALIZADO":
                return_str = return_str + '<th class="header" style="background-color: #333333; color: #FFFFFF;">' + "CONTEUDO REALIZADO" + '</th>'
            elif key == "CONTEUDOPREVISTO":
                return_str = return_str + '<th class="header" style="background-color: #333333; color: #FFFFFF;">' + "CONTEUDO PREVISTO" + '</th>'
            else:
                return_str = return_str + '<th class="header" style="background-color: #333333; color: #FFFFFF;">' + key + '</th>'


        return_str = return_str + '</tr>'

        for key in dicData[1].keys():
            return_str = return_str + '<tr>'
            for subkey in dicData[1][key]:
                if subkey == "FREQUENCIALIBERADA" and dicData[1][key][subkey] == "NÃO":
                    return_str = return_str + '<td style="background-color: yellow; border: 1px solid #333333; padding: 8px;">' + str(dicData[1][key][subkey]) + '</td>'
                elif subkey == "CONTEUDOREALIZADO" and dicData[1][key][subkey] == "VAZIO":
                    return_str = return_str + '<td style="background-color: yellow; border: 1px solid #333333; padding: 8px;">' + str(dicData[1][key][subkey]) + '</td>'
                else:
                    return_str = return_str + '<td style="border: 1px solid #333333; padding: 8px;">' + str(dicData[1][key][subkey]) + '</td>'

        return_str = return_str + '</tr></table>'

        return return_str