import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

class ConectSheets():

    def __init__(self):
        self.creds = None
        self.service = None
        self.sheet = None
        self.ID = "1Q_Wn0HJkP9D-qEJ43ODrwHE89taw8pN4iryY9f-4fvM"
        self.dadosEncomenda = "Dadosdeencomenda!A3:O"
        self.caixa = "Caixa!A4:L"
        self.materiais = "Códigodemateriais!A3:G"

    def connect(self):
        SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
        if os.path.exists("token.json"):
            self.creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
                self.creds = flow.run_local_server(port=0)
            
            with open("token.json", "w") as token:
                token.write(self.creds.to_json())

        try:
            self.service = build("sheets", "v4", credentials=self.creds)
            self.sheet = self.service.spreadsheets()
            print("Conexão com Google Sheets estabelecida com sucesso!")
        except HttpError as err:
            print(f"Erro ao conectar ao Google Sheets: {err}")
            raise

    def registraEncomenda(self, dadosEncomenda):
        try:
            # Pega o index na planilha
            linha_Index = self.sheet.values().get(
                spreadsheetId=self.ID,
                range="Dadosdeencomenda!O2",  # Sem o valueInputOptions
            ).execute()

            # Converte a linha em int
            linha = int(linha_Index['values'][0][0])

            # Atualiza a linha a ser adicionada
            IDrange = "Dadosdeencomenda!A" + str(linha) + ":L" + str(linha)

            # Passa os valores para a planilha
            result = self.sheet.values().update(
                spreadsheetId=self.ID,
                range=IDrange,
                valueInputOption = "USER_ENTERED",
                body={'values': dadosEncomenda}  # Sem o valueInputOptions
            ).execute()

            linha+=1

            # Atualiza o index
            result = self.sheet.values().update(
                spreadsheetId=self.ID,
                range="Dadosdeencomenda!O2",
                valueInputOption = "USER_ENTERED",
                body={'values': [[linha]]}  # Sem o valueInputOptions
            ).execute()
                     

        except HttpError as err:
            return err
            raise


    def registraCaixa(self, caixa):
        try:
            # Pega o index na planilha
            linha_Index = self.sheet.values().get(
                spreadsheetId=self.ID,
                range="Caixa!L2",
            ).execute()

            # Converte a linha em int
            linha = int(linha_Index['values'][0][0])

            # Atualiza a linha a ser adicionada
            IDrange = "Caixa!A" + str(linha) + ":J" + str(linha)

            # Passa os valores para a planilha
            result = self.sheet.values().update(
                spreadsheetId=self.ID,
                range=IDrange,
                valueInputOption = "USER_ENTERED",
                body={'values': caixa}  # Sem o valueInputOptions
            ).execute()

            linha+=1

            # Atualiza o index
            result = self.sheet.values().update(
                spreadsheetId=self.ID,
                range="Caixa!L2",
                valueInputOption = "USER_ENTERED",
                body={'values': [[linha]]}  # Sem o valueInputOptions
            ).execute()
                     

        except HttpError as err:
            return err
            raise

    def registraMaterial(self, material):
        try:
            # Pega o index na planilha
            linha_Index = self.sheet.values().get(
                spreadsheetId=self.ID,
                range="Códigodemateriais!G2",
            ).execute()

            # Converte a linha em int
            linha = int(linha_Index['values'][0][0])

            # Atualiza a linha a ser adicionada
            IDrange = "Códigodemateriais!A" + str(linha) + ":J" + str(linha)

            # Passa os valores para a planilha
            result = self.sheet.values().update(
                spreadsheetId=self.ID,
                range=IDrange,
                valueInputOption = "USER_ENTERED",
                body={'values': material}  # Sem o valueInputOptions
            ).execute()

            linha+=1

            # Atualiza o index
            result = self.sheet.values().update(
                spreadsheetId=self.ID,
                range="Códigodemateriais!G2",
                valueInputOption = "USER_ENTERED",
                body={'values': [[linha]]}  # Sem o valueInputOptions
            ).execute()
                     

        except HttpError as err:
            return err
            raise

    def buscaMat(self, material):
        try:
            # Pega o index na planilha
            linha_Index = self.sheet.values().get(
                spreadsheetId=self.ID,
                range="Códigodemateriais!G2",
            ).execute()

            # Converte a linha em int
            linha = int(linha_Index['values'][0][0])

            # Atualiza a linha a ser adicionada
            IDrange = "Códigodemateriais!A" + str(linha) + ":J" + str(linha)

            # Pega o index na planilha
            linha_Index = self.sheet.values().get(
                spreadsheetId=self.ID,
                range=IDrange,
            ).execute()   

            #Buscando um valor de material
            for i in range(linha-1):
                if material[0] == linha_Index['values'][i][0] and material[1] == linha_Index['values'][i][1]:
                    valorMat = str(linha_Index['values'][i][6])
                else:
                    i+=1

        except HttpError as err:
            return err
            raise