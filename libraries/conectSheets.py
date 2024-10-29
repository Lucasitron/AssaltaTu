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
        self.Lista = None
        self.ID = "1Q_Wn0HJkP9D-qEJ43ODrwHE89taw8pN4iryY9f-4fvM"
        self.dadosEncomenda = "Dadosdeencomenda!A3:O"
        self.caixa = "Caixa!A4:L"
        self.materiais = "Códigodemateriais!A3:G"

    def connect(self):
        SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
        if os.path.exists("ConectSheets/token.json"):
            self.creds = Credentials.from_authorized_user_file("ConectSheets/token.json", SCOPES)
        
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file("ConectSheets/credentials.json", SCOPES)
                self.creds = flow.run_local_server(port=0)
            
            with open("ConectSheets/token.json", "w") as token:
                token.write(self.creds.to_json())

        try:
            self.service = build("sheets", "v4", credentials=self.creds)
            self.sheet = self.service.spreadsheets()
            print("Conexão com Google Sheets estabelecida com sucesso!")
        except HttpError as err:
            print(f"Erro ao conectar ao Google Sheets: {err}")
            raise

        try:
            # Pega o index na planilha
            linhaVal = self.sheet.values().get(
                spreadsheetId=self.ID,
                range="Códigodemateriais!H2",
            ).execute()

            # Converte a linha em int
            linha = int(linhaVal['values'][0][0])

            # Atualiza a linha base
            IDrange = "Códigodemateriais!A4" + ":F" + str(linha)

            # Pega o index na planilha
            self.Lista = self.sheet.values().get(
                spreadsheetId=self.ID,
                range=IDrange,
            ).execute()

            # Transforma self.Lista para conter apenas os dados
            self.Lista = self.Lista.get('values', [])
        except HttpError as err:
            return err
            raise


    def registraEncomenda(self, dadosEncomenda):
        try:
            # Pega o index na planilha
            linhaVal = self.sheet.values().get(
                spreadsheetId=self.ID,
                range="Dadosdeencomenda!O2",  # Sem o valueInputOptions
            ).execute()

            # Converte a linha em int
            linha = int(linhaVal['values'][0][0])

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
            linhaVal = self.sheet.values().get(
                spreadsheetId=self.ID,
                range="Caixa!L2",
            ).execute()

            # Converte a linha em int
            linha = int(linhaVal['values'][0][0])

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
            linhaVal = self.sheet.values().get(
                spreadsheetId=self.ID,
                range="Códigodemateriais!G2",
            ).execute()

            # Converte a linha em int
            linha = int(linhaVal['values'][0][0])

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

        