from conectSheets import ConectSheets

class Data():
    
    def __init__(self) -> None:

        self.numero = None
        self.descricao = None
        self.material = None
        self.dataInicio = None
        self.dataFim = None
        self.nomeSolicitante = None
        self.formaDeContato = None
        self.responsavel = None
        self.valorFinal = None

        self.caixa = None
        self.codigosMateriais = None
         #Inicializando a comunicação
        self.save = ConectSheets()
        self.save.connect()

    def gravarEncomenda(self):
        #Armazena os dados da encomenda em uma lista
        dadosEncomenda = [
            self.numero,
            self.descricao,
            self.material,
            self.dataInicio,
            self.dataFim,
            self.nomeSolicitante,
            self.formaDeContato,
            self.responsavel,
        ]
        
        #chama a função gravador
        self.save.registraEncomenda(dadosEncomenda) 
        
    def calculo3D(self, peso, filamento, tempoImpressao, modelagem):
        
        valorFil = peso #* #save.buscaMaterial(filamento)

        # Define o valor do tempo de impressao
        if int(tempoImpressao) > 320:
            self.trabalhoMaquina = 24
        else:
            self.trabalhoMaquina = tempoImpressao * 0,23

        if modelagem is not None:
            vModel = modelagem*16
        
        #calcula o valor da impressao
        self.valorFinal = int(valorFil + tempoImpressao + modelagem)/0.66

        return self.valorFinal
    
    def calculoLaser(self, area, material, tempo_Trabalho, desing):

        valorMat = area #* save.buscaMaterial(material)
        
        custo_Tempo = (tempo_Trabalho * 0.01) / 0,66

         # Calculo do custo do desing
        valor_Desing = 0
        if desing is None:
            valor_Desing = desing * 16.00

        self.valorFinal = (valorMat + custo_Tempo + valor_Desing)/0,66

        return self.valorFinal
    
    def calculoPCB(self):
    
        pass

    #Busca a lista de itens com o mesmo nome na base
    def buscaInsumo(self, nome):
        material = None
        for i in self.save.Lista:
            if (nome == self.save.Lista[i][0]):
                cont = i
                while nome == self.save.Lista[0][i]:
                    material.append(self.save.Lista[cont])
                    cont+=1
            break
        return material
