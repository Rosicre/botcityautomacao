from botcity.core import DesktopBot
# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
# from botcity.maestro import *

import pandas as pd

class Bot(DesktopBot):
    def action(self, execution=None):

        # Importar a base de dados
        tabela = pd.read_excel('Contatos.xlsx')
        print(tabela)
        # Para cada linha da base de dados


        # Opens the Whatssap website.
        self.browse("https://web.whatsapp.com/")

        for linha in tabela.index:

            contato = tabela.loc[linha, 'Contato']
            arquivo = tabela.loc[linha,'Arquivo']
            msg = tabela.loc[linha, 'Msg']

            if not self.find( "lupa", matching=0.97, waiting_time=60000):
                self.not_found("lupa")
            self.click()

            self.type_keys_with_interval(100,contato)
            self.enter()

            if not self.find( "anexar", matching=0.97, waiting_time=10000):
                self.not_found("anexar")
            self.click()

            if not self.find( "imagem", matching=0.97, waiting_time=10000):
                self.not_found("imagem")
            self.click()

            if not self.find( "nome", matching=0.97, waiting_time=10000):
                self.not_found("nome")
            self.paste(arquivo)
            self.enter()

            if not self.find( "enviar", matching=0.97, waiting_time=10000):
                self.not_found("enviar")
            self.click()
            self.wait(2000)

            if pd.isna(arquivo): # tem arquivo
                self.paste(msg)
                self.enter()
            else:
                self.paste(msg)
                self.wait(3000)
                self.enter()

            if self.find( "seta", matching=0.97, waiting_time=10000):
                self.click()


    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()

