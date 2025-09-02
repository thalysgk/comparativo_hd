# main.py
# Ponto de entrada principal da aplicação.
# Sua única responsabilidade é instanciar e executar a interface gráfica.

from gui import App

if __name__ == "__main__":
    # Cria a instância da aplicação principal
    app = App()
    # Inicia o loop de eventos do Tkinter
    app.mainloop()