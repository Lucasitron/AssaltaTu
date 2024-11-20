from kivymd.uix.screen import MDScreen
from kivymd.uix.screenmanager import MDScreenManager
from kivy.core.window import Window
from kivymd.app import MDApp
from kivy.lang import Builder
from kivymd.uix.menu import MDDropdownMenu

Window.size = (360, 640)

class TelaPrincipal(MDScreen):
    pass

class TelaImpressao3D(MDScreen):
    pass

class MainApp(MDApp):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.menu = None

    def menu_open(self):
        # Obter a tela "TelaImpressao3D" do ScreenManager
        tela_impressao = self.root.get_screen('TelaImpressao3D')
        
        # Configura o caller usando o botão dentro da tela específica
        menu_items = [
            {
                "text": f"Item {i}",
                "on_release": lambda x=f"Item {i}": self.menu_callback(x),
            } for i in range(5)
        ]
        self.menu = MDDropdownMenu(caller=tela_impressao.ids.button, items=menu_items)
        self.menu.open()

    def menu_callback(self, text_item):
        print(text_item)
        if self.menu:
            self.menu.dismiss()

    def build(self):
        self.theme_cls.theme_style = "Dark"
        Builder.load_file('tela.kv')
        
        managerScreen = MDScreenManager()
        managerScreen.add_widget(TelaPrincipal(name="TelaPrincipal"))
        managerScreen.add_widget(TelaImpressao3D(name="TelaImpressao3D"))

        return managerScreen

MainApp().run()
