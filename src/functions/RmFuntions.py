
from pywinauto import Application
from datetime import datetime, timedelta
from time import sleep
from pyautogui import press, hotkey, write, moveTo, doubleClick
import threading



class RM:
    def __init__(self, logger):
        self.rmPath = "C:\TOTVS\RM.NET\RM.exe"
        self.logger = logger
        

    def Main(self, reportPath):
        try:
            self.StartRm()
            sleep(20)
            self.logger.info("Fazendo login no RM.")
            self.Login("m2bot", "m2botRPA", "CorporeRM")
            sleep(240)
            self.logger.info("Preparando ambiente.")
            self.CloseAllWindows()
            self.GoToVisoesDados()
            self.logger.info("Selecionando consulta SQL.")
            self.SelectConsultaSql()
            self.SelectQuery()
            self.logger.info("Executando query e exportando arquivo.")
            diaAnterior = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
            self.FillAndExportQuery(reportPath, diaAnterior, diaAnterior)
            return "Success"
        except Exception as e:
            return f"Error in RM: {e}"


    def StartRm(self):
        self.app = Application(backend="uia").start(self.rmPath)
        self.loginWindow = self.app.TOTVS

    def Login(self, username, password, alias):

        userEmail= self.loginWindow.child_window(title="cUser", auto_id = "cUser", control_type="Edit")
        userEmail.wait('visible', timeout=240)
        userEmail.set_text(username)

        userPassword = self.loginWindow.child_window(title="cUser", auto_id = "cPassword" ,control_type="Edit")
        userPassword.set_text(password)

        aliasDropdown = self.loginWindow.child_window(title="cAlias", control_type="ComboBox")
        aliasDropdown.select(alias)
        sleep(2)
        press("enter")
    
    def CloseAllWindows(self):
        app = Application(backend="uia").connect(path=self.rmPath)
        mainWindow = app.window(found_index=0)
        mainWindow.set_focus()

        hotkey("alt", "a")
        sleep(1)
        hotkey("a", "m")
        sleep(1)
        press("j")
        press("f")


    def GoToVisoesDados(self):

        app = Application(backend="uia").connect(path=self.rmPath)
        mainWindow = app.window(found_index=0)
        mainWindow.set_focus()

        
        sleep(3)
        hotkey("alt", "g")
        sleep(3)
        hotkey("v", "i")
        sleep(5)

    def SelectConsultaSql(self):

        sleep(15)
        app = Application(backend="uia").connect(title="Filtros - Consultas SQL")
        mainWindow = app.window(title="Filtros - Consultas SQL")
        mainWindow.set_focus()

        press("s", presses=3, interval=1)
        sleep(1)
        press("enter")

        sleep(15)

    def SelectQuery(self):
        app = Application(backend="uia").connect(path=self.rmPath)
        mainWindow = app.window(found_index=0)

        customToolbar = mainWindow.child_window(title="customToolBar", control_type="ToolBar", found_index=0)

        buttonSearchCode = customToolbar.child_window(title="01 - Por Código", control_type="Button", found_index=0)
        buttonSearchCode.click()

        sleep(5)
        write("DEP.032")
        
        sleep(1)
        press("enter")
        
        sleep(15)

        mainWindow = app.window(found_index=0)
        
        treeview = mainWindow.child_window(title="treeList", auto_id="treeList", control_type="Tree", found_index=0)
        treeview.wait("exists", timeout=60)
        treeitem_n0 = treeview.child_window(control_type="TreeItem", found_index=0)
        text_row_0 = treeitem_n0.child_window(title="Texto row 0", control_type="DataItem", found_index=0)
        rect = text_row_0.rectangle()
        left, top, right, bottom = rect.left, rect.top, rect.right, rect.bottom

        mid_x = (left + right) // 2
        mid_y = (top + bottom) // 2

        mainWindow.set_focus()
        moveTo(mid_x, mid_y, duration=1)
        doubleClick()
        sleep(15)


    def FillAndExportQuery(self, filePath, startDate, endDate):
        app = Application(backend="uia").connect(path=self.rmPath)
        mainWindow = app.window(found_index=0)

        executeSqlWindow = mainWindow.child_window(title="Execução de Sentença SQL", found_index=0)
        executeSqlWindow.wait("ready", timeout=60)

        mainWindow.set_focus()
        sleep(5)
        press("right")
        write(startDate, interval=0.1)
        press("enter")
        write(endDate, interval=0.1)
        press("enter")

        self.logger.info("Procurando botao execute.")
        buttonExecute = executeSqlWindow.child_window(title="buttonExecute", control_type="Button", found_index=0)
        self.logger.info("Esperando botao execute estar disponivel para clicar.")
        buttonExecute.wait("enabled", timeout=60)
        self.logger.info("Clicando no botao execute")
        buttonExecute.click()

        sleep(15)

        self.logger.info("Procurando botão export")
        buttonExport = executeSqlWindow.child_window(title="buttonExport", control_type="MenuItem", found_index=0)
        sleep(10)
        self.logger.info("Esperando botão export ficar disponivel para clicar")
        buttonExport.wait("enabled", timeout=60)
        self.logger.info("Clicando no botão export")
        sleep(10)
        buttonExport.invoke()

        sleep(5)

        press("down", presses=4, interval=0.1)
        press("enter")
        
        sleep(10)

        press("tab")
        sleep(1)
        press("enter")
        press("r")
        sleep(2)
        write(filePath)
        press("enter")
        sleep(10)
        press("n")


