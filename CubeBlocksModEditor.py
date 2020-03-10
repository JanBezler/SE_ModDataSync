import os
import numpy
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from xml.etree import ElementTree
import datetime
import tkinter as tk
import sys

class Main(object):
    def __init__(self,log):
        self.tab = []
        self.index = 0
        self.log = log

        self.MakeTree()
        self.MakeTab()
        self.MakeComponents()
        self.CheckQuantityOfComponents()
        self.MakeAdditionalColumns()
        self.MakeBlob()
        self.SendToGoogle()

    def MakeTree(self):
        # ---reading from file---
        file = open("DataSync.conf", "r")
        self.path = file.readline().rstrip('\n')
        file.close()
        #self.path = "M:\Gity\SEModsPack\Rebels Industry System\src\Data"
        self.tree = ElementTree.parse(os.path.abspath(os.path.join(self.path, "Blocks_CubeBlocks.sbc")))
        self.root = self.tree.getroot()
        # --finding all needed things---
        self.TypeId = self.tree.findall('CubeBlocks/Definition/Id/TypeId')
        self.SubtypeId = self.tree.findall('CubeBlocks/Definition/Id/SubtypeId')
        self.DisplayName = self.tree.findall('CubeBlocks/Definition/DisplayName')
        self.CubeSize = self.tree.findall('CubeBlocks/Definition/CubeSize')
        self.BlockPairName = self.tree.findall('CubeBlocks/Definition/BlockPairName')
        self.Components = self.tree.findall('CubeBlocks/Definition/Components')

    def MakeBlob(self):
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('PythonInputRG-fdcbaeee7393.json', scope)
        client = gspread.authorize(creds)
        self.sheet = client.open('Rebel ECONOMY - sync mod data tests').worksheet("PythonInput")
        self.logsheet = client.open('Rebel ECONOMY - sync mod data tests').worksheet("SynchroLogs")

        self.logrow=[self.log.get('nick'),self.log.get('datatime'),self.log.get('reason')]
        self.tab = numpy.array(self.tab).transpose()
        self.cell_list = self.sheet.range("A1:{0}{1}".format(self.column(len(self.tab[1])), len(self.tab) + 1))
        x = 0
        for i in range(len(self.tab)):
            for j in range(len(self.tab[0])):
                self.cell_list[x].value = self.tab[i][j]
                x += 1

    def MakeComponents(self):
        self.dictionaries = []
        for cubeBlocks in self.root:
            for cbDefinition in cubeBlocks:
                self.index += 1
                displayName = str()
                for definitionChild in cbDefinition:
                    if (definitionChild.tag == "DisplayName"):
                        displayName = definitionChild.text
                    if definitionChild.tag == "Components":
                        for component in definitionChild:
                            self.dictionaries.append(
                                {'displayName': displayName, 'index': self.index, "comp": component.attrib.get("Subtype"),
                                 "count": component.attrib.get("Count")})

    def CheckQuantityOfComponents(self):
        for item in self.dictionaries:
            for i in range(len(self.tab)):
                if self.tab[i][0] == item.get("comp"):
                    count = item.get("count")
                    if (self.tab[i][int(item.get("index"))]):
                        self.tab[i][int(item.get("index"))] = str(int(self.tab[i][int(item.get("index"))]) + int(count))
                    else:
                        self.tab[i][int(item.get("index"))] = count

    def MakeTab(self):
        for i in range(46):
            self.tab.append([])
            for j in range(len(self.SubtypeId) + 1):
                self.tab[i].append(j)
                self.tab[i][j] = ""

        self.tab[0][0] = "Blocks \ Components"
        self.tab[1][0] = "SteelPlate"
        self.tab[2][0] = "SteelPlateBox"
        self.tab[3][0] = "InteriorPlate"
        self.tab[4][0] = "InteriorPlateBox"
        self.tab[5][0] = "Construction"
        self.tab[6][0] = "Computer"
        self.tab[7][0] = "Motor"
        self.tab[8][0] = "SmallTube"
        self.tab[9][0] = "LargeTube"
        self.tab[10][0] = "Display"
        self.tab[11][0] = "BulletproofGlass"
        self.tab[12][0] = "MetalGrid"
        self.tab[13][0] = "Superconductor"
        self.tab[14][0] = "ReactorComponents"
        self.tab[15][0] = "Thrust"
        self.tab[16][0] = "Girder"
        self.tab[17][0] = "SolarCell"
        self.tab[18][0] = "Detector"
        self.tab[19][0] = "RadioCommunication"
        self.tab[20][0] = "MedicalComponents"
        self.tab[21][0] = "GravityGenerator"
        self.tab[22][0] = "PowerCell"
        self.tab[23][0] = "Explosives"
        self.tab[24][0] = "Robotics"
        self.tab[25][0] = "Coolant"
        self.tab[26][0] = "CoolantBox"
        self.tab[27][0] = "PowerCore"
        self.tab[28][0] = "PowerCoreBox"
        self.tab[29][0] = "VolumeExtender"
        self.tab[30][0] = "Stabilizer"
        self.tab[31][0] = "StabilizerBox"
        self.tab[32][0] = "Screw"
        self.tab[33][0] = "ScrewBox"
        self.tab[34][0] = "PalladiumComponent"
        self.tab[35][0] = "PalladiumComponentBox"
        self.tab[36][0] = "WeaponComponent"
        self.tab[37][0] = "ToolComponent"
        self.tab[38][0] = "PirateComponent"
        self.tab[39][0] = "AdvancedComponent"
        self.tab[40][0] = "AdvancedTransmitter"
        self.tab[41][0] = "AdvancedProcessor"
        self.tab[42][0] = "TypeID"
        self.tab[43][0] = "SubtypeID"
        self.tab[44][0] = "BlockPairName"
        self.tab[45][0] = "CubeSize"

    def MakeAdditionalColumns(self):
        for i in range(len(self.DisplayName)):
            self.tab[0][i + 1] = str(i + 1) + "." + self.DisplayName[i].text
            self.tab[42][i + 1] = str(self.TypeId[i].text)
            self.tab[43][i + 1] = str(self.SubtypeId[i].text)
            self.tab[44][i + 1] = str(self.BlockPairName[i].text)
            # self.tab[45][i+1]=str(self.CubeSize[i].text)

    def column(self,n):
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string

    def SendToGoogle(self):
        self.sheet.update_cells(self.cell_list)
        self.logsheet.append_row(self.logrow)
        window = tk.Tk()
        window.title("ModDataSynchroCB")
        tk.Label(window, text="Success!", padx=20).pack()
        butt = tk.Button(window, text="OK", width=10, command=sys.exit)
        butt.pack()
        tk.mainloop()

class Menu(object):

    def __init__(self):
        self.window = tk.Tk()
        self.window.title("ModDataSynchroCB")
        self.name = tk.StringVar()
        self.road = tk.StringVar()
        tk.Label(self.window, text="Tool for synchronising data", padx=20).pack()

        try:
            file = open("DataSync.conf", "r")
            self.filelines=file.readlines()
            self.nick=self.filelines[1].rstrip('\n')
            self.path=self.filelines[0].rstrip('\n')
            self.name.set("Your nick: " + self.nick)
            self.road.set("Your path: " + self.path)
            tk.Label(self.window, textvariable=self.name, padx=100).pack()
            tk.Label(self.window, textvariable=self.road, padx=100).pack()
            file.close()

        except FileNotFoundError:
            tk.Label(self.window, text="Write your nick:", padx=20).pack()
            self.nick = tk.Entry(self.window, width=16)
            self.nick.pack()

            tk.Label(self.window, text="Write your path:", padx=20).pack()
            self.path = tk.Entry(self.window, width=60)
            self.path.pack()
            tk.Label(self.window, text="Example: M:\Gity\SEModsPack\Rebels Industry System\src\Data", padx=20).pack()

            button = tk.Button(self.window, text="Save", width=10, command=self.SaveData)
            button.pack()

        tk.Label(self.window, text="Reason for synchro:").pack()
        self.reason = tk.Text(self.window, width=30, height=6)
        self.reason.pack()

        submit = tk.Button(self.window, text="Synchronize", width=10, command=self.SendIt)
        submit.pack()

        tk.mainloop()

    def SendIt(self):
        log ={
            'nick': str(self.nick),
            'reason': str(self.reason.get("1.0",'end-1c')),
            'datatime': str(datetime.datetime.now().strftime("%I:%M %p %d.%B.%Y"))
        }

        Main(log)
        sys.exit(0)

    def SaveData(self):
        file = open("DataSync.conf", "w")
        file.write(self.path.get()+"\n")
        file.write(self.nick.get())
        self.nick = self.nick.get()
        file.close()

if __name__=="__main__":
    Menu()