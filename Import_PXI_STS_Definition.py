from xml.dom.minidom import parse
import xml.dom.minidom
import xlwings as xw


EPath = r'C:\Users\jluo\Desktop\Python\Preliminary_Chanmap_Telink.xlsm'
wb = xw.Book(EPath)

# Parse the Deinition page, Collect related info
Model = {} #{Model Name,[ModelInfo]}
InstrumentTypes = set([])
SelectedArea = 'A2:E'+ str(wb.sheets['Definition'].range('K1').options(numbers=int).value+1)
InstrumentRng = wb.sheets['Definition'].range(SelectedArea)
for i in range(len(InstrumentRng.rows)):
    ModelInfos = InstrumentRng.rows[i]
    ModelInfo = [] #[Name,Slots,Color]
    ModelInfo.append(str(ModelInfos[1].options(numbers=int).value))
    ModelInfo.append(ModelInfos[2].options(numbers=int).value)
    PogoStr = ModelInfos[3].value
    if PogoStr != None:
        ModelInfo.append(PogoStr.split(','))
    else:
        ModelInfo.append([''])
    ModelInfo.append(ModelInfos[4].color)
    Model[ModelInfos[0].value] = ModelInfo


SelectedArea = 'H2:H'+str(wb.sheets['Definition'].range('K2').options(numbers=int).value+1)
SpringPinTypeRng = wb.sheets['Definition'].range(SelectedArea)
for SpringPinType in SpringPinTypeRng:
    InstrumentTypes.add(SpringPinType.value)

# set the path
FilePath = 'C:\\Users\\jluo\\Desktop\\Python'
PXI_Path = FilePath + '\\PXI_Definition.xml'
STS_Path = FilePath + '\\STS_Definition.xml'

# Parse the PXI Definition and get the all the instrument info into PXI_ID
PXI_ID = {}  #{Instrument ID,[InP]}
Ins_Type = {} #{Modle,Index}
# Function to get Instrument Info, support VST,PXI,USB
def GetInstrumentInfo(PXIChassis, InstrumentName):
    try:
        i = int(PXIChassis.getAttribute("Number"))
    except:
        i = 0
    Instruments = PXIChassis.getElementsByTagName(InstrumentName)
    for Instrument in Instruments:
        InP = []  #[Chassis Number, Model, Index, Slot Number]
        InP.append(i)
        if Instrument.hasAttribute("Name") & (Instrument.getAttribute("Name")!=''):
            ModelName = Instrument.getAttribute("Model")
            InP.append(ModelName)
            if ModelName in Ins_Type:
                Ins_Type[ModelName] += 1
            else:
                Ins_Type[ModelName] = 1
            InP.append(Ins_Type[ModelName])
            try:
                InP.append(int(Instrument.getAttribute("Slot")))
            except:
                InP.append(0)
            InP.append(Model[ModelName])
            PXI_ID[Instrument.getAttribute("Name")] = InP

DOMTreePXI = xml.dom.minidom.parse(PXI_Path)
PXI_SD = DOMTreePXI.documentElement
PXIChassises = PXI_SD.getElementsByTagName("PXIChassis")
# Collection all the info
GetInstrumentInfo(PXI_SD,"USB")
for PXIChassis in PXIChassises:
    GetInstrumentInfo(PXIChassis,"PXI")
    GetInstrumentInfo(PXIChassis,"VST")

# Parse the PXI Defination and get the pogo info to the PXI_ID
# Function to get all pogo info
Ins_Type_set = set([])
def GetPogoInfo(STS_SD, InstrumentType):
    Pogo = []
    try:
        Instruments = STS_SD.getElementsByTagName(InstrumentType)
        if InstrumentType =='SpringPinSystemPSU':
            Pogo = ['P0']
            PXI_ID[Instruments[0].getAttribute("Instrument")].append(Pogo)
            return
        elif InstrumentType=='SpringPinSystemDMM':
            Pogo = ['P1']
            PXI_ID[Instruments[0].getAttribute("Instrument")].append(Pogo)
            return
        elif InstrumentType=='SpringPinSystemSMU':
            Pogo = ['P2']
            PXI_ID[Instruments[0].getAttribute("Instrument")].append(Pogo)
            return
        for x in range(len(Instruments)):
            Instrument = Instruments[x].getAttribute("Instrument")
            if Instrument not in Ins_Type_set:
                Ins_Type_set.add(Instrument)
                Pogo = []
            Pogo.append(Instruments[x].getAttribute("Position"))
            if x == len(Instruments)-1:
                PXI_ID[Instrument].append(Pogo)
            elif Instruments[x+1].getAttribute("Instrument") not in Ins_Type_set:
                PXI_ID[Instrument].append(Pogo)
    except:
        None

DOMTreeSTS = xml.dom.minidom.parse(STS_Path)
STS_SD = DOMTreeSTS.documentElement

for InstrumentType in InstrumentTypes:
    GetPogoInfo(STS_SD,InstrumentType)


#########################################################################################################
#wb.sheets['Spring Probe Map'].range('I14').value = 1
#wb.sheets['Spring Probe Map'].range('J15:J14').api.MergeCells = True
#wb.sheets['Spring Probe Map'].range('I14:J15').color = (217,217,217)

# Fucntion to get the slot selection area
def GetSlotArea(ChassisNumber, slot, slotsize):
    Area = ''
    CS_1_2_3_4 = [['AE','AF',12,1],['I','J',31,-1],[33,34],[9,10]]
    CS = ['L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC']
    if ChassisNumber <3:
        CS_1_2 = CS_1_2_3_4[ChassisNumber-1]
        X = CS_1_2[2]+slot*CS_1_2[3]
        S = X+(slotsize-1)*CS_1_2[3]
        MAX = str(max(X,S))
        MIN = str(min(X,S))
        Area = CS_1_2[0] + MIN + ':' + CS_1_2[1] + MAX
    else:
        CS_3_4 = CS_1_2_3_4[ChassisNumber-1]
        MAX = str(max(CS_3_4[1],CS_3_4[0]))
        MIN = str(min(CS_3_4[1],CS_3_4[0]))
        if ChassisNumber==3:
            CS.reverse()
            Area = CS[slot + slotsize - 2] + MIN + ":" + CS[slot - 1] + MAX
        else:
            Area = CS[slot - 1] + MIN + ":" + CS[slot + slotsize - 2] + MAX
    return Area

# Function to get Pogo block selection area
def GetPogoArea(Pogo):
    PNum = [[9,15,21,27,33],[3,5,37,39]]
    PKeyA = [['C','D'],['E','F'],['AI','AJ'],['AK','AL']]
    PKeyB = [['H','I'],['N','O'],['T','U'],['Z','AA'],['AF','AG']]
    PogoNum = int(Pogo.split('P')[1])
    if 100< PogoNum <120:
        (PogoD,PogoM) = divmod(PogoNum - 102,4)
        PNum[0].reverse()
        Area = PKeyA[2+PogoM][0] + str(PNum[0][PogoD]) + ":" + PKeyA[2+PogoM][1] + str(PNum[0][PogoD]+1)
    elif 120< PogoNum <140:
        (PogoD,PogoM) = divmod(PogoNum - 122,4)
        Area = PKeyA[1-PogoM][0] + str(PNum[0][PogoD]) + ":" + PKeyA[1-PogoM][1] + str(PNum[0][PogoD]+1)
    elif 140< PogoNum <160:
        (PogoD,PogoM) = divmod(PogoNum - 142,4)
        Area = PKeyB[PogoD][0] + str(PNum[1][2+PogoM]) + ":" + PKeyB[PogoD][1] + str(PNum[1][2+PogoM]+1)
    elif 160< PogoNum <180:
        (PogoD,PogoM) = divmod(PogoNum - 162,4)
        PKeyB.reverse()
        Area = PKeyB[PogoD][0] + str(PNum[1][1-PogoM]) + ":" + PKeyB[PogoD][1] + str(PNum[1][1-PogoM]+1)
    elif PogoNum ==0:
        Area = "H40:H40"
    elif PogoNum ==1:
        Area = "I40:I40"
    elif PogoNum ==2:
        Area = "H37:I38"
    return Area


# Fucntion to fill the value and color to the excel
def SetCell(Sheets, Area, Value, Colors):
    StartArea = Area.split(":")[0]
    if Sheets.range(StartArea).value != None:
        Value = Value + " &" +str(Sheets.range(StartArea).value)
    Sheets.range(StartArea).value = Value
    Sheets.range(Area).api.MergeCells = True
    Sheets.range(Area).color = Colors
    Sheets.range(Area).WrapText = True

#wb.sheets['Spring Probe Map'].api.Copy(Before=wb.sheets['Spring Probe Map'].api)
def Main():
    for ID in PXI_ID:
        print(ID,end='')                   #####################################################
        Info = PXI_ID[ID]
        print(Info)
        ModelInfo = Info[4]
        PXI_Value = ModelInfo[0]
        if Ins_Type[Info[1]] != 1:
            PXI_Value +='_'+str(Info[2])
        if 0< Info[0] <5: # Differe from USB
            SetCell(wb.sheets['Spring Probe Map'],GetSlotArea(Info[0],Info[3],ModelInfo[1]),PXI_Value,ModelInfo[3])
        try: # Some instrument don't have a pogo block
            if (len(Info[5]))>1:
                for i in range(len(Info[5])):
                    Pogo_Value = PXI_Value + '_'+ ModelInfo[2][i]
                    SetCell(wb.sheets['Spring Probe Map'], GetPogoArea(Info[5][i]), Pogo_Value, ModelInfo[3])
            else:
                Pogo_Value = PXI_Value
                SetCell(wb.sheets['Spring Probe Map'], GetPogoArea(Info[5][0]), Pogo_Value, ModelInfo[3])
        except:
            None



































