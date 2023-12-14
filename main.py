import openpyxl
import csv
import xml.etree.ElementTree as ET


class VisionBox:
    def __init__(self, name):
        # Inițializează liste pentru interfețele de rețea
        self.name = name
        self.eth0 = []
        self.eth1 = []
        self.eth2 = []
        self.eth3 = []
        self.eth4 = []
        self.stations = {}
        self.number_of_stations = 0
# adaugam statiitle nume, ordinea pentru station name, ordinea gloabala,pozitia in masina,nr spin

    def add_station(self, station_name, station_order, stglobalorder, tipst, spinnr=''):
        self.stations[station_name] = [station_order, stglobalorder, tipst, spinnr]
        self.number_of_stations += 1

    def add_ip(self, ip):  # assignam adresele ip
        for name, value in ip.items():
            if name == 'ETH0[IP]':
                for ipadd in value:
                    self.eth0.append(ipadd)
            if name == 'ETH1[IP]':
                for ipadd in value:
                    self.eth1.append(ipadd)
            if name == 'ETH2[IP]':
                for ipadd in value:
                    self.eth2.append(ipadd)
            if name == 'ETH3[IP]':
                for ipadd in value:
                    self.eth3.append(ipadd)
            if name == 'ETH4[IP]':
                for ipadd in value:
                    self.eth4.append(ipadd)

    def showip(self):  # afisam ip-urile
        print('ETHO: {}'.format(self.eth0))
        print('ETH1: {}'.format(self.eth1))
        print('ETH2: {}'.format(self.eth2))
        print('ETH3: {}'.format(self.eth3))
        print('ETH4: {}'.format(self.eth4))

    def writetags(self, datatags):  # cream fisierul install param pentru fiecare statie
        for numest, st in self.stations.items():
            base_tree = ET.ElementTree(ET.Element("Root"))
            for items in datatags:
                print(items['VariableName'])
                tag_tree = create_tag_list(items['VariableName'], items['PlcName'].format(st[1]), items['type'])
                base_tree.getroot().append(tag_tree)
            ET.indent(base_tree, space='  ', level=0)
            base_tree.write(numest+"output.xml", encoding="utf-8", xml_declaration=False, method="xml",
                            short_empty_elements=False)


def readips(excelfile, startrangeip, endrangeip, numarul_coloanei):  # citim ip-uriel din excel file
    source_excel_file_path = excelfile
    work_book = openpyxl.load_workbook(source_excel_file_path)
    work_sheet = work_book['Sheet1']
    list_ip_addr = []
    for actual_row_number in range(startrangeip, endrangeip + 1):
        value_cell = work_sheet.cell(row=actual_row_number, column=numarul_coloanei).value
        if cell_value is not None:
            list_ip_addr.append(value_cell)
    work_book.close()
    return list_ip_addr


def write_to_file(vision_box, destination_path_file):
    with open(destination_path_file, 'w') as file:
        file.write('#---------ETH0------------\n')
        file.write('{}\n'.format(vision_box.eth0[0]))
        file.write('# -------ETH0 Alias--------\n')
        for count in range(1, len(vision_box.eth0)):
            file.write('{}\n'.format(vision_box.eth0[count]))
        file.write('#---------ETH1------------\n')
        file.write('{}\n'.format(vision_box.eth1[0]))
        file.write('#---------ETH2------------\n')
        file.write('{}\n'.format(vision_box.eth2[0]))
        file.write('#---------ETH3------------\n')
        file.write('{}\n'.format(vision_box.eth3[0]))
        file.write('#---------ETH4------------\n')
        file.write('{}\n'.format(vision_box.eth4[0]))


def cautaindict(sursa, val):  # o folosesc pentru a gasi statiunile de pe carousel 1 si a le pune in ordine
    sursad = sursa
    value_to_find = val[:3]
    found_in_key = None
    # Parcurge fiecare cheie și lista asociată pentru a găsi valoarea
    for key, value_list in sursad.items():
        if value_to_find.upper() in value_list:
            found_in_key = key
            break
    if found_in_key is not None:
        return found_in_key
    else:
        print(f"Valoarea '{value_to_find}' nu a fost găsită în nicio listă asociată cu cheie.")


def create_tag_list(root_element, plcname, datatype):  # crearea tagului  este chemata mai sus in metoda writetag
    root = ET.Element(root_element)
    plc_var_name = ET.SubElement(root, "PLCVarName", AuditItemType="-1",
                                 Description="Name of variable defined in the PLC")
    plc_var_name.text = plcname
    db_number = ET.SubElement(root, "DBNumber", AuditItemType="-1", Description="Memory block")
    db_number.text = "9"
    byte_offset = ET.SubElement(root, "ByteOffset", AuditItemType="-1", Description="Byte Offset")
    byte_offset.text = "120"
    bit_offset = ET.SubElement(root, "BitOffset", AuditItemType="-1", Description="Bit Offset")
    bit_offset.text = "0"
    data_type = ET.SubElement(root, "Type", AuditItemType="-1", Description="Data type")
    data_type.text = datatype
    return root


excel_file_path = 'ConfigExcel.xlsm'
vbnr = 4
columnarray = []
initcol = 2
for i in range(0, 4):
    columnarray.append(initcol)
    initcol += 2
column_number = 2
startrangeIp = 5
endrangeIp = 13
dictIp = {}
vbarr = []
for i in range(0, vbnr):
    creare_vb = VisionBox(str(i+1))
    vbarr.append(creare_vb)
# Citirea ip-urilor din excel file i fiind coloana, de fiecare data pleaca la urmatoarea si cheama metoda readips
k = 0
for i in columnarray:
    dictIp['ETH0[IP]'] = readips('ConfigExcel.xlsm', 5, 13, i)
    dictIp['ETH1[IP]'] = readips('ConfigExcel.xlsm', 14, 14, i)
    dictIp['ETH2[IP]'] = readips('ConfigExcel.xlsm', 15, 15, i)
    dictIp['ETH3[IP]'] = readips('ConfigExcel.xlsm', 16, 16, i)
    dictIp['ETH4[IP]'] = readips('ConfigExcel.xlsm', 17, 17, i)
    vbarr[k].add_ip(dictIp)
    k += 1
###############################################################

for i in vbarr:
    i.showip()

write_to_file(vbarr[0], 'output.txt')

# deschidem fisierul excel si citim statiunile si ordinea lor
workbook = openpyxl.load_workbook(excel_file_path)
sheet = workbook['Sheet1']
statdict = {}
listorder = []
for row_number in range(23, 45 + 1):
    cell_value = sheet.cell(row=row_number, column=column_number).value
    if cell_value is not None:
        statdict[cell_value] = [sheet.cell(row=row_number, column=column_number-1).value, sheet.cell(row=row_number,
                                                        column=column_number+1).value]
        listorder.append(cell_value)
workbook.close()
#####################################################################
# deschidem fisierul si verificam fiecare statie daca se afla in lista statiilor de pe c1,c2,leak,single
workbook = openpyxl.load_workbook(excel_file_path)
sheet = workbook['Sheet2']
pozDict = {}
temp_list = []
for row_number in range(3, 25 + 1):
    cell_value = sheet.cell(row=row_number, column=1).value
    if cell_value is not None:
        temp_list.append(cell_value)
pozDict['C1'] = temp_list
temp_list = []
for row_number in range(3, 25 + 1):
    cell_value = sheet.cell(row=row_number, column=2).value
    if cell_value is not None:
        temp_list.append(cell_value)
pozDict['C2'] = temp_list
temp_list = []
for row_number in range(3, 25 + 1):
    cell_value = sheet.cell(row=row_number, column=3).value
    if cell_value is not None:
        temp_list.append(cell_value)
pozDict['Leak'] = temp_list
temp_list = []
for row_number in range(3, 25 + 1):
    cell_value = sheet.cell(row=row_number, column=4).value
    if cell_value is not None:
        temp_list.append(cell_value)
pozDict['Single'] = temp_list
workbook.close()
################################################################
# cream dictioanrul gen statia : nrspin
c1 = {}
k = 1
for i in listorder:
    if i[:3] in pozDict['C1']:
        c1[i] = round(k/2+0.1)
        k += 1
#########################################
# adaugam stattile, ordinea,pozitia(single,c1,c2,leak)
for vb in vbarr:
    for stname, orderlist in statdict.items():
        if str(orderlist[1]) == str(vb.name):
            poz = cautaindict(pozDict, stname)
            if poz == 'C1':
                vb.add_station(stname, orderlist[0], listorder.index(stname)+1, poz, str(c1[stname]))
            else:
                vb.add_station(stname, orderlist[0], listorder.index(stname) + 1, poz)
###############################################################
file_path = 'InstalParamVariable.csv'
data = []
with open(file_path, 'r', newline='', encoding='utf-8') as csvfile:  # Deschide fișierul CSV și citește conținutul
    csvreader = csv.DictReader(csvfile)
    for row in csvreader:    # Iterează prin rândurile din fișierul CSV
        data.append(row)
# Scrie variabilele in fisier
vbarr[0].writetags(data)
