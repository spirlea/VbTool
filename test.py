
import xml.etree.ElementTree as ET

# Crează un element radacina
root = ET.Element("CurrentSpeed")

# Adaugă subelementele cu valorile corespunzătoare
plc_var_name = ET.SubElement(root, "PLCVarName", AuditItemType="-1", Description="Name of variable defined in the PLC")
plc_var_name.text = "VRI_Definitions.VRI_SetSpeed"

db_number = ET.SubElement(root, "DBNumber", AuditItemType="-1", Description="Memory block")
db_number.text = "9"

byte_offset = ET.SubElement(root, "ByteOffset", AuditItemType="-1", Description="Byte Offset")
byte_offset.text = "120"

bit_offset = ET.SubElement(root, "BitOffset", AuditItemType="-1", Description="Bit Offset")
bit_offset.text = "0"

data_type = ET.SubElement(root, "Type", AuditItemType="-1", Description="Data type")
data_type.text = "3"

# Creează un arbore XML
tree = ET.ElementTree(root)

# Salvează arborele XML într-un fișier
tree.write("output.xml")