import pandas as pd
from create_pdf import PDF  


# Variables for accessing
FILE_NAME = 'exfile.xlsx'
MOST_RECENT_SHEET = -1

class Aqua():
    def __init__(self, collar_count = 0, leash_count = 0, harness_count = 0) -> None:
        self.collar_count = collar_count
        self.leash_count = leash_count
        self.harness_count = harness_count

class Cat():
    def __init__(self, collar_count = 0) -> None:
        self.collar_count = collar_count

class Ikonic():
    def __init__(self, collar_count = 0, leash_total = 0,harness_total = 0, martingal_total = 0, metal_buckle_total = 0, 
                customized_total = 0, h_harness_total = 0, fabric_martingale = 0 ) -> None:
        self.collar_total = collar_count
        self.leash_total = leash_total
        self.harness_total = harness_total
        self.martingal_total = martingal_total
        self.metal_buckle_total = metal_buckle_total
        self.customized_total = customized_total
        self.h_harness_total = h_harness_total
        self.fabric_martingale = fabric_martingale
    
def get_sheet_names(file_name):
    xl_file = pd.ExcelFile(file_name)
    return xl_file.sheet_names

def process_aqua_count(file_name) -> Aqua:
    sheet_names = get_sheet_names(file_name)
    xl_df = pd.read_excel(file_name, sheet_name=sheet_names[MOST_RECENT_SHEET])
    collar_count = leash_count = harness_count = 0
    for i in range(2, 6):
        print()
        collar_count += xl_df.iloc[80][i]
    print(collar_count)
    for i in range(6, 8):
        leash_count += xl_df.iloc[80][i]
    print(leash_count)
    for i in range(8,12):
        harness_count += xl_df.iloc[80][i]
    print(harness_count)
    aqua : Aqua = Aqua(collar_count=collar_count, 
                       leash_count=leash_count, 
                       harness_count=harness_count)
    return aqua

def process_ikonic_count(file_name) -> Ikonic:
    sheet_names = get_sheet_names(file_name)
    xl_df = pd.read_excel(file_name, sheet_name=sheet_names[MOST_RECENT_SHEET])
    # collars = 2 - 5 column 
    collar_count = leash_total = harness_total =  martingale_total = \
    metal_buckle_total = customized_total = h_harness_total = \
    fabric_martingale = 0
    for i in range(2, 6):
        collar_count += xl_df.iloc[68][i]
    for i in range(6, 8):
        leash_total += xl_df.iloc[68][i]
    for i in range(8,12):
        harness_total += xl_df.iloc[68][i]
    for i in range(12,16):
        martingale_total += xl_df.iloc[68][i]
    for i in range(16, 20):
        metal_buckle_total += xl_df.iloc[68][i]
    for i in range(20,22):
        customized_total += xl_df.iloc[68][i]
    for i in range(22,26):
        h_harness_total += xl_df.iloc[68][i]
    for i in range(26,29):
        fabric_martingale += xl_df.iloc[68][i]
    ikonic : Ikonic = Ikonic(collar_count=collar_count, leash_total=leash_total, harness_total=harness_total, 
    martingal_total=martingale_total, metal_buckle_total=metal_buckle_total, customized_total=customized_total, 
    h_harness_total=h_harness_total, fabric_martingale=fabric_martingale)
    return ikonic

def process_cat_count(file_name) -> Cat:
    sheet_names = get_sheet_names(file_name)
    xl_df = pd.read_excel(file_name, sheet_name=sheet_names[MOST_RECENT_SHEET])
    pattern_position_start = 7
    index = 40
    total = 0
    while not pd.isna(xl_df.iloc[pattern_position_start][38]):
        if not pd.isna(xl_df.iloc[pattern_position_start][40]):
            total += xl_df.iloc[pattern_position_start][40]
        pattern_position_start += 1
    cat : Cat = Cat(collar_count=total)
    return cat

    
if __name__ == "__main__":
    FILE_NAME = "exfile.xlsx"
    sheet_names = get_sheet_names(FILE_NAME)
    xl_df = pd.read_excel(FILE_NAME, sheet_name=sheet_names[MOST_RECENT_SHEET])
    aqua : Aqua = process_aqua_count(FILE_NAME)
    cat : Cat = process_cat_count(FILE_NAME)
    ikonic : Ikonic = process_ikonic_count(FILE_NAME)
    # Instantiation of inherited class
    pdf = PDF()
    pdf.alias_nb_pages()
    pdf.add_page()
    pdf.set_font('Times', '', 12)
    #adding Aqua values
    pdf.cell(0, 10, '', 0, 1)
    #adding Ikonic values
    pdf.cell(0, 10, 'Report for date ' + str(sheet_names[MOST_RECENT_SHEET]) , 0, 1)
    pdf.cell(0, 10, 'Aqua Orders', 0, 1)
    pdf.cell(0, 10, '                   Collar Orders : ' + str(aqua.collar_count), 0, 1)
    pdf.cell(0, 10, '                   Leash Orders : ' + str(aqua.leash_count), 0, 1)
    pdf.cell(0, 10, '                   Harness Orders : ' + str(aqua.leash_count), 0, 1)
    pdf.cell(0, 10, 'Ikonic Orders', 0, 1)
    pdf.cell(0, 10, '                   Collar Orders : ' + str(ikonic.collar_total), 0, 1)
    pdf.cell(0, 10, '                   Leash Orders : ' + str(ikonic.leash_total), 0, 1)
    pdf.cell(0, 10, '                   Harness Orders : ' + str(ikonic.harness_total), 0, 1)
    pdf.cell(0, 10, '                   Martingale Orders : ' + str(ikonic.martingal_total), 0, 1)
    pdf.cell(0, 10, '                   Metal Buckle Collar Orders : ' + str(ikonic.metal_buckle_total), 0, 1)
    pdf.cell(0, 10, '                   Customized Leash Orders : ' + str(ikonic.customized_total), 0, 1)
    pdf.cell(0, 10, '                   H-Harness Orders : ' + str(ikonic.h_harness_total), 0, 1)
    pdf.cell(0, 10, '                   Fabric Orders : ' + str(ikonic.fabric_martingale), 0, 1)
    pdf.cell(0, 10, 'Cat Collar Orders', 0, 1)
    pdf.cell(0, 10, '                   Collar Orders : ' + str(cat.collar_count), 0, 1)
    pdf.output('report.pdf', 'F')
    print("created")
