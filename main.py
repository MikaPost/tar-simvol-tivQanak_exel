"""
This file is for our new theme: tar simvol tiv Qanak in exel
Create by: Miqayel Postoyan
Date: 15 April
"""
import xlsxwriter

def get_content(fname):
    with open(fname) as f:
        return f.read()

def create_list_of_names(cnt):
    letter = {}
    nam = {}
    simvol = {}
    for i in cnt:
        if i.isalpha():
            if i in letter:
                letter[i] += 1
            else:
                letter[i] = 1
        elif i.isdigit():
            if i in nam:
                nam[i] += 1
            else:
                nam[i] = 1
        else:
            if i in simvol:
                simvol[i] += 1
            else:
                simvol[i] = 1
    return list(letter.items()), list(nam.items()), list(simvol.items())


def writer_excel(letter_list, num_list, simvol_list):
    workbook = xlsxwriter.Workbook("mika.xlsx")
    letter_sheet = workbook.add_worksheet("Letters")
    num_sheet = workbook.add_worksheet("Numbers")
    simvol_sheet = workbook.add_worksheet("Symbols")

    write_sheet(letter_sheet, letter_list, True)
    write_sheet(num_sheet, num_list)
    write_sheet(simvol_sheet, simvol_list)

    workbook.close()

def write_sheet(sheet, data_list, n=None):
    row = 0
    if n:
        vowels, consonants = sort_tarer(data_list)
        row = 1
        for i in vowels:
            sheet.write(0, 0, "vowels")
            sheet.write(row, 0, i[0])
            sheet.write(row, 1, i[1])
            row += 1
        sheet.write(row, 0, "consonants")
        row += 1
        for i in consonants:
            sheet.write(row, 0, i[0])
            sheet.write(row, 1, i[1])
            row += 1
    else:
        for item in data_list:
            sheet.write(row, 0, item[0])
            sheet.write(row, 1, item[1])
            row += 1

def sort_tarer(data_list):
    a = "aeiouy"
    vowels = []
    consonants = []
    for i in range(len(data_list)):
        if data_list[i][0] in a:
            vowels.append(data_list[i])
        else:
            consonants.append(data_list[i])
    return vowels, consonants



def main():
    cnt = get_content("a.txt")
    letter_list, num_list, simvol_list = create_list_of_names(cnt)
    letter_list.sort(key=lambda x: x[1], reverse=True)
    num_list.sort(key=lambda x: x[1], reverse=True)
    simvol_list.sort(key=lambda x: x[1], reverse=True)
    writer_excel(letter_list, num_list, simvol_list)


if __name__ == "__main__":
    main()
