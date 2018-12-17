import re
import openpyxl
import os
import csv


def format_9x12(fle):
    header = ['AllSeq', 'PerSeq', 'WellmarkID', 'First',
              'Last', 'Add1', 'Add2', 'City', 'St', 'Zip',
              'Envelope', 'LtrDate', 'FormNum', 'Order']

    open_fle = openpyxl.load_workbook(filename=fle)
    with open("{0}.txt".format(fle[:-5]), 'w', newline='\n') as s:
        csvr = csv.writer(s, header, delimiter='\t', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        csvr.writerow(header)

        ws = open_fle.active
        wellid = ""
        perseq = 0
        for n, line in enumerate(ws.values):
            if line[0] == "AllSeq":
                pass
            else:
                if line[3] == wellid:
                    pass
                else:
                    wellid = line[3]
                    perseq += 1

                csvr.writerow([n, perseq, line[2], line[3], line[4], line[5],
                               line[6], line[7], line[8], line[9],
                               line[10], line[11], line[12], line[13], line[0]])

    open_fle.close()

def format_6x9(fle):
    header = ['AllSeq','PerSeq','WellmarkID','First',
              'Last','Add1','Add2','City','St','Zip',
              'Envelope','LtrDate','FormNum','Order','Page Ct']

    open_fle = openpyxl.load_workbook(filename=fle)
    with open("{0}.txt".format(fle[:-5]), 'w', newline='\n') as s:
        csvr = csv.writer(s, header, delimiter='\t', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        csvr.writerow(header)

        ws = open_fle.active
        wellid = ""
        perseq = 0
        for n, line in enumerate(ws.values):
            if line[0] == "Page Ct":
                pass
            else:
                if line[3] == wellid:
                    pass
                else:
                    wellid = line[3]
                    perseq += 1

                csvr.writerow([n, perseq, line[3], line[4], line[5],
                               line[6], line[7], line[8], line[9],
                               line[10], line[11], line[12], line[13],
                               line[14], line[0]])


    open_fle.close()

def process_file(fle):
    print(fle)
    find_6x9 = re.compile("6x9")
    find_9x12 = re.compile("9x12")

    # Run 6x9
    if re.search(find_6x9, fle):
        format_6x9(fle)
    # Run 9x12
    if re.search(find_9x12, fle):
        format_9x12(fle)


def main():
    xl = [f for f in os.listdir(os.curdir) if f[-4:].upper() == 'XLSX']
    for n, f in enumerate(xl):
        process_file(f)

if __name__ == '__main__':
    main()
