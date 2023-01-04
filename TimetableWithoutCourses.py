import pdfplumber


with pdfplumber.open('') as pdf:
    tables = [page.extract_table() for page in pdf.pages]
    for i in range(len(tables) - 1):
        if (tables[i][-1][1] == '' or tables[i+1][0][1] == ''):
            for col in range(len(tables[i+1][0])):
                if (tables[i+1][0][col] != ''):
                    rowOff = 0
                    while tables[i][-1-rowOff][col] == None:
                        rowOff += 1
                    tables[i][-1-rowOff][col] += tables[i+1][0][col]
            tables[i+1] = tables[i+1][1:]
    table = []
    for t in tables:
        table += t
    for row in table:
        print(row)
