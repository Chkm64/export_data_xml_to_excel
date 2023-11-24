import xlsxwriter

class Excel:
    def __init__(self, name, data):
        self.name = name
        self.data = data

    def initWorbook(self):
        workbook = xlsxwriter.Workbook('files/'+str(self.name)+'.xlsx')
        return workbook

    def createWorsheet(self, workbook):
        worksheet = workbook.add_worksheet()
        return worksheet

    def setHeaders(self):
        headers = self.data[0].keys()
        return headers

    def write(self, worksheet, row, column, value, attr):
        worksheet.write(row, column, value, attr)

    def generate(self):
        # init
        workbook = self.initWorbook()
        worksheet = self.createWorsheet(workbook)
        # header
        style_header = workbook.add_format({'bold': True})
        header = self.setHeaders()
        for count, item in enumerate(header):
            worksheet.set_column(0, count, 42)
            self.write(worksheet, 0, count, item, style_header)
        # body
        style_body = workbook.add_format({'bold': False})
        for row, item in enumerate(self.data):
            for colum, colum_name in enumerate(header):
                self.write(worksheet, row+1, colum, item[colum_name], style_body)
        workbook.close()
        return True
