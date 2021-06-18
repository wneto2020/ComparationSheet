import gspread


class Sup:
    def __init__(self, importation, solicitation):
        self.solicitation = self.request_sheet(solicitation, 'Página1', 'dict')
        self.importation = self.request_sheet(importation, 'Página1', 'dict')
        self.median_dict = self.request_sheet(importation, 'intermediaria', 'dict')
        self.median_list = self.request_sheet(importation, 'intermediaria', 'list')
        self.send = self.request_sheet(importation, 'envio')
        self.list = []

    @staticmethod
    def request_sheet(spreadsheet, sheet, method=None):
         """The function requests a spreadsheet,
            returning the contents of your chosen sheet and the return method
        """
        gc = gspread.service_account('service_account.json')
        sh = gc.open(spreadsheet)
        work = sh.worksheet(sheet)

        if method == 'dict':
            return work.get_all_records()
        elif method == 'list':
            return work.get_all_values()
        else:
            return work

    def comparison(self):
        """Function of comparison spreadsheets importation and solicitation
        """
        for person_solicit in self.solicitation:
            for person_import in self.importation:
                if person_solicit['Nome'] == person_import['Nome'] or person_solicit['E-mail'] == person_import['E-mail']:
                    self.list.append(person_solicit)
        return self.list

    def write(self, spreadsheet, sheet, add: list):
         """The function write in a sheet
           not return
        """
        writer = self.request_sheet(spreadsheet, sheet)
        writer.append_row(add)

    def update(self, spreadsheet, sheet, cont):
        """Update cell"""
        
        update_cell = self.request_sheet(spreadsheet, sheet)
        update_cell.update(f'D{cont}', 'Enviado')

    def update_send(self):
        """Function update the spreadsheet on key situation and send data from other spreadsheet"""
        
        i = 2
        for item in self.median_dict:
            if item['situacao'] == 'Enviado':
                print("enviado")

            elif item['situacao'] != 'Enviado':
                self.write('importada', 'envio', [item['Nome'], item['E-mail'], item['cpf']])
                self.update('importada', 'intermediaria', i)
            i += 1

    def run(self):
        """Execute POO (initialize)
        """

        list_comparison = self.comparison()

        for person_comparison in list_comparison:
            i = 0
            for item in self.median_list:
                if (person_comparison['Nome'] or person_comparison['E-mail']) in item:
                    print('Contido')
                    continue
                else:
                    i += 1
            if i == len(self.median_list):
                self.write('importada', 'intermediaria', [person_comparison['Nome'], person_comparison['E-mail'],
                                                          person_comparison['cpf']])
                print('Adicionado')


if __name__ == '__main__':
    # importation = input('import: ')
    # solicitation = input('solicit: ')
    sup = Sup('importada', 'solicitacao')
    sup.run()
    sup.update_send()
