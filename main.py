
from openpyxl import load_workbook


mappings = {
    'RCGIS DROPS': 'DROPS',
    'DIALOGTJENESTE-API': 'DIALOGTJENESTE',
    'DIALOGTJENESTE-UT':  'DIALOGTJENESTE',
    'IFS': 'IFS CLOUD',
    'IFSCLOUD': 'IFS CLOUD',
    'MCPS-INNBETALINGSFILER': 'MCPS',
    'MCPS-REMITTERINGSFILER': 'MCPS',
    'NETBAS-MÅLEPUNKT-UT': 'NETBAS',
    'NETBAS-SAMLESKINNENAVN (SØR)': 'NETBAS',
    'NETBAS-TILKNYTNINGSPUNKT-UT': 'NETBAS',
    'STATNETT-MARGINALTAPSNAVN (NORD)': 'STATNETT',
    'STATNETT-MARGINALTAPSSATSER (NORD)': 'STATNETT',
    'UN': 'ARCGIS'
}

#versions = ['DV4/5', 'DV6', 'DV7']
#versions = ['DV7']
versions = ['DV4/5']


def get_apis(api_string):
    funcs = []
    a_name = ''
    split_position = 0
    the_name = ''
    apis = []
    if type(api_string) == str:
        if ',' in api_string:
            funcs = list(map(str.strip, api_string.split(',')))
        else:
            #                print(api)
            funcs = [api_string]
        for a_name in funcs:
            a_name = a_name.upper()
            if '-' in a_name:
                split_position = a_name.index('-')
                the_name = a_name[:split_position]
                if the_name in mappings:
                    the_name = mappings[the_name]
            #                    if name not in systems:
            #                       name = api_name
            else:
                the_name = a_name
                if the_name in mappings:
                    the_name = mappings[the_name]
            if the_name != '':
                apis.append(the_name)
        if len(apis) >= 1:
            return apis
        else:
            return None
    return None


if __name__ == '__main__':

    EXCEL_WORKBOOK_NAME = '/users/djr/Downloads/DV4_Status_Integrasjoner (3).xlsx'
    EXCEL_SHEET_NAME = 'Integrasjon-Dataflyt status'

    workbook = load_workbook(EXCEL_WORKBOOK_NAME)
    worksheet = workbook[EXCEL_SHEET_NAME]

    systems = set()
    startrow = 2
    endrow = worksheet.max_row + 1
    name = ''
    version = ''
    for i in range(startrow, endrow):
        version = worksheet.cell(row=i, column=29).value
        if not version or version == '' or version in versions:
            name = worksheet.cell(row=i, column=1).value
            systems.add(name.upper())

    sorted_systems = list(sorted(systems))
#    print(sorted_systems)

    dependent_on_apis = set()
    functions = []
    api = ''
    for i in range(startrow, endrow):
        functions = get_apis(worksheet.cell(row=i, column=4).value)
        if functions:
            for api in functions:
                dependent_on_apis.add(api)


    sorted_apis = list(sorted(dependent_on_apis))
#    print(sorted(sorted_apis))

    integrations_sheet = workbook.create_sheet(title='Integrasjoner')

    the_row = 1
    the_col = 2
    for name in sorted_apis:
        integrations_sheet.cell(row=the_row, column=the_col).value = name
        the_col += 1

    the_row = 2
    the_col = 1
    for name in sorted_systems:
        integrations_sheet.cell(row=the_row, column=the_col).value = name
        the_row += 1

    for i in range(startrow, endrow):
        version = worksheet.cell(row=i, column=29).value
        if not version or version == '' or version in versions:
            name = worksheet.cell(row=i, column=1).value
            name = name.upper()
            functions = get_apis(worksheet.cell(row=i, column=4).value)
            if functions:
                for api in functions:
                    integrations_sheet.cell(row=sorted_systems.index(name)+2, column=sorted_apis.index(api)+2).value = 'X'

    workbook.save(filename='/users/djr/Downloads/foo.xlsx')



