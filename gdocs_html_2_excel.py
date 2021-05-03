import requests
from bs4 import BeautifulSoup
import pandas

base_download_url = 'https://docs.google.com/spreadsheets/u/0/d/'
spreadsheet_id = '1LUtqhOEjUMySCfn3zj8Arhzcmazr3vrPzy7VzJwIshE'
force_html_view = '/htmlview#'

excel_name = None

sheet_map = {}
data_map = {}

with requests.get(base_download_url + spreadsheet_id + force_html_view) as web_page:

    if not web_page.status_code == 200:
        print('Page did not load')
        print('Status Code :: {}'.format(web_page.status_code))
        print('Reason      :: {}'.format(web_page.reason))
    else:
        parsed_page = BeautifulSoup(web_page.content, 'html.parser')
        excel_name = parsed_page.title.text
        sheet_menu = parsed_page.find_all(id='sheet-menu')
        for sheet in sheet_menu[0].children:
            if sheet.get('id') is not None:
                sheet_id = sheet.get('id').replace('sheet-button-', '')
                sheet_map[sheet_id] = sheet.text
                data_map[sheet_id] = pandas.read_html(str(parsed_page.find_all(id=sheet_id)[0]), index_col=0,
                                                      skiprows=1)[0]
            else:
                pass

if excel_name is None or len(sheet_map) == 0 or len(data_map) == 0 or len(sheet_map) != len(data_map):
    print('Excel Extraction for {} (id="{}") was not successful...'.format(excel_name, spreadsheet_id))
    print('We extracted {} sheets based on the metadata and found {} table data in the body'.format(len(sheet_map),
                                                                                                    len(data_map)))
    print('Extracted sheets are below')
    print(sheet_map)
else:
    with pandas.ExcelWriter(str(excel_name) + '.xlsx') as excel_writer:
        for sheet_id in sheet_map.keys():
            sheet_data = data_map[sheet_id]
            sheet_data.to_excel(excel_writer, sheet_name=sheet_map[sheet_id], index=False)
