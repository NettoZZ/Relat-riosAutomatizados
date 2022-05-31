import win32com.client
from datetime import datetime
import os

drive_letter = r'F:\\' 
folder_name = r'Docs\Relatórios\Relatórios\Relatório '  
folder_time = datetime.now().strftime("%d-%m-%Y")
pdf_itau = r"\Divergencias CC ITAU com CxCentral"
pdf_itau2 = r"\Divergencias CC ITAU x CxCentral"
pdf_santander2 = r"\Divergencias CC Santander PJ x CxCentral"
pdf_global = r"\Divergencias Global - Geral"
pdf_santander = r"\Divergencias Santander com CxCentral"
pdf_periodicas = r"\Relatorio - tarefas periodicas"
pdf_sangrias = r"\Sangrias com divergencias"
pdf_rendimento = r"\Relatório Bitrix - Todos funcionários"
pdf_metas = r"\Relatório de roteiros - Todos funcionários"
pdf_borrador = r"\Indicadores de prazos dos qualificadores"
pdf_bitrix = r"\Relatórios Bitrix"
folder_to_save_files = drive_letter + folder_name + folder_time
path = folder_to_save_files
path = os.path.realpath(path)
if not os.path.exists(folder_to_save_files):
    os.mkdir(folder_to_save_files)

################################----INSTRUÇÕES

os.system("cls")
print('\n  ############################################ INTRUÇÕES DE USO ############################################')
print("  #                                                                                                        #")
print("  #    1- Os arquivos serão gerados em F:\Docs\Relatórios\Relatórios                                       #")
print("  #    2- Em caso de erro certifique-se que ninguem está com alguma planilha dos relatórios aberta         #")
print("  #    3- Em caso de movimentação das planilhas Excel informar o T.I para alterar o caminho no programa    #")
print("  #    4- Criado e desenvolvido por Antonio Netto - Gerador de relatórios Excel - BOM USO! :D              #")
print("  #                                                                                                        #")
print("  #    Pressione qualquer tecla para começar a gerar os relatórios...                                      #")
print("  #                                                                                                        #")
input("  ##########################################################################################################")
os.system("cls")
print('\n  ##### RELATÓRIOS SENDO GERADOS #####')
print("  #                                  #")
print("  #    |                    | 0%     #")
print("  #                                  #")
print("  ####################################")
    
################################----RENDIMENTO

o = win32com.client.Dispatch("Excel.Application")
o.Visible = False
o.DisplayAlerts = False
o.EnableEvents = False
o.ScreenUpdating = False
wb_pathrendimento = r'F:\Docs\Bitrix\Relatório Bitrix.xlsx'
wbrendimento = o.Workbooks.Open(wb_pathrendimento, False, True, None, None, None, True)
o.Calculation = -4135
ws_index_listrendimento = [3]
path_to_pdf = folder_to_save_files + pdf_rendimento
print_area = 'A1:U164'
for index in ws_index_listrendimento:
    ws = wbrendimento.Worksheets[index - 1]
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.PrintArea = print_area
wbrendimento.WorkSheets(ws_index_listrendimento).Select()
wbrendimento.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
wbrendimento.Close(False)
os.system("cls")
print('\n  ##### RELATÓRIOS SENDO GERADOS #####')
print("  #                                  #")
print("  #    |--                  | 10%    #")
print("  #                                  #")
print("  ####################################")

################################----METAS

wb_pathmetas = r'F:\Docs\Relatório Roteiros\Relatório Roteiros.xlsx'
wbmetas = o.Workbooks.Open(wb_pathmetas, False, True, None, None, None, True)
o.Calculation = -4135
ws_index_listmetas = [1]
path_to_pdf = folder_to_save_files + pdf_metas
print_area = 'A1:M1443'
for index in ws_index_listmetas:
    ws = wbmetas.Worksheets[index - 1]
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.PrintArea = print_area
wbmetas.WorkSheets(ws_index_listmetas).Select()
wbmetas.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
wbmetas.Close(False)
os.system("cls")
print('\n  ##### RELATÓRIOS SENDO GERADOS #####')
print("  #                                  #")
print("  #    |----                | 20%    #")
print("  #                                  #")
print("  ####################################")

################################----BORRADOR

wb_pathborrador = r'F:\Docs\RI\Borrador do recolhimento.xlsm'
wbborrador = o.Workbooks.Open(wb_pathborrador, False, True, None, None, None, True)
o.Calculation = -4135
ws_index_listborrador = [5]
path_to_pdf = folder_to_save_files + pdf_borrador
print_area = 'A1:D35'
for index in ws_index_listborrador:
    ws = wbborrador.Worksheets[index - 1]
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.PrintArea = print_area
wbborrador.WorkSheets(ws_index_listborrador).Select()
wbborrador.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
wbborrador.Close(False)
os.system("cls")
print('\n  ##### RELATÓRIOS SENDO GERADOS #####')
print("  #                                  #")
print("  #    |------              | 30%    #")
print("  #                                  #")
print("  ####################################")

################################----PERIODICAS

wb_pathperiodicas = r'\\10.1.1.2\f\Docs\POP\Tarefas Periódicas Centralizadas (USAR ESTA).xlsx'
wbperiodicas = o.Workbooks.Open(wb_pathperiodicas, False, True, None, None, None, True)
o.Calculation = -4135
ws_index_listperiodicas = [6]
path_to_pdf = folder_to_save_files + pdf_periodicas
wbperiodicas.ActiveSheet.Columns("C:CZ").Hidden = True
print_area = 'A1:DV20'
for index in ws_index_listperiodicas:
    ws = wbperiodicas.Worksheets[index - 1]
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.PrintArea = print_area
wbperiodicas.WorkSheets(ws_index_listperiodicas).Select()
wbperiodicas.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
os.system("cls") 
print('\n  ##### RELATÓRIOS SENDO GERADOS #####')
print("  #                                  #")
print("  #    |--------            | 40%    #")
print("  #                                  #")
print("  ####################################")

################################----GLOBAL
################################----SANGRIAS

wb_pathglobal = r'F:\Docs\financeiro\Global\PJGlobal_numero.xlsm'
wb = o.Workbooks.Open(wb_pathglobal, False, True, None, 'woner606', None, True)
o.Calculation = -4135
ws_index_listglobal = [8]
path_to_pdf = folder_to_save_files + pdf_sangrias
o.Selection.AutoFilter(8, "<>0,00")
wb.ActiveSheet.Columns("D:E").Hidden = True
print_area = 'A4:H375'
for index in ws_index_listglobal:
    ws = wb.Worksheets[index - 1]
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.PrintArea = print_area
wb.WorkSheets(ws_index_listglobal).Select()
wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
o.Selection.AutoFilter(8)
os.system("cls") 
print('\n  ##### RELATÓRIOS SENDO GERADOS #####')
print("  #                                  #")
print("  #    |----------          | 50%    #")
print("  #                                  #")
print("  ####################################")

################################----ITAU1

path_to_pdf = folder_to_save_files + pdf_itau
o.Selection.AutoFilter(17, "<>0,00")
wb.ActiveSheet.Columns("D:N").Hidden = True
print_area = 'A4:Q375'
for index in ws_index_listglobal:
    ws = wb.Worksheets[index - 1]
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.PrintArea = print_area
wb.WorkSheets(ws_index_listglobal).Select()
wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
o.Selection.AutoFilter(17)
os.system("cls")
print('\n  ##### RELATÓRIOS SENDO GERADOS #####')
print("  #                                  #")
print("  #    |------------        | 60%    #")
print("  #                                  #")
print("  ####################################")

################################----SANTANDER1

path_to_pdf = folder_to_save_files + pdf_santander
o.Selection.AutoFilter(23, "<>0,00")
wb.ActiveSheet.Columns("D:T").Hidden = True
print_area = 'A4:W375'
for index in ws_index_listglobal:
    ws = wb.Worksheets[index - 1]
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.PrintArea = print_area
wb.WorkSheets(ws_index_listglobal).Select()
wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
o.Selection.AutoFilter(23)
os.system("cls")
print('\n  ##### RELATÓRIOS SENDO GERADOS #####')
print("  #                                  #")
print("  #    |--------------      | 70%    #")
print("  #                                  #")
print("  ####################################")

################################----ITAU2

path_to_pdf = folder_to_save_files + pdf_itau2
o.Selection.AutoFilter(31, "<>0,00")
wb.ActiveSheet.Columns("D:AB").Hidden = True
print_area = 'A4:AE375'
for index in ws_index_listglobal:
    ws = wb.Worksheets[index - 1]
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.PrintArea = print_area
wb.WorkSheets(ws_index_listglobal).Select()
wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
o.Selection.AutoFilter(31)
os.system("cls")
print('\n  ##### RELATÓRIOS SENDO GERADOS #####')
print("  #                                  #")
print("  #    |----------------    | 80%    #")
print("  #                                  #")
print("  ####################################")

################################----SANTANDER2

path_to_pdf = folder_to_save_files + pdf_santander2
o.Selection.AutoFilter(34, "<>0,00")
wb.ActiveSheet.Columns("D:AE").Hidden = True
print_area = 'A4:AH375'
for index in ws_index_listglobal:
    ws = wb.Worksheets[index - 1]
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.PrintArea = print_area
wb.WorkSheets(ws_index_listglobal).Select()
wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
o.Selection.AutoFilter(34)
os.system("cls") 
print('\n  ##### RELATÓRIOS SENDO GERADOS #####')
print("  #                                  #")
print("  #    |------------------  | 90%    #")
print("  #                                  #")
print("  ####################################")

################################----GERAL

ws_index_listgeral = [4]
path_to_pdf = folder_to_save_files + pdf_global
print_area = 'A1:I39'
for index in ws_index_listgeral:
    ws = wb.Worksheets[index - 1]
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.PrintArea = print_area
wb.WorkSheets(ws_index_listgeral).Select()
wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
wb.Close(False)
os.system("cls")
print('\n  ##### RELATÓRIOS SENDO GERADOS #####')
print("  #                                  #")
print("  #    |--------------------| 100%   #")
print("  #                                  #")
print("  ####################################")

################################----BITRIX

# wb_pathbitrix = r'F:\Docs\Relatório Bitrix\Relatório Bitrix.xlsm'
# wbbitrix = o.Workbooks.Open(wb_pathbitrix, False, True, None, None, None, True)
# o.Calculation = -4135
# ws_index_listbitrix = [1]
# path_to_pdf = folder_to_save_files + pdf_bitrix
# print_area = 'A:I'
# for index in ws_index_listbitrix:
#     ws = wbperiodicas.Worksheets[index - 1]
#     ws.PageSetup.Zoom = False
#     ws.PageSetup.FitToPagesTall = 1
#     ws.PageSetup.FitToPagesWide = 1
#     ws.PageSetup.PrintArea = print_area
# wbbitrix.WorkSheets(ws_index_listbitrix).Select()
# wbbitrix.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
# wb.Close(False)
os.system("cls")
print('\n  ##### RELATÓRIOS GERADOS COM SUCESSO! #####')
print("  #                                         #")
print("  #    Pressione enter para finalizar...    #")
print("  #                                         #")
input("  ###########################################")
os.startfile(path)