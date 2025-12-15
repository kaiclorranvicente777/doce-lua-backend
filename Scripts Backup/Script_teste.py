import win32com.client as win32
from PIL import ImageGrab

arquivo = r"C:\Users\Káic e Gabi\OneDrive\Documentos\Sistema.Central_Doce.Lua_09-2025.xlsx"

try:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(arquivo)
    ws = wb.Sheets('Acompanhamento-Geral')

    ws.Range("I1:AP27").CopyPicture(Format=win32.constants.xlPicture)
    temp_sheet = wb.Sheets.Add()
    temp_sheet.Paste()

    temp_sheet.Range("A1").CopyPicture(Format=win32.constants.xlPicture)
    image = ImageGrab.grabclipboard()

    imagem_path = r"C:\Users\Káic e Gabi\OneDrive\Documentos\calendario_teste.png"
    image.save(imagem_path)

    temp_sheet.Delete()
    wb.Close(SaveChanges=False)
    excel.Quit()

    print("Imagem capturada com sucesso!")
except Exception as e:
    print(f"Erro: {e}")
