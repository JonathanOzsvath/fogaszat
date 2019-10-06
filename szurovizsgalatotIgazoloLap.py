from time import gmtime, strftime
import pdfkit
import xlrd

def pdf_template():

    lastIndex = 548
    file = 'nevsor2.xlsx'
    workbook = xlrd.open_workbook(file)
    sheet = workbook.sheet_by_index(0)

    osztaly = sheet.cell_value(1, 1)

    parts = []
    felso_allando = """<html>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<style>
    html {nyitoZarojel}
        font-family: "Times New Roman", Georgia, Serif;
        letter-spacing: normal;
    {zaroZarojel}
    
    .kisSorkoz {nyitoZarojel}
        line-height: 3px;
        letter-spacing: 3px;
    {zaroZarojel}

    .oldalTores{nyitoZarojel}
        page-break-inside: avoid;
    {zaroZarojel}
</style>
<body>""".format(nyitoZarojel="{", zaroZarojel="}")

    also_allando = """</body>
</html>"""

    parts.append(felso_allando)

    for i in range(1, lastIndex):

        nev = sheet.cell_value(i, 2) + ' ' + sheet.cell_value(i, 3)

        s = """
<div class="oldalTores" style="border-right: thin">
    <div style="font-weight: bold; text-align: center;">
        Szűrővizsgálatot igazoló lap
    </div>
    <p>Igazolom, hogy <b>{nev}</b> fogorvosi szűrővizsgálaton megjelent.</p>
    <p>A fogászati állapota alapján további fogászati kezelés</p>
    <table style="margin-top: 0px; width: 100%;">
        <tr>
            <td>
                szükséges: &#x2610;
            </td>
            <td style="margin-left:40px; padding:0px 0px 0px 0px;  text-align: center">
                nem szükséges: &#x2610;
            </td>
        </tr>
    </table>
    <p>Hajdúhadház</p>
</div>
</body>
</html>""".format(nev=nev)

        parts.append(s)

        if i == lastIndex-1 or osztaly != sheet.cell_value(i+1, 1):
            parts.append(also_allando)
            kesz = "".join(parts)
            print(kesz, osztaly.replace(". ", "") + ".pdf")
            if i != lastIndex-1:
                osztaly = sheet.cell_value(i+1, 1)
                parts = []
                parts.append(felso_allando)

def print(kesz, pdf_name):

    option = {
        'page-size': 'A8',
        'margin-top': '0.1in',
        'margin-right': '0in',
        'margin-bottom': '0.1in',
        'margin-left': '0.1in',
        'encoding': 'UTF-8',
        'orientation': 'landscape'
    }

    pdfkit.from_string(kesz, pdf_name, options=option)

if __name__ == '__main__':
    pdf_template()