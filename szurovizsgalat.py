from time import gmtime, strftime
import pdfkit
import xlrd

LAST_INDEX = 538

def pdf_template():
    file = 'Nevsor_2023_04_25.xlsx'
    workbook = xlrd.open_workbook(file)
    sheet = workbook.sheet_by_index(0)

    osztaly = sheet.cell_value(1, 0)

    parts = []
    felso_allando = """<html>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<style>
    html {nyitoZarojel}
        font-family: "Times New Roman", Georgia, Serif;
        letter-spacing: normal;
    {zaroZarojel}

    .oldalTores{nyitoZarojel}
        page-break-inside: avoid;
    {zaroZarojel}
    
    .kisSorkoz{nyitoZarojel}
        margin: 5px;
        margin-left: 0px;
    {zaroZarojel}
    
    td{nyitoZarojel}
        padding-right: 15px;
    {zaroZarojel}
    
    .boarder-bottom{nyitoZarojel}
        border-bottom: solid;
        border-width: thin;
    {zaroZarojel}
    
    .boarder-right{nyitoZarojel}
        border-right: solid;
        border-width: thin;
    {zaroZarojel}
</style>
<body>""".format(nyitoZarojel="{", zaroZarojel="}")

    also_allando = """</body>
</html>"""

    parts.append(felso_allando)

    for i in range(1, LAST_INDEX):
        nev = sheet.cell_value(i, 1) + ' ' + sheet.cell_value(i, 2)
        anev = sheet.cell_value(i, 3) + ' ' + sheet.cell_value(i, 4)
        lakhely = sheet.cell_value(i, 8) + ' ' + sheet.cell_value(i, 9) + ' ' + sheet.cell_value(i, 10) + ' ' + sheet.cell_value(i, 11) + ' ' + sheet.cell_value(i, 12)
        szul_datum = sheet.cell_value(i, 5)
        taj = sheet.cell_value(i, 7)

        s = """<div class="oldalTores">
    <div style="font-weight: bold; text-align: center;">
        Szűrővizsgálati igazoló lap
    </div>
    <p class="kisSorkoz">Név: {nev}</p>
    <p class="kisSorkoz">Anyja neve: {anev}</p>
    <p class="kisSorkoz">Születési dátum: {szul_datum}</p>
    <p class="kisSorkoz">TAJ szám: {taj}</p>
    <p class="kisSorkoz">Lakhely: {lakhely}</p>
    <p>Igazolom, hogy a páciens szűrővizsgálaton megjelent.</p>
    <p>A fogazati állapota alapján további fogászati kezelés</p>
    <table style="width: 100%">
        <tr>
            <td>
                szükséges: &#x2610;
            </td>
            <td>
                nem szükséges: &#x2610;
            </td>
        </tr>
    </table>
    <p>Szakorvosi ellátés szükséges:</p>
    <p class="kisSorkoz">Szájsebészet:........................................................................&#x2610;</p>
    <p class="kisSorkoz">Paradontológia:....................................................................&#x2610;</p>
    <p class="kisSorkoz">Fogpótlás:............................................................................&#x2610;</p>
    <p class="kisSorkoz">Fogszabályozás:...................................................................&#x2610;</p>
    <table style="margin-top: 50px; border-spacing: 0px;">
        <tr>
            <td class="boarder-bottom boarder-right">8</td>
            <td class="boarder-bottom boarder-right">7</td>
            <td class="boarder-bottom boarder-right">6</td>
            <td class="boarder-bottom boarder-right">5</td>
            <td class="boarder-bottom boarder-right">4</td>
            <td class="boarder-bottom boarder-right">3</td>
            <td class="boarder-bottom boarder-right">2</td>
            <td class="boarder-bottom boarder-right">1</td>
             
            <td class="boarder-bottom boarder-right">1</td>
            <td class="boarder-bottom boarder-right">2</td>
            <td class="boarder-bottom boarder-right">3</td>
            <td class="boarder-bottom boarder-right">4</td>
            <td class="boarder-bottom boarder-right">5</td>
            <td class="boarder-bottom boarder-right">6</td>
            <td class="boarder-bottom boarder-right">7</td>
            <td class="boarder-bottom">8</td>
        </tr>
        <tr>
            <td class="boarder-right">8</td>
            <td class="boarder-right">7</td>
            <td class="boarder-right">6</td>
            <td class="boarder-right">5</td>
            <td class="boarder-right">4</td>
            <td class="boarder-right">3</td>
            <td class="boarder-right">2</td>
            <td class="boarder-right">1</td>
            
            <td class="boarder-right">1</td>
            <td class="boarder-right">2</td>
            <td class="boarder-right">3</td>
            <td class="boarder-right">4</td>
            <td class="boarder-right">5</td>
            <td class="boarder-right">6</td>
            <td class="boarder-right">7</td>
            <td>8</td>
        </tr>
    </table>
    <table style="margin-top: 50px; width: 100%;">
        <tr>
            <td>
                orvosi pecsét
            </td>
            <td style="margin-left:40px; padding:0px 10px 0px 10px; border-top: dotted; text-align: center">
                orvosi aláírás
            </td>
        </tr>
    </table>
</div>""".format(nev=nev, anev=anev, szul_datum=szul_datum, lakhely=lakhely, taj=taj)

        parts.append(s)

        # if (osztaly != sheet.cell_value(i+1, 1) or osztaly == "1. a"):
        if i == LAST_INDEX - 1 or osztaly != sheet.cell_value(i + 1, 0):
            parts.append(also_allando)
            kesz = "".join(parts)
            print(kesz, osztaly.replace(". ", "") + ".pdf")
            if i != LAST_INDEX - 1:
                osztaly = sheet.cell_value(i + 1, 0)
                parts = []
                parts.append(felso_allando)


def print(kesz, pdf_name):
    option = {
        'page-size': 'A6',
        'margin-top': '0.5in',
        'margin-right': '0.5in',
        'margin-bottom': '0.5in',
        'margin-left': '0.5in',
        'encoding': 'UTF-8'
    }

    pdfkit.from_string(kesz, 'osztaly/' + pdf_name, options=option)


if __name__ == '__main__':
    pdf_template()
