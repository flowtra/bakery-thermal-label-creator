from mergepdf import mergePDFs
import time
import gspread
from gspread.cell import Cell
from google.oauth2 import service_account
import google.auth.transport.requests
import urllib.parse
import requests

def getOrders():
    headers = {"X-BUSINESS-API-KEY": "REDACTED",
               "Content-Type": "application/x-www-form-urlencoded"}
    querystring = {"perPage": "100",
                   "version":"2.0"}
    response = requests.get("https://api.hit-pay.com/v1/orders?page=1", headers=headers, params=querystring)
    if response.status_code == 200:
        return response.json()

def filterAliveAndDeliveryOnly(ordersJson): #FILTER for NON-CANCELLED and DELIVERY
    delOrders = []
    orders = ordersJson["data"]
    for order in orders:
        if order["status"] != 'canceled':
            if order["customer_pickup"] == False:
                delOrders.append(order)

    return delOrders


def filterAliveAndPickupOnly(ordersJson): #FILTER for NON-CANCELLED and PICK UP
    puOrders = []
    orders = ordersJson["data"]
    for order in orders:
        if order["status"] != 'canceled':
            if order["customer_pickup"] == True:
                puOrders.append(order)

    return puOrders

def cleanUpOrders(inArr):
    cleanArr = []
    for order in inArr:
        if order['channel'] != 'store_checkout':
            continue
        unitNo = None
        lineAddress = None
        if order['customer']['address']['street'] is not None:
            addStreet = order['customer']['address']['street']
            unitNo = '#'
            if '#' in addStreet: #Postcode validation
                unitNoLong = addStreet.split('#')[1]
                for char in str(unitNoLong):
                    if char.isdigit() or char == '-':
                        unitNo += char
                    if char == ' ':
                        break
            else:
                unitNo = ""

            lineAddress = ""
            city = str(order['customer']['address']['city']).upper()
            state = str(order['customer']['address']['state']).upper()
            lineAddress += order['customer']['address']['street'] #Add Street to Line Address
            sg = ['SG', 'SINGAPORE', 'SPORE', 'NIL']
            if city not in sg:
                 lineAddress += ' ' + city
            if state not in sg:
                lineAddress += ' ' + state
            lineAddress += ' ' + order['customer']['address']['postal_code']
            lineAddress = lineAddress.upper()

        obj = {
            "id": order["id"],
            "order_display_number": order["order_display_number"],
            "customer": order["customer"],
            "line_items": order["line_items"],
            "remark": order["remark"],
            "unitNo": unitNo,
            "lineAddress": lineAddress
        }
        cleanArr.append(obj)
    return cleanArr

def initDownloadToken():
    scope = ['https://www.googleapis.com/auth/drive']
    service_account_json_key = 'key.json'
    credentials = service_account.Credentials.from_service_account_file(
        filename=service_account_json_key,
        scopes=scope)

    credentials.refresh(google.auth.transport.requests.Request())
    accessToken = credentials.token

    return accessToken

def downloadPDF(sheetType, fileName, accessToken):
    while True:
        outputFilename = str(fileName) + '.pdf'  # Please set the output filename.
        spreadsheetId = "REDACTED"  # Please set your Spreadsheet ID.

        q = {
            'format': 'pdf',
            'size': '3.937x5.906',
            'portrait': 'true',
            'source': 'labnol',
            'sheetnames': 'false',
            'printtitle': 'false',
            'pagenumbers': 'false',  # or 'pagenum': 'UNDEFINED',
            'gridlines': 'false',
            'top_margin': '0.00',
            'bottom_margin': '0.00',
            'left_margin': '0.00',
            'right_margin': '0.00',
            'gid': sheetType #First sheet value is 0
        }
        queryParameters = urllib.parse.urlencode(q)
        url = f'https://docs.google.com/spreadsheets/d/{spreadsheetId}/export?{queryParameters}'
        headers = {'Authorization': 'Bearer ' + accessToken}
        res = requests.get(url, headers=headers)
        if 'html' in str(res.content):
            print(f"ERROR | Rate limited while trying to download order #{fileName}. Retrying in 10 Seconds. HTML Response:")
            print(res.content)
            time.sleep(10)
        else:
            break
    with open(f'output/{outputFilename}', 'wb') as f:
        f.write(res.content)
        return True

def initGSpread(sheet_name):
    Sheet_credential = gspread.service_account("key.json")
    spreadsheet = Sheet_credential.open_by_key('REDACTED')
    ws = spreadsheet.worksheet(sheet_name)

    return ws

def updateDelSheet(orderArr, ws):

    cells = []
    cells.append(Cell(row=12, col=4, value=orderArr['customer']['name']))
    cells.append(Cell(row=13, col=4, value=orderArr['customer']['phone_number']))
    cells.append(Cell(row=14, col=4, value=orderArr['lineAddress']))
    cells.append(Cell(row=8, col=4, value=orderArr['unitNo']))


    cells.append(Cell(row=19, col=4, value='#' + str(orderArr['order_display_number'])))
    if orderArr['remark'] != None:
        cells.append(Cell(row=21, col=2, value=str('Remark: ' + orderArr['remark'])))
    else:
        cells.append(Cell(row=21, col=2, value=str('Remark: N/A')))

    #QR Insertion
    cells.append(Cell(row=3, col=5, value='=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=500x500&data={}")'.format(orderArr["id"])))

    #Item Insertion
    itemRow = 26
    itemCount = 0
    totalQty = 0
    for item in orderArr['line_items']:
        if 'Delivery' in item["name"]:
            continue
        totalQty += int(item["quantity"])
        itemCount += 1
        cells.append(Cell(row=itemRow, col=2, value=itemCount))
        cells.append(Cell(row=itemRow, col=3, value=item["name"]))
        cells.append(Cell(row=itemRow, col=5, value=item["quantity"]))
        itemRow += 1
    while itemRow <= 34:
        cells.append(Cell(row=itemRow, col=2, value=""))
        cells.append(Cell(row=itemRow, col=3, value=""))
        cells.append(Cell(row=itemRow, col=5, value=""))
        itemRow += 1

    cells.append(Cell(row=19, col=2, value=totalQty))

    ws.update_cells(cells, value_input_option='USER_ENTERED')
    return True


def updatePickUpSheet(orderArr, ws):

    cells = []
    cells.append(Cell(row=12, col=4, value=str(orderArr['customer']['name']).upper()))
    cells.append(Cell(row=8, col=4, value=str(orderArr['customer']['name']).upper()))
    cells.append(Cell(row=13, col=4, value=orderArr['customer']['phone_number']))
    cells.append(Cell(row=14, col=4, value="N/A"))


    cells.append(Cell(row=19, col=4, value='#' + str(orderArr['order_display_number'])))
    cells.append(Cell(row=3, col=4, value='#' + str(orderArr['order_display_number'])))

    if orderArr['remark'] != None:
        cells.append(Cell(row=21, col=2, value=str('Remark: ' + orderArr['remark'])))
    else:
        cells.append(Cell(row=21, col=2, value=str('Remark: N/A')))

    #QR Insertion
    cells.append(Cell(row=3, col=5, value='=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=500x500&data=PICKUP%20{}")'.format(orderArr["id"])))

    #Item Insertion
    itemRow = 26
    itemCount = 0
    totalQty = 0
    for item in orderArr['line_items']:
        if 'Delivery' in item["name"]:
            continue
        totalQty += int(item["quantity"])
        itemCount += 1
        cells.append(Cell(row=itemRow, col=2, value=itemCount))
        cells.append(Cell(row=itemRow, col=3, value=item["name"]))
        cells.append(Cell(row=itemRow, col=5, value=item["quantity"]))
        itemRow += 1
    while itemRow <= 34:
        cells.append(Cell(row=itemRow, col=2, value=""))
        cells.append(Cell(row=itemRow, col=3, value=""))
        cells.append(Cell(row=itemRow, col=5, value=""))
        itemRow += 1

    cells.append(Cell(row=19, col=2, value=totalQty))

    ws.update_cells(cells, value_input_option='USER_ENTERED')
    return True


def getLastDownloadedOrderNo():
    with open('lastDownloaded.txt', 'r') as inFile:
        lastOrder = inFile.read().strip()

    return int(lastOrder)

def setLastDownloadedOrderNo(lastOrderNo):
    with open('lastDownloaded.txt', 'w') as outFile:
        outFile.write(lastOrderNo)

    return True

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    option = input("LABEL MAKER\n[M] Merge PDF\n[1] Delivery Orders\n[2] Pick Up Orders\nEnter option: ")
    downloadAccessToken = initDownloadToken()
    lastDownloadedOrderNo = getLastDownloadedOrderNo()
    print(f'Last Download Detected: Order #{str(lastDownloadedOrderNo)}. Will retrieve orders after.')
    firstOrderNo = True

    ######## TO FIX - LAST DOWNLOADED CAN CONFLICT BETWEEN PU / DELIVERY. MUST MAKE A SAVED NUMBER FOR EACH TYPE. BUT ITS 230AM SO IT CAN WAIT.
    if option == 'M':
        if mergePDFs():
            print('PDFs Merged. Process Complete')
    if option == '1':
        ws = initGSpread('DELIVERY')
        allOrders = getOrders()
        delOrders = filterAliveAndDeliveryOnly(allOrders)
        cleanOrders = cleanUpOrders(delOrders)
        for order in cleanOrders:
            orderDispNo = order["order_display_number"]
            if firstOrderNo:
                setLastDownloadedOrderNo(str(orderDispNo))
                firstOrderNo = False
            # if int(orderDispNo) != [INSERT SINGLE ORDER NO WANTED HERE]:
            #     continue
            if int(orderDispNo) <= lastDownloadedOrderNo:
                break
            updateDelSheet(order, ws)
            downloadPDF('0', orderDispNo, downloadAccessToken)
            print(f'PDF Created - Order #{orderDispNo}')
            time.sleep(5)
    elif option == '2':
        ws = initGSpread('PICKUP')
        allOrders = getOrders()
        puOrders = filterAliveAndPickupOnly(allOrders)
        puOrders.reverse()
        cleanOrders = cleanUpOrders(puOrders)
        for order in cleanOrders:
            orderDispNo = order["order_display_number"]
            # if str(orderDispNo) == '1172':
            #     pass
            # else:
            #     print(f'{orderDispNo} not it')
            #     continue
            if int(orderDispNo) <= lastDownloadedOrderNo:
                print(f'{orderDispNo} alr print')
                continue
            if firstOrderNo:
                setLastDownloadedOrderNo(str(orderDispNo))
                firstOrderNo = False

            updatePickUpSheet(order, ws)
            downloadPDF('REDACTED', orderDispNo, downloadAccessToken)
            print(f'PDF Created - Order #{orderDispNo}')
            time.sleep(2)
        if mergePDFs():
            print('PDFs Merged. Process Complete')

    # print_hi('PyCharm')

