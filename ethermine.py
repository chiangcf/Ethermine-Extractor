from bs4 import BeautifulSoup
import requests
import xlwt

style0 = xlwt.easyxf('font: name Calibri, color-index black, bold on,height 320')
style1 = xlwt.easyxf('font: name Calibri, color-index black,height 280')

# Creates workbook sheet
def createWs(wb, sheetName):
    ws = wb.add_sheet(sheetName)
    ws.write(0, 0, "ETHEREUM", style0)
    ws.write(0, 1, "Wallet Coins", style0)
    ws.write(0, 2, "Unpaid", style0)
    ws.write(0, 3, "Total", style0)
    ws.write(0, 4, "Ethereum Price (usd)", style0)
    ws.write(0, 5, "Total Value(usd)", style0)
    ws.write(3, 0, "Date", style0)
    ws.write(3, 1, "Duration(h)", style0)
    ws.write(3, 2, "Amount", style0)
    return ws

# Takes the address and get yours values into and excel file
def extractEth(address, filename):
    # Excel File
    wb = xlwt.Workbook()
    ws = createWs(wb, 'Ethereum')

    # Creating sessions
    s = requests.Session()
    s.headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.116 Safari/537.36'

    r  = s.get("https://ethermine.org/miners/"+ address +"/payouts")
    data = r.text
    soup = BeautifulSoup(data, "html.parser")
    containers = soup.findAll("table", {"class":"table"})
    payouts = containers[0].findAll("tr")

    row = 4
    wallet = 0
    for pays in payouts[1:]:
        currPay = pays.findAll("td")
        date = currPay[0].text[5:10]
        duration = currPay[3].text
        if duration != '':
            duration = float(duration)
        amount = float(currPay[4].text.partition("E")[0])
        wallet += amount

        ws.write(row, 0, date, style1)
        ws.write(row, 1, duration, style1)
        ws.write(row, 2, amount, style1)
        row+=1


    # Unpaid Balance from the pool
    r  = s.get("https://ethermine.org/miners/" + address)
    data = r.text
    soup = BeautifulSoup(data, "html.parser")
    containers = soup.findAll("div", {"class":"panel panel-info"})
    unpaids = containers[0].findAll("h4")
    unpaid = float(unpaids[0].text.partition("E")[0])

    # Ethereum value from coinmarketcap
    r  = s.get("http://coinmarketcap.com/currencies/ethereum/")
    data = r.text
    soup = BeautifulSoup(data, "html.parser")
    containers = soup.findAll("div", {"col-xs-6 col-sm-8 col-md-4 text-left"})
    price = float(containers[0].findAll("span")[0].text.partition("$")[2])
    totalCoins = wallet + unpaid
    totalValue = price * totalCoins

    # Adds all the values to the file
    ws.write(1, 1, wallet, style1)
    ws.write(1, 2, unpaid, style1)
    ws.write(1, 3, totalCoins, style1)
    ws.write(1, 4, price, style1)
    ws.write(1, 5, totalValue, style1)
    wb.save(filename + ".xls")



if __name__ == '__main__':
    print("Extracting info... Please wait...")
    # Address, Excel filename
    address = ""
    filename = "ethmine"
    extractEth(address,filename)
