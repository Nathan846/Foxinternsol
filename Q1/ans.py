import xlwings as xw
from geopy.geocoders import Nominatim
import requests
from tempconv import Fahrenheit,Cels
def main():
    geolocator = Nominatim(user_agent="waade")
    wb = xw.Book.caller()
    ws = wb.sheets(1)
    final_row = lastRow(1,wb)
    for i in range(2,final_row+1):
        if(ws.range((i,5)).value==0):
            pass
        else:
            ws.range((i,5)).value=1
            address = ws.range((i,1)).value
            loc = geolocator.geocode(address)
            lat = loc.latitude
            lon = loc.longitude
            apikey = "bfd59b9c746830b19814398cf0766475"
            url = "http://api.openweathermap.org/data/2.5/weather?lat={lat}&lon={lon}&appid={apikey}"
            jsonval = requests.get(url.format(lat=lat,lon=lon,apikey=apikey))
            data = jsonval.json()['main']
            if(ws.range(i,4).value=='F'):
                datas = Fahrenheit(data['temp'])
            else:
                ws.range(i,4).value= 'C'
                datas = Cels(data['temp'])
            ws.range((i,2)).value = datas
            ws.range((i,3)).value = data['humidity']

        
def lastRow(idx, workbook, col=1):
    ws = workbook.sheets(1)

    lwr_r_cell = ws.cells.last_cell      # lower right cell
    lwr_row = lwr_r_cell.row             # row of the lower right cell
    lwr_cell = ws.range((lwr_row, col))  # change to your specified column

    if lwr_cell.value is None:
        lwr_cell = lwr_cell.end('up')    # go up untill you hit a non-empty cell

    return lwr_cell.row
main()