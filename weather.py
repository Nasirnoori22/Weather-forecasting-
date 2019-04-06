import requests, json
import xlsxwriter
from datetime import datetime


api_key = "86232f552e2ba216fb26dc537ff85942"  # Key used for  api

base_url = "https://api.openweathermap.org/data/2.5/forecast?" #url of my this api

city_name = raw_input("Enter city name : ")   # Input From User
city_names = city_name.split(',')#margin  Multiple City
print ("Weather Dashboard of"), city_names
print("________________________________________________________________________________________________")

print(" Date                     current_tempreture        min tempereture           max tempereture ")

print("______________________________________________________________________________________")

counter = 0
workbook = xlsxwriter.Workbook('task8.xlsx') #create workbook for exl sheet
worksheet = workbook.add_worksheet("My sheet") # create my exl sheet


for c in city_names:
    print("Weather Dashboard of"), city_names
    print("\n")
    complete_url = base_url + "DE&appid=" + api_key + "&q= " + city_names[counter] #concadinate city_names with api url
    response = requests.get(complete_url)
    weather_data = response.json()#calling for Json Respons
    row_Headers = ['Date', 'Temperature in Celsius', 'High', 'Low'] # create Header for exl sheet
    worksheet.set_column(4, 5, 25)
    worksheet.set_column(3, 4, 25)
    row = 4
    col = 3
    worksheet.write_row(row, col, tuple(row_Headers))
    row += 1
    merge_format = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'fg_color': 'gray'})

    # Merge 3 cells.

    if weather_data["cod"] != "404":  # cheek for condition if request True or false

        for i in range(35): # put a variable in range of time
            current_temp = weather_data['list'][i]['main']['temp'] # access to json request
            temp_max = weather_data['list'][i]['main']['temp_min']
            temp_min = weather_data['list'][i]['main']['temp_max']
            date = weather_data['list'][i]['dt_txt']
            data_time_str = datetime.strptime(date, '%Y-%m-%d %H:%M:%S')
            d = data_time_str.strftime("%Y %B %A") # change time into string and date type
            a = data_time_str.strftime("%p")  # by using %p access to time format

            if i == 0 or i == 1 or i == 2 or i == 3: # create a condition for access to week of days
                print "today",data_time_str.hour,a

            elif i == 4 or i == 5 or i == 6 or i == 7 or i == 8 or i == 9 or i == 10 or i == 11:

                print "towmorrow",data_time_str.hour, a

            else:
                print(str(d))


            print("                          "

                      + str(current_temp)+"F" + "                   "
                      + str(temp_max) + "F" + "                  "
                      + str(temp_min)+"f")
            print("_______________________________________________________________________________________________")

            rowValues = [date,  # values to be writen into exl sheet
                         weather_data['list'][i]['main']['temp'],
                         weather_data['list'][i]['main']['temp_min'],
                         weather_data['list'][i]['main']['temp_max']]

            worksheet.write_row(row, col, tuple(rowValues))
            row += 1
            worksheet.merge_range('D2:G3', 'Weather Forecast Data', merge_format)

counter += 1

workbook.close()

# else:
#     print(" City Not Found ")
#     #     to find City
#


# End of counter