#Importing Packages Used
from tkinter import *
from tkinter import ttk
from tkinter import messagebox ,simpledialog
from PIL import Image, ImageTk
import openpyxl as xl
import os
import requests

class Weather_App:
    def __init__(self, root):
        self.root = root
        self.root.geometry("700x450")
        self.root.maxsize(700,450)
        self.root.minsize(700,450)
        self.root["bg"]="#45A0EA"
        self.root.title("Weather APP --Rudra Kumar Mishra")

        # Creating Heading for the APP
        heading = Label(root, text = "Earth including over 2 million cities", fg = "red", bg = "sky blue", font = ("Copper Black", 24, "bold")).pack()


        # Creating Frame to Take input from user about Cities
        frame1 = Frame(root, bg = "#42c2f4", bd =5)
        frame1.place(x = 80, y = 50, width = 500, height = 60)


        # Creating a textbox Entry Field to get name of city
        txt_box = Entry(frame1, font = ("Lucida Handwritting", 25, "bold"), width = 17, border = 5, relief = SUNKEN)
        txt_box.grid(row = 0, column = 0, sticky = "W", padx = 10)


        # Creating Button to Submit the name of City for Output
        btn = Button(frame1, text = "Get Weather", fg = "green", font = ("Lucida Handwritting", 12, "bold"), command = lambda : self.get_weather(txt_box.get()))
        btn.grid(row = 0, column = 1, padx = 30)


        # Working on net frame to give output
        frame2 = Frame(root, bg = "#42c2f4", bd =5)
        frame2.place(x = 80, y = 150, width = 550, height = 250)


        #self.Result Label which will give Output
        self.result = Label(frame2, bg = "white", font = ("Algerian", 16, "italic"), justify = "left", anchor = "nw")
        self.result.place(relwidth = 1, relheight = 1)

    # Function to get Weather From OpenWeatherMap API
    def get_weather(self, city):
        api_key = "dd63ffdda0b2fd81b4f350346ca74bc9"
        url = "https://api.openweathermap.org/data/2.5/weather"
        param = {"APPID" : api_key, "q":city, "units": "metric"}
        rquestresult = requests.get(url, param)
        response = rquestresult.json()
        str = self.execption_handling(response)
        self.result["text"] = str

    #Function for removing any errors in the program
    def execption_handling(self, response):
        try:
            city = response["name"]
            country = response["sys"]["country"]
            weather = response["weather"][0]["description"]
            temp_min = response["main"]["temp_min"]
            temp_max = response["main"]["temp_max"]
            pressure = response["main"]["pressure"]
            humidity = response["main"]["humidity"]
            timezone = response["timezone"]
            str_out = f"City = {city}\nCountry = {country}\nWeather = {weather}\nMax Temparature = {temp_max}\nMin Temparature = {temp_min}\nPressure = {pressure}\nHumidity = {humidity}\nTimzone = {timezone}"
        except:
            str = "There was a problem in retrieveing that information"
        return str_out


root = Tk()
    
ob = Weather_App(root)

root.mainloop() 


