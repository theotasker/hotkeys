#SingleInstance Force
CoordMode, Mouse, Client
CoordMode, Pixel, Client
msgbox done

Driver := ComObjCreate("Selenium.ChromeDriver")
Driver.SetCapability("debuggerAddress", "127.0.0.1:9222")
Driver.Start()




; Driver.Get("google.com")

msgbox done