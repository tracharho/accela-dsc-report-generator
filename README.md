# accela-dsc-report-generator
This is an automation script used to login into Accela, download two report spreadsheets, apply a simple late/unlate process to the applicable columns, store the data, and email the reports to the appropriate personel.

The libraries used are pyautogui, seleniumn, openpyxl and stanrdard python libraries (shutil, os, time, etc.)
Needs to be refractored to so that the passwords and login aren't written in.
Uploaded files do not have any usernames, passwords, or emails for privacy
The selenium webbrowser opens on a 1366 x 768 resolution monitor.
This is an unstable build as moving the mouse during the process will result in errors and the mouse positions are statically based on the screen resolution.

#TO DO
Implement privacy importing for login credentials
Dynamically apply x,y for mouse coordinates based on monitor size
