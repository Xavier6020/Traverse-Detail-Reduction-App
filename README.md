# Traverse-Detail-Reduction-App
This code uses Qt5 to create a windowed application. The application is capable of reducing traverse and detail survey observations that come are formatted in an excel file. The code also allows the user to input  values that a typical detail survey would cover and reduce said observations from there. 
The code for the application is breakable and relise on the user input following correct units. The units given must be as follows:

installation requirments
Users must have installed Qt5 in their python enviroment. This can be done by opening the command prompt on your pc and typing the following command: pip install PyQt5 (This step must be done after installing python).

Bearings: dd.mmss (degrees minutes seconds)
Distances: m (meters)
Heights: m (meters)


Email tab requirments
A user can give a gmail account to give the exports a storage location. An app specific password must be created so that the email can be send. Google has rigerouse sercurity and requires that users opperating programes such as python provide a password to send emails. The reason the password is needed is due to a high number of hackers using python to hack Gmail accounts. You can generate an app specific password by following these steps: 
Go to the security tab in your google account: https://myaccount.google.com/security > Signing in to Google > App passwords > Select app > Other (Custom name) > “Python” >  Generate. Now Copy and paste the password into this app.


Traverse Survey requirments
The formate of the traverse survey must follow the face left (FL), face right (FR) observation method of data collection. 
The backsite (BS) must be smaller that the foresight (FS) and both must be circle. E.g. BS: 40° 31' 21", FS: 120° 41' 12"
An example of a formated traverse survey can be found in the main branch of this page.


Detail Survey requirments
Units must be followed


Station input tab requirments
Units must be followed


Detail Survey input tab requirments
Units must be followed

