Instructions for Using the PDF Script
	-It is recomended that you only have 1 instance of VBS running as the script
	is not reliable when multiple instances are open

	-The keyboard and mouse should not be used while the script is running with the 
	 only exceptions being when the script prompts you for information (will prompt
	 once directly after starting the program)

	-To kill the script at anytime, press windows key + z

1. Open up VBS, login and then minimize the screen

2. Double click on "Print Orders To PDF" to launch the script

3. The program will ask you to enter which plant you want to print orders from.
   Type the corresponding number and press enter. The program will then launch a new 
   instance of VBS and navigate to the orders page.

4. The script will paste the website page data into an excel sheet to copy the order 
   numbers from, and then begin switching back and forth printing orders and copying
   order numbers. When the script is finished, a message box saying "Complete" will 
   pop up.

Areas to Improve:
   -The script uses sleep commands to wait for two page reloads. There is code for waiting
    for a page to reload, however, it does not behave consistently across different computers.
   -By using an AutoHotkey script the computer cannot be used while the script is running. If
    given more time, I would pursue a python script that uses a library like Selenium webdriver.