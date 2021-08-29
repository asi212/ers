This folder contains documentation for the various projects I worked on

Programs that are used within ERS:
Alle/ERSTools/Endtest Document Generator/
Alle/ERSTools/EndtestData/
Alle/ERSTools/Reparatur Document Generator/



Projects that can be implemented later:
- "Label Printer" folder: contains example code and documentating of printer language
	- C# solution for printing found in C# solutions folder

- "Image Recognition" foler: contains image recognition script to read serial number labels
	- needs to use a faster OCR package (30 seconds to read label)?
		- might run fast enough to be reasonable on a newer computer...
	- Check extracted text again what the serial number in the Seriernnummern.xls spread sheet 
	  should say (Look at Doc Generator code for ideas on how to do this)

- "C# solutions" folder: contains examples of codes that can move data in and out of excel spreadsheets
	- has cool possibilities if you export at .exe and run it as a scheduled task
	- can syncronize information across different databases(spreadsheets) across ERS