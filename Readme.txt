
1. Printer setup:
	label dimensions 4" x 6"
	Port: C:\Output\output.PRN
	Set the parameters to use the printer settings.  
	Make sure it is online
	
	
2.  Spreadsheet setup:

	Several rows, 1 for each label format to be printed
	
	Active = Y, N or M for Yes No Manual (Manual means don't print the label, but process the manually generated ZPL)
	Printer: Name of printer to test
	dpi: dpi of printer
	IPAddress: IP Address of printer to use to generate .bmp image
	Application:  Application to print label formats with
	Label:  Name of label in label directory
	Language:  Language of label/Printer

	Active	Printer						dpi	IP Address  Application			Label								Language
	Y		ZDesigner ZD500-203dpi ZPL	200	10.80.22.72 ZebraDesigner Pro	TF_7_8_9_External Data Sources.lbl	ZPL
	.
	.
	.
	etc.


python RunTests.py /wbInputFileName=<Path to Spreadsheet>  /labelfiledirectory=<Path to directory holding label formats>  PrinterAddress=<ip.address of printer>

	Run parameters
		 wbInputFileName = ""  # Name of input spreadsheet

		labelfiledirectory = ""  # Directory name containing Label format


		
To compare output of two drivers, one will need to run once with one version of the driver,  
	Save the output directories before re-running -
	
								<Label Directory>\<Language eg. ZPL>\<Application name eg.ZebraDesigner Pro>\<Printername eg.ZDesigner ZD500-203dpi ZPL>\BMP
								D:\LabelWork\DriverTestData\ZPL\ZebraDesigner Pro\ZDesigner ZD500-203dpi ZPL\BMP
	
								<Label Directory>\<Language eg. ZPL>\<Application name eg.ZebraDesigner Pro>\<Printername eg.ZDesigner ZD500-203dpi ZPL>\Output
								D:\LabelWork\DriverTestData\ZPL\ZebraDesigner Pro\ZDesigner ZD500-203dpi ZPL\Output
								
								
								