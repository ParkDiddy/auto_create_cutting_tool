#Requires AutoHotkey v2.0
esc::exitapp
Excel := ComObjActive("Excel.Application")
Workbook := Excel.Workbooks.Open("C:\Users\nparker\PycharmProjects\ScrapeDrillSize\DRILL CABINET 1.xlsx")
Worksheet := Workbook.Worksheets.Item("Jobber LENGTH DRILLS")
rownum := 3 ;starting row number of excel file
maxrownum := 119 ;maximum row number of excel file

while rownum <= maxrownum ;iterate on all rows from rownum to maxrownum
{
	drillsize := Round(Worksheet.Cells(rownum, 8).Value,4) ;set drill size to size data from excel
	oal := Round(Worksheet.Cells(rownum, 6).Value,3) ;set oal to oal data from excel
	fl := Round(Worksheet.Cells(rownum, 7).Value,3) ;set flute length to fl data from excel
	realdesc := Worksheet.Cells(rownum, 5).Value ;set description to description from excel
	WinActivate 'CreateDrills.vnc - GibbsCAM' ;activate gibbs create drills program
	MouseMove 24, 170 ;move to tool window 
	Click 2 ;click to bring up tool dialog box
	WinWait 'Milling Tool #1' ;wait until it shows up
	sleep 1000 ;sleep for a bit
	
	;set tool description
	ControlClick "WindowsForms10.EDIT.app.0.223fba4_r26_ad18"
	ControlSetText realdesc, "WindowsForms10.EDIT.app.0.223fba4_r26_ad18"

	;set to "use diameter"
	ControlClick "WindowsForms10.COMBOBOX.app.0.223fba4_r26_ad19"
	ControlSend "{home}", "WindowsForms10.COMBOBOX.app.0.223fba4_r26_ad19"
	ControlSend "{enter}", "WindowsForms10.COMBOBOX.app.0.223fba4_r26_ad19"
	
	;set size
	ControlClick "WindowsForms10.EDIT.app.0.223fba4_r26_ad13"
	ControlSetText drillsize, "WindowsForms10.EDIT.app.0.223fba4_r26_ad13"
	
	;set oal
	ControlClick "WindowsForms10.EDIT.app.0.223fba4_r26_ad12"
	ControlSetText oal, "WindowsForms10.EDIT.app.0.223fba4_r26_ad12"
	
	;set flute length and length out of holder
	ControlClick "WindowsForms10.EDIT.app.0.223fba4_r26_ad14"
	ControlSetText fl, "WindowsForms10.EDIT.app.0.223fba4_r26_ad14"
	sleep 500
	ControlClick "WindowsForms10.EDIT.app.0.223fba4_r26_ad112"
	ControlSetText fl, "WindowsForms10.EDIT.app.0.223fba4_r26_ad112"
	
	sleep 1000
	
	WinClose 'Milling Tool #1'
	
	;this block opens up the save selected tool dialog box and saves the tool
	MouseMove 24, 170
	Click "Right"
	Sleep 1000
	MouseMove 106, 97
	Click
	WinWait 'Save Tool File'
	ControlClick "Edit1"
	ControlSetText realdesc, "Edit1"
	sleep 100
	SendInput "{enter}"
	
	rownum++
}

MsgBox "Complete!"
ExitApp

;Workbook.Close()
;Excel.Quit()