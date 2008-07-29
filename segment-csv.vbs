'---------------------------------------- OUTPUT FILE


   Const ForReading = 1, ForWriting = 2, ForAppending = 8
   Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

   Dim fso, f, ts

   Set fso = CreateObject("Scripting.FileSystemObject")

   fso.CreateTextFile "output.csv"

   Set f_export = fso.GetFile("output.csv")
   Set ts_export = f_export.OpenAsTextStream(ForWriting, TristateUseDefault)



'----------------------------- INPUT FILE
Set objFSOName = CreateObject("Scripting.FileSystemObject")
Set objNameFile = objFSOName.OpenTextFile("input.chr", ForReading)


x=0

'write the first line
strNextLine = objNameFile.Readline
tmpMailCode = "Mail Code"
tmpWriteLine = tmpMailCode & "|" & strNextLine
tmpWriteLine2 = Replace(tmpWriteLine, "|", ",")
ts_export.write  tmpWriteLine2 & vbCRLF

Do Until objNameFile.AtEndOfStream
    strNextLine = objNameFile.Readline
    tmpArray = Split(strNextLine , "|") 


	tmpMailCode = "XXX"


	varRecordDate = Replace(tmpArray(13), chr(34),"")
	varRecordAmount  = Replace(tmpArray(14), chr(34),"")
'msgbox varRecordDate




'--- 25-36 Months
	tmpMyDate = "1/1/2005"
	If cDate(FormatDateTime(varRecordDate, vbShortDate)) >= cDate(FormatDateTime(tmpMyDate, vbShortDate)) Then
		tmpMailCode = "R84RVC7"
		If varRecordAmount < 1000 then 
			tmpMailCode = "R84RVC6" 
		End If
		If varRecordAmount < 500  then 
			tmpMailCode = "R84RVC5" 
		End If
		If varRecordAmount < 250  then 
			tmpMailCode = "R84RVC4" 
		End If
		If varRecordAmount < 100  then 
			tmpMailCode = "R84RVC3" 
		End If
		If varRecordAmount < 50   then 
			tmpMailCode = "R84RVC2" 
		End If
		If varRecordAmount < 25   then 
			tmpMailCode = "R84RVC1"
		End If
	End If

	'--- 13-24 Months
	tmpMyDate = "1/1/2006"
	If cDate(FormatDateTime(varRecordDate, vbShortDate)) >= cDate(FormatDateTime(tmpMyDate, vbShortDate)) Then
		tmpMailCode = "R84RVB7"
		If varRecordAmount < 1000 then 
			tmpMailCode = "R84RVB6" 
		End If
		If varRecordAmount < 500  then 
			tmpMailCode = "R84RVB5" 
		End If
		If varRecordAmount < 250  then 
			tmpMailCode = "R84RVB4" 
		End If
		If varRecordAmount < 100  then 
			tmpMailCode = "R84RVB3" 
		End If
		If varRecordAmount < 50   then 
			tmpMailCode = "R84RVB2" 
		End If
		If varRecordAmount < 25   then 
			tmpMailCode = "R84RVB1"
		End If
	End If

	'--- 0-12 Months
	tmpMyDate = "1/1/2007"
	If cDate(FormatDateTime(varRecordDate, vbShortDate)) >= cDate(FormatDateTime(tmpMyDate, vbShortDate)) Then
		tmpMailCode = "R84RVA7"
		If varRecordAmount < 1000 then 
			tmpMailCode = "R84RVA6" 
		End If
		If varRecordAmount < 500  then 
			tmpMailCode = "R84RVA5" 
		End If
		If varRecordAmount < 250  then 
			tmpMailCode = "R84RVA4" 
		End If
		If varRecordAmount < 100  then 
			tmpMailCode = "R84RVA3" 
		End If
		If varRecordAmount < 50   then 
			tmpMailCode = "R84RVA2" 
		End If
		If varRecordAmount < 25   then 
			tmpMailCode = "R84RVA1"
		End If
	End If














	tmpWriteLine = chr(34) & tmpMailCode & chr(34) & "|" & strNextLine
	tmpWriteLine2 = Replace(tmpWriteLine, chr(34) & "|" & chr(34), chr(34) & "," & chr(34))

	ts_export.write  tmpWriteLine2 & vbCRLF

    x = x + 1
Loop
objNameFile.Close
ts_export.Close

msgbox "Done"