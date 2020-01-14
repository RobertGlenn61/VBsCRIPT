' TODO 
' Do not change the value of input parameters
' Add Comments, headers, etc
' Create Functions only - return "error" if the function fails
' Prevent Write File writing during validation
' Create arrays using lists, especially for worklists where the number of elements
' is not always 96 elements

''Dim fxControl: Set fxControl = CreateObject("fxControlLib.FXControl")

'vbScript OpenTextFile constants. 
Const ForAppending = 8
Const ForReading = 1
Const ForWriting = 2

' =====================================================

Class Barcodes

' =====================================================
'/////////////////////////////////////////////
' Create default barcode objects for validation/simulation
Sub CreateDefaultBarcodeObject()
   DIM myBcObject, idx

   For idx=1 To 3
      Set myBcObject = CreateObject("Othros.EEOR") 
      myBcObject("Position") = "null"
      myBcObject("Description") = "null"
      myBcObject("Barcode") = "null"
      myBcObject("Name") = "null"
      myBcObject("TotalWells") = 0
      myBcObject("IsRack") = vbFalse
      myBcObject("WorklistHeader") = "null"
      HGSC.Barcodes.Source.Add(myBcObject)
      
      Select Case idx
         Case 1
            HGSC.Barcodes.Source.Add(myBcObject)
         Case 2
            HGSC.Barcodes.Destination.Add(myBcObject)
         Case 3
            HGSC.Barcodes.Reagent.Add(myBcObject)
      End Select
   Next

End Sub
'/////////////////////////////////////////////'/////////////////////////////////////////////
' Function CreateBarcodeObject
' Create a barcode object
' myBarcode: 1-d barcode if known
' mySdrType: s,d,r for source, destination, or reagent
' myPos: deck ALP position:  Scanner, P1..P20  do not use P0, it has special meaning
' myDescription: help further define object.  Use in the case where the labware is not on the deck
' myHeader: actual header used in the worklist where the 1-d barcode is found
' myBcObjId: numeric id into the s,d,or r list which locates the barcode object
'       list do not have any particular order, this is needed to easily id 
'       the object after returning from the routine
'/////////////////////////////////////////////
Function CreateBarcodeObject(myBarcode,mySdrType,myPos,myDescription, myHeader, ByRef myBcObjId)
   DIM myBcObject, lw, deckClass, rack, fileClass

   On Error Resume Next
   Set deckClass = New Deck
   Set fileClass = New TextFile
   Set myBcObject = CreateObject("Othros.EEOR") 

   myBcObject("Position") = myPos
   If myDescription<>"null" Then
      myBcObject("Description") = myDescription
   End If
   myBcObject("Barcode") = myBarcode
   myBcObject("Name") = "null"
   myBcObject("TotalWells") = 0
   myBcObject("IsRack") = vbFalse
   myBcObject("WorklistHeader") = myHeader

   If deckClass.IsDeckPosition(myPos) Then   
      Set lw = Labware.Deck.Positions(myPos).Labware
      myBcObject("Barcode") = lw.Properties.Barcode
      myBcObject("Name") = lw.Properties.Name
      myBcObject("TotalWells") = lw.Class.WellsX * lw.Class.WellsY
      If InStr(lw.Class.Type, "rack")>0 Then
         myBcObject("IsRack") = vbTrue
      Else
         myBcObject("IsRack") = vbFalse
      End If
   End If
   
   myBcObjId = -1
   Select Case mySdrType
      Case "s"
         HGSC.Barcodes.Source.Add(myBcObject)
         For rack=0 To HGSC.Barcodes.Source.Count-1
            If  HGSC.Barcodes.Source(rack).Position = myPos Then
               myBcObjId = rack
            End If
         Next
      Case "d"
         HGSC.Barcodes.Destination.Add(myBcObject)
         For rack=0 To HGSC.Barcodes.Destination.Count-1
            If  HGSC.Barcodes.Destination(rack).Position = myPos Then
               myBcObjId = rack
            End If
         Next
      Case "r"
         HGSC.Barcodes.Reagent.Add(myBcObject)
         For rack=0 To HGSC.Barcodes.Reagent.Count-1
            If  HGSC.Barcodes.Reagent(rack).Position = myPos Then
               myBcObjId = rack
            End If
         Next
      Case Else
         Call Err.Raise(60000, "vbScript Library", "s-d-r type must have the value 's','d,' or 'r'. Value entered = '" & mySdrType & "'")
   End Select

 If Err.Number<>0 Then
    CreateBarcodeObject = "Error in CreateBarcodeObject: " & Err.Description & " at " & Err.Source
    Err.Clear
    Call FileClass.LogData("", CreateBarcodeObject)
 Else
    CreateBarcodeObject = "success"
    Call FileClass.LogData("", "CreateBarcodeObject for barcode: " & myBcObj.Barcode)
 End If

End Function
'/////////////////////////////////////////////
' Function GetBarcodeObjectIndex(myBcObjectList, myPos)
' Given a list of barcode objects, find the index of the object
' with a given pos
' Returns integer 0...n if the objec is found
'/////////////////////////////////////////////
Function GetBarcodeObjectIndex(myBcObjectList,myPos)
   DIM ii
   
   GetBarcodeObjectIndex = -1
   For ii=0 To myBcObjectList.Count-1
      If myBcObjectList(ii).Position = myPos Then
         GetBarcodeObjectIndex = ii
         Exit For
      End If
   Next
  ' World.Globals.PauseGenerator.BtnPromptUser "GetBarcodeObjectId=" & GetBarcodeObjectIndex & " for " & myPos, Array("OK"), "OK"

End Function

'/////////////////////////////////////////////
' Function ExactBarcodeCompare
' Given two arrays, this subroutine will compare the values at each index. If the values do not match 
' they will be thrown into an errorList and reported at the end of the subroutine. If this list is reported
' the user will be prompted to abort.
'/////////////////////////////////////////////
Function ExactBarcodeCompare(scannerArray, worklistArray)
	Dim wellClass, errorList, resp, i, elementCount, totalWells
	
	Set wellClass = New Wells
 	errorList = ""
   elementCount = 0
   ExactBarcodeCompare = "success"
   totalWells = UBound(scannerArray)

	If IsArray(scannerArray) And IsArray(worklistArray) Then
		elementCount = UBound(worklistArray)
	    If UBound(scannerArray)<elementCount Then
	    	elementCount = UBound(scannerArray)
	   	End If

		For i = 1 To elementCount
			
			' DEBUG
			' resp = "Well: " & Wells.GetAlphaNumericID(i, totalWells) & "  WorklistBC: " & worklistArray(i) & vbCrlf
			' resp = "Well: " & Wells.GetAlphaNumericID(i, totalWells) & "  ScannerBC: " & scannerArray(i) & vbCrlf
			' resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")
			
			If worklistArray(i) <> scannerArray(i) Then
				errorList = errorList & "Well: " & WellClass.GetAlphaNumericID(i, totalWells) & " Barcode: " & worklistArray(i) & vbCrLf
			End If
		Next
	   If UBound(worklistArray)>elementCount Then
		   	For i=elementCount+1 To UBound(worklistArray) 
				  errorList = errorList & "Well: " & WellClass.GetAlphaNumericID(i, totalWells) & " Barcode: " & worklistArray(i) & vbCrLf
		   	Next
	   End If
	   If UBound(scannerArray)>elementCount Then
		   	For i=elementCount+1 To UBound(scannerArray) 
				  errorList = errorList & "Well: " & WellClass.GetAlphaNumericID(i, totalWells) & " Barcode: " & scannerArray(i) & vbCrLf
		   	Next
		End If

		If Len(errorList)>0 Then
			ExactBarcodeCompare = "Error from ExactBarcodeCompare:" & vbCrLf & "These barcodes were not found in the scanned rack: " & vbCrLf & errorList
		End If

	Else
		ExactBarcodeCompare = "Error from ExactBarcodeCompare:" & vbCrLf & "The input parameters for Sub ExactBarcodeCompare are not arrays."
	End If


End Function

'/////////////////////////////////////////////
' Function BarcodeCompare
' Given two arrays, this subroutine will compare the values at each index. If the values do not match 
' they will be thrown into an errorList and reported at the end of the subroutine. If there is a barcode
' in the scannerArray that is not present in the worklistArray it will compile a notificationList and present 
' it to the user. The user can choose to abort or continue at this point.
' Condsider the case where the array's have unequal elements
'/////////////////////////////////////////////
Function BarcodeCompare(scannerArray, worklistArray)
	Dim wellClass, errorList, notificationList, resp, ii, elementCount
	Dim totalWells
	
	Set wellClass = New Wells
	errorList = ""
	notificationList = ""
	elementCount = 0
	BarcodeCompare = "success"
	

   If Not IsArray(scannerArray) Or Not IsArray(worklistArray) Then
      BarcodeCompare = "Error From BarcodeCompare: " & vbCrLf & "The input parameters for Sub BarcodeCompare are not arrays."
      Exit Function
   End If
   
   totalWells = UBound(scannerArray)
	elementCount = UBound(worklistArray)
	If UBound(scannerArray)<elementCount Then
	   elementCount = UBound(scannerArray)
	End If
	For ii = 1 To elementCount

		' DEBUG
'   	 resp = "Well: " & wellClass.GetAlphaNumericID(i, totalWells) & "  WorklistBC: " & worklistArray(i) & vbCrlf
'		 resp = "Well: " & wellClass.GetAlphaNumericID(i, totalWells) & "  ScannerBC: " & scannerArray(i) & vbCrlf
'		 resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")
		
		' If the worklist has a valid tube, it must be in the scanner list
		If (worklistArray(ii) <> "0") Then
		   If (worklistArray(ii) <> scannerArray(ii))Then
		      errorList = errorList & "Well: " & wellClass.GetAlphaNumericID(ii, totalWells) & " scanned barcode " & scannerArray(ii) & " does not match worklist barcode " & worklistArray(ii) & vbCrLf
		   End If
		' If there's a scanned tube not in the worklist, notify operator
		ElseIf (scannerArray(ii) <> worklistArray(ii)) Then
		   notificationList = notificationList & "Well: " & wellClass.GetAlphaNumericID(ii, totalWells) & " scanned barcode " & scannerArray(ii) & " not found in worklist" & vbCrLf
		End If
				
	Next
	
	' Worklist barcodes not found in scanned barcodes
	If UBound(worklistArray)>elementCount Then
	   For ii=elementCount+1 To UBound(worklistArray) 
	     If worklistArray(ii)<> "0" Then
		     errorList = errorList & "Well: " & wellClass.GetAlphaNumericID(ii, 96) & " worklist barcode " & worklistArray(ii) & " not found in scanned barcodes." & vbCrLf
		  End If
	  	Next
	End If
	' Scanned barcodes not found in worklist
   If UBound(scannerArray)>elementCount Then
   	For ii=elementCount+1 To UBound(scannerArray) 
         If (scannerArray(ii)<> "0") Then
		      notificationList = notificationList & "Well: " & wellClass.GetAlphaNumericID(ii, 96) & " scanned barcode " & scannerArray(ii) & " not found in worklist." & vbCrLf
		   End If
   	Next
	End If

   ' Don't combine the two since an error ==> abort and a notice ==> OK
	If Len(errorList)>0 Then
	   BarcodeCompare = "Error From BarcodeCompare: " & vbCrLf & errorList
	ElseIf Len(notificationList)>0 Then
	   BarcodeCompare = "Notice (error?) From BarcodeCompare: " & vbCrLf & notificationList
	End If

End Function

' =====================================================
'////////////////////////////////////////
' Function GetBarcodeDataSet(barcode, searchStr)
' Given the barcode of piece of labware on the deck, 
' and a data set search string, 
' return the tube barcodes in its dataset.
'///////////////////////////////////////
Function GetBarcodeDataSet(barcode, searchStr)
   DIM alp, pos, key, alpBC, alpFound, nWells, myArray
   Dim ii, ds, resp, wellClass, name, lw, well
   Dim deckClass
  
  On Error Resume Next
  
  	Set wellClass = New Wells
  	Set deckClass = New Deck
  	Set fileClass = New TextFile
  	
  	alpFound = vbFalse
  	alp = deckClass.GetAlp(barcode)
  	If InStr(LCase(alp), "error") Then
  	   alpFound = vbFalse
      GetBarcodeDataSet = "Error in GetBarcodesDataSet: No labware found with "& _
      "barcode " & barcode
      Exit Function
   Else
      alpFound = vbTrue
   End If

   If alpFound Then
   
   resp = "GetBarcodeDataSet: Found " & barcode & " on alp " & alp & ". Searching for ds with " & searchStr
   Call fileClass.WriteToDebugFile("", resp, vbFalse)
   resp = ""
   
      Set lw = Positions(alp).Labware
      nWells = lw.Class.WellsX * lw.Class.WellsY
     
      redim myArray(nWells)
      for ii=0 To nWells
         myArray(ii) = "0"
      next

      ' Fetch barcode for each tube on rack
      For ds=0 To lw.DataSets.Count-1
         name = LCase(lw.DataSets.VariantDictionary.Keys(ds))
         nameList = nameList & name & vbCrLf
         If InStr(name, LCase(searchStr))>0 Then
            nameList = nameList & "***" & name & "***" & vbCrLf
            For ii=1 To nWells 
               well = WellClass.GetAlphaNumericId(ii, nWells)
               If InStr(name, LCASE(well)) > 0 Then
                  myArray(ii) = lw.DataSets.VariantDictionary.Values(ds)(ii)
               End If
            Next
         End If
      Next
      GetBarcodeDataSet = myArray
  End If

' DEBUG
   resp = "GetBarcodeDataSet:" & vbCrLf & "ALP Found = " & alp
   resp = resp & vbCrLf & "Barcode = " & lw.Properties.Barcode
   resp = resp & vbCrLf & "Labware = " & lw.Class.Name
   resp = resp & vbCrLf & "n wells = " & nWells
   resp = resp & vbCrLf & "Error " & Err.Description
   If alpFound Then
      resp = resp & vbCrLf & "NAMES: " & vbCrLf & nameList
      for ii=0 To nWells
         resp = resp & vbCrLf & ii & "==>" & myArray(ii)
      next
   Else
      resp = resp & myArray
   End If
 '''  resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")
   Err.Clear
   
End Function
' =====================================================


End Class  ' Class Barcodes

' =====================================================

Class Wells

' =====================================================

'/////////////////////////////////////////////
' FUNCTION GetAlphaNumericID
' Given the Biomek friendly row-wise well ID
' convert to the alpha-numeric well format
' A01 = 1, A02 = 2, B01 = 13 (for 96-well)
'/////////////////////////////////////////////

Function GetAlphaNumericID(id, totalWells)
   Dim pos, row, col, numCols
 
   If id<=0 OR id>totalWells Then
      ' Throw error
   End If

   If totalWells = 24 Then
      numCols = 6
   ElseIf totalWells = 96 Then
      numCols = 12
   ElseIf totalWells = 384 Then 
      numCols = 24
   Else
      numCols = 12
   End If

   pos = "000"
   row = (id-1)\numCols 
   pos = Chr( Asc("A") + row )
   col = id MOD numCols
   If col = 0 Then
      col = numCols
   End If
   If col<10 Then
      pos = pos & "0" & CStr(col)
   Else
      pos = pos & CStr(col)
   End If

 ' DEBUG
   resp = "GetAlphaNumericId for " & totalWells & " wells"
   resp = resp & vbCrLf & "row = " & row & "   col = "  & col 
   resp = resp & vbCrLf & "Well id: " & id & " = " & pos
'  resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")

   GetAlphaNumericID = pos

End Function

'/////////////////////////////////////////////
' FUNCTION CorrectWellFormat
' Put well in correct 3-char format
'/////////////////////////////////////////////

Function CorrectWellFormat(well, totalWells)
Dim myWell,alpha, number, lastLetter, lastColumn

 Select Case totalWells 
    Case 24
       lastLetter = ASC("D")
       lastColumn = 6
    Case 96
       lastLetter = ASC("H")
       lastColumn = 12
    Case 384
       lastLetter = ASC("P")
       lastColumn = 24
 End Select

 myWell = UCase(well)
 alpha = Mid(myWell, 1, 1)                              
 number = CInt(Replace(myWell, alpha, ""))
 
 If (ASC(alpha)<ASC("A")) OR (ASC(alpha)>lastLetter) Then
    myWell = "000"
 ElseIf (number<1) or (number>lastColumn) Then
    myWell = "000"
 Else
    myWell = alpha
    If number<10 Then
       myWell = myWell & "0" 
    End If
    myWell = myWell & CStr(number)
 End If

 CorrectWellFormat = myWell

' resp = "Parsing well: " & well & vbCrLf
' resp = resp & "Alpha: " & alpha & vbCrLf 
' resp = resp & "Number: " & number & vbCrLf
' resp = resp & "Final Format: " & myWell
' resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")

End Function

'/////////////////////////////////////////////
' FUNCTION GetNumericID
' Given the alpha-numeric well, convert to a
' Biomek friendly row-wise well ID
' A01 = 1, A02 = 2, B01 = 13 (for 96-well)
'/////////////////////////////////////////////

Function GetNumericID( myWell, totalWells )
 Dim row, col, pos, wellOK, numberOfColumns, lastLetter,lastColumn 
 DIM well, resp
 
 well = myWell                                                                  
 Select Case totalWells 
    Case 24
       lastLetter = ASC("D")
       lastColumn = 6
    Case 96
       lastLetter = ASC("H")
       lastColumn = 12
    Case 384
       lastLetter = ASC("P")
       lastColumn = 24
 End Select

   wellOK = vbFalse
   pos = -1
   well = UCase(well)      
 ' resp = World.Globals.PauseGenerator.BtnPromptUser(well, Array("OK"), "OK")

If Len(well)>=2 Then
   If Asc(Mid(well,1,1)) >= Asc("A") Then
      If Asc(Mid(well,1,1)) <= lastLetter Then
         If Cint(Mid(well,2,2)) >= 1 Then
            If Cint(Mid(well,2,2)) <= lastColumn Then
              wellOK = vbTrue              
            End If
         End If
      End If
   End If
End If
If wellOK = vbTrue Then
   row = CInt(Asc(Mid(well,1,1)) - Asc("A") )
   col = CInt(Mid(well,2,2)) 
   pos = row*lastColumn + col
End If      

' DEBUG
 resp = "GetNumericId for " & totalWells & " wells"
 resp = resp & vbCrLf & "WELL: " & well & "  POS: " & pos
'resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")

   GetNumericID = CInt(pos)
End Function

' =====================================================
'////////////////////////////////////////
' Function GetWellCount(barcode)
' Given the barcode of piece of labware on the deck, 
' return the number of wells
'///////////////////////////////////////
Function GetWellCount(barcode)
   DIM alp, nWells, resp, lw
   Dim deckClass

   On Error Resume Next
   Set deckClass = New Deck
   alp = deckClass.GetAlp(barcode)

   If InStr(LCase(alp), "error")<=0 And alp<>vbFalse Then
      Set lw = Positions(alp).Labware
      nWells = lw.Class.WellsX * lw.Class.WellsY
      GetWellCount = nWells
      resp = "ALP Found = " & alp & "   n wells = " & nWells 
   Else
      GetWellCount = alp
      resp = "Error in GetWellCount: " & CStr(alp)
  End If
  Call fileClass.WriteToDebugFile("null", Replace(resp,vbCrLf, vbTab))
  
  Err.Clear 
' DEBUG
' resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")

End Function
' =====================================================




' =====================================================

End Class     'Class Wells

' =====================================================

' =====================================================

Class TextFile

' =====================================================


'///////////////////////////////////////////////////////
'Sub CreateSampleIntakeWorklist
'Takes the exemplar worklist and inserts alternating water source wells into the WaterWellPosition column. 
'
'NOTE:
'Should this be done for each TFF worklist? Or should we create a subroutine that does it all?
'///////////////////////////////////////////////////////

Sub CreateSampleIntakeWorklist(wl)
Dim wlArray, header, headerArray, file
Dim ii, waterIdx, line, newLine, wlData, well

wlArray = Split(HGSC.Lims.Gen2.Data.Worklist, vbCrLf)

header = wlArray(0)
headerArray = Split(header, ",")

	For ii = 0 to Ubound(headerArray)
	   If headerArray(ii) = "WaterWellPosition" Then
   		waterIdx = ii
         Exit For
	   End If
	Next

	For ii = 1 to UBound(wlArray)

		line = Split(wlArray(ii), ",")

		If Not Len(wlArray(ii)) < UBound(wlArray) Then

			Select Case ii Mod 4

				Case 0 
   				well = "D1"

				Case 1
   				well = "A1"

				Case 2
   				well = "B1"

				Case 3 
				well = "C1"

			End Select

 
			line(waterIdx) = well
			newLine = Join(line, ",")
			wlData = wlData & newLine & vbCrLf

		End If

	Next      

	data = header & vbCrLf & data

   Call Write(data,HGSC.Lims.Gen2.Data.Worklist,vbFalse, vbTrue)

End Sub

'////////////////////////////////////////
' SUB LogData
' Takes a message as a string and writes it to the given log file. This method appends a date 
' and a time stamp to the beginning of each string sent to the log file. 
' 
' Do not log data unless a method is running
'
' This subroutine assumes the file path is unique.
'///////////////////////////////////////
Sub LogData(logFileName, data)
Dim fso, logFile, fileName

  On Error Resume Next
  ' Stop forward processing
  World.Globals.PauseGenerator.StallUntilETSIs 0 

   Set fso = CreateObject("Scripting.FileSystemObject")

   If Len(LogFileName)<=0 Then
      fileName = "C:\HGSC\Log\Log.txt"
   Else
      fileName = logFileName
   End If
   
   If NOT (World.Globals.BrowserMan.Simulating) Then
      If fso.FileExists(fileName) Then
      	Set logFile = fso.OpenTextFile(fileName, ForAppending)
      Else
      	Set logFile = fso.CreateTextFile(fileName)
      End If
      logFile.Write(Date & " " & Time & ": " & data & vbCrLf)
      logFile.Close
   End If

   Set fso = Nothing
   
  If Err.Number<>0 Then
     Dim msg: msg = "Error in LogData: " & vbCrLf & Err.Description
     Call World.Globals.PauseGenerator.BtnPromptUser(msg, Array("OK"), "OK")
     Err.Clear
  End If

   
   Call WriteToDebugFile("null", data, vbFalse)
   
End Sub
'///////////////////////////////////////

'///////////////////////////////////////
' Function WriteToDebugFile
' Used for debugging, writes data to file
' regardless of simulation mode
' intended to capture values of variables in scripts
'///////////////////////////////////////
Sub WriteToDebugFile(fileName,data,displayPopUp)
Dim fso, logFile

' Stop forward processing
World.Globals.PauseGenerator.StallUntilETSIs 0 

Set fso = CreateObject("Scripting.FileSystemObject")

   ' Just fall through if problems
   On Error Resume Next 
   If Not fso.FileExists(fileName) Then
      fileName = "C:\HGSC\Log\DebugFile.txt"
   End If
   If fso.FileExists(fileName) Then
   	Set logFile = fso.OpenTextFile(fileName, ForAppending)
   Else
   	Set logFile = fso.CreateTextFile(fileName)
   End If
   logFile.Write(Date & " " & Time & ": " & data & vbCrLf)
   logFile.Close
   Set fso = Nothing
   
   If displayPopUp = vbTrue Then
      Call World.Globals.PauseGenerator.BtnPromptUser(data, Array("OK"), "OK")
   End If
   If Err.Number<>0 Then
      Dim msg: msg = "Error in WriteToDebugFile: " & vbCrLf & Err.Description
      Call World.Globals.PauseGenerator.BtnPromptUser(msg, Array("OK"), "OK")
      Err.Clear
   End If

End Sub
'////////////////////////////////////////////////////

'////////////////////////////////////////////////////
' FUNCTION GetWellDataList
' Given a file containing a column of alpha numeric
' wells in column wellHeader, return the data in 
' the associated cell in the dataHeader column
' Headers can be either an integer or a string
'
' Data returned is a string of the form A01,bc1, ..., H12,bc96 
' 
' TODO:
' What happens if 
' ...header is not in file?
' ...file does not exist?
' Pass in the total wells?
'////////////////////////////////////////////////////
Function GetWellDataList(fileName, wellHeader, dataHeader, lwBarcode, lwHeader, defaultTotalWells)
Dim dataArray, myList, wellId, ii, wellClass
Dim resp

  On Error Resume Next

   resp = "GetWellDataList: " & vbCrLf
   resp = resp & "File: " & fileName & vbCrLf 
   resp = resp & "Well Header: " & wellHeader & vbCrLf
   resp = resp & "Data Header: " & dataHeader & vbCrLf
   resp = resp & "Barcode: " & lwBarcode
   Call WriteToDebugFile("", resp, vbFalse)
'   Call World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")

   Set wellClass = new Wells
   dataArray = GetWellDataArray(fileName, wellHeader, dataHeader, lwBarcode, lwHeader, defaultTotalWells)
   If Not IsArray(dataArray) Then
      GetWellDataList = "Error in GetWellDataList: " & dataArray
      Call WriteToDebugFile("", GetWellDataList, vbFalse)
'      Call World.Globals.PauseGenerator.BtnPromptUser(dataArray, Array("OK"), "OK")
      Exit Function
   End If
 
 ' Assemble the list
   myList = ""
   For ii=1 To UBound(dataArray)
      If dataArray(ii)<>"0" Then
         wellID = wellClass.GetAlphaNumericID(ii, UBound(dataArray))
         If Len(myList)<=0 Then
           myList = CStr(wellId) & "," & dataArray(ii)
         Else
            myList = myList & "," & CStr(wellId) & "," & dataArray(ii)
         End If
      End If
   Next

   If Err.Number<>0 Then
      GetWellDataList = "GetWellDataList error, file: " & fileName & vbCrLf & "Error= " & err.Description
'      Call World.Globals.PauseGenerator.BtnPromptUser(GetWellDataList, Array("OK"), "OK")
      Err.Clear
      Exit Function
   Else
      GetWellDataList = myList
   End If
   
   Call WriteToDebugFile("", GetWellDataList, vbFalse)
'   Call World.Globals.PauseGenerator.BtnPromptUser("GetWellDataList:" & vbCrLf & GetWellDataList, Array("OK"), "OK")
End Function

'////////////////////////////////////////////////////
' FUNCTION GetWellDataArray
' Given a file containing a column of alpha numeric
' wells in column wellHeader, return the data in 
' the associated cell in the dataHeader column
' Headers can be either an integer or a string
'
' Unless there's an error, GetWellDataArray returns an array 1..nWells
' If there's no data, the value is "0"
' 
' TODO:
' July 11, 2017 - Refactor, this can be cleaner
' What happens if 
' ...header is not in file?
' ...file does not exist?
' Pass in the total wells?
' 
' In Get Pattern, we just want to read a file w/ header 'well,'active'
' If no barcode, how do we know the size of array?
'   make assumption based on file size?  ==> max well idx is the size, 
'      meaning the file 1..24, 1..96, 1..384
'   make barcode=arraySize?  
'   if null, popup query asking user for plate size?
'////////////////////////////////////////////////////
Function GetWellDataArray(fileName, wellHeader, dataHeader, lwBarcode, lwHeader,defaultTotalWells)
Dim fso, lineArray, line, headerArray, file, dataArray, totalWells, data
Dim wellIdx, dataIdx, lwIdx, header, ii, wellClass, resp, numericID
Dim isLwMatch, numElements

   On Error Resume Next
   
   resp = "GetWellDataArray: " & vbCrLf
   resp = resp & "File: " & fileName & vbCrLf 
   resp = resp & "Well Header: " & wellHeader & vbCrLf
   resp = resp & "Data Header: " & dataHeader & vbCrLf
   resp = resp & "Labware Barcode (optional): " & lwBarcode & vbCrLf
   resp = resp & "Labware Header (optional): " & lwHeader
'   Call World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")
   Call WriteToDebugFile("", resp, vbFalse)
   
   Set wellClass = new Wells
   Set fso = CreateObject("Scripting.FileSystemObject")
   If NOT fso.FileExists(fileName) Then
      GetWellDataArray = "Error in GetWellDataArray, file does not exist: " & fileName
      Set fso = Nothing
      Call World.Globals.PauseGenerator.BtnPromptUser(GetWellDataArray, Array("OK"), "OK")
      Exit Function
   End If

   Set File = fso.OpenTextFile(fileName, ForReading)
   If Err.Number>0 Then
      GetWellDataArray = "Error in GetWellDataArray, can't open file: " & fileName & vbCrLf & "Error= " & err.Description
      Err.Clear
      Set fso = Nothing
'      Call World.Globals.PauseGenerator.BtnPromptUser(GetWellDataArray, Array("OK"), "OK")
      Call WriteToDebugFile("", GetWellDataArray, vbFalse)
      Exit Function
   End If
   
   If IsNull(defaultTotalWells) Then
      totalWells = 96
   Else
      totalWells = defaultTotalWells
   End If
   If lwBarcode<>"null" Then
      totalWells = wellClass.GetWellCount(lwBarcode)
   Else
      ' TODO: read file to find determine the totalWells
      ' ask operator?
   End If
   If Not IsNumeric(totalWells) Then
      GetWellDataArray = "Error in GetWellDataArray, can't find totalWells in barcode " & lwBarcode & vbCrLf & "Error= " & totalWells
'      Call World.Globals.PauseGenerator.BtnPromptUser(GetWellDataArray, Array("OK"), "OK")
      Call WriteToDebugFile("", GetWellDataArray, vbFalse)
      Set fso = Nothing
      Exit Function
   End If

   ReDim dataArray(CInt(totalWells))
   For ii=0 To totalWells
      dataArray(ii) = "0"
   Next
   If Err.Number>0 Then
      GetWellDataArray = "Error in GetWellDataArray, 'totalWells' " & CStr(totalWells) & vbCrLf & "Error= " & err.Description
      Call WriteToDebugFile("", GetWellDataArray, vbFalse)
      Err.Clear
      Set fso = Nothing
      Exit Function
   End If

' Any problems reading the header line?
   line = File.ReadLine
   If Err.Number>0 Then
      GetWellDataArray = "Error in GetWellDataArray, 'line = File.ReadLine' file: " & fileName & vbCrLf & "Error= " & err.Description
      Err.Clear
      Set fso = Nothing
      Call WriteToDebugFile("", GetWellDataArray, vbFalse)
'      Call World.Globals.PauseGenerator.BtnPromptUser(GetWellDataArray, Array("OK"), "OK")
      Exit Function
   End If
   If Len(line)<3 Then
      GetWellDataArray = "Error in GetWellDataArray, file does not contain header data: " & fileName
      Call WriteToDebugFile("", GetWellDataArray, vbFalse)
'      Call World.Globals.PauseGenerator.BtnPromptUser(GetWellDataArray, Array("OK"), "OK")
      Set fso = Nothing
      Exit Function
   End If

' Find the header indices
   wellIdx = GetHeaderIndex(line,wellHeader)
   dataIdx = GetHeaderIndex(line,dataHeader)
   lwIdx = GetHeaderIndex(line,lwHeader)
   numElements = UBound(Split(line, ","))
  
  ' Indices not found
  If wellIdx<0 Or dataIdx<0 Then
      resp = "Error in GetWellDataArray, could not find headers in file '" & fileName & "' "
      resp = resp & "while searching for well header '" & wellHeader & "' and data header '" & dataHeader & "'"     
      GetWellDataArray = resp
      Call WriteToDebugFile("", GetWellDataArray, vbFalse)
'      Call World.Globals.PauseGenerator.BtnPromptUser(GetWellDataArray, Array("OK"), "OK")
      Set fso = Nothing
      Exit Function
  End If
     
Do Until File.AtEndOfStream
   line = File.ReadLine     
   lineArray = Split(line, ",")
   
   ' DEBUG
'    resp = "LINE: " & line & vbCrLf & "Array Size: " & UBound(lineArray)
'    For ii=0 To UBound(lineArray)
'       resp = resp & vbCrLf & "Element " & CStr(ii) & " = " & lineArray(ii)
'    Next
'    Call WriteToDebugFile("", resp, vbFalse)
'    Call World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")

   ' Check optional match for labware barcode
   ' Option selected id lwIdx>=0
   isLwMatch = vbFalse
   If lwIdx>=0 Then
      ' Make sure the lwIdx element is in the array
      If UBound(lineArray)<lwIdx Then
         resp = "GetWellDataArray, error: " & vbCrLf
         resp = resp & "The line from the worklist does not seem to match the expected rack barcode header" & vbCrLf
         resp = resp & "Line from file: " & line & vbCrLf & "Labware barcode header index = " & lwIdx
         Call WriteToDebugFile("", GetWellDataArray, vbFalse)
'         Call World.Globals.PauseGenerator.BtnPromptUser(GetWellDataArray, Array("OK"), "OK")
         Set fso = Nothing
         Exit Function
      Else
         isLwMatch = vbFalse
         resp = "isLwMatch?" & vbCrLf & lineArray(lwIdx) & " ?=? " & lwBarcode
'         Call World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")
         If lineArray(lwIdx)=lwBarcode Then
            isLwMatch = vbTrue
         End If
     End If
      
   Else
      isLwMatch = vbTrue
   End If
    
   If isLwMatch = vbTrue Then

      ' Well index may not be alpha-numeric
      If IsNumeric(lineArray(wellIdx)) Then
         numericID = CInt(lineArray(wellIdx))
      Else
         numericID = wellClass.GetNumericID(lineArray(wellIdx), totalWells)
      End If
      If numericID>0 Then
         data = Trim(lineArray(dataIdx))
         If Len(data)>0 Then
           dataArray(numericID) = data
         End If
         resp = "GetWellDataArray, numeric ID " & numericID & " > 0" & vbCrLf & line & vbCrLf
         resp = resp & "WELL: " & lineArray(wellIdx) & "  POS: " & numericID & "  Data " & dataArray(numericID)
'         Call WriteToDebugFile("", resp, vbFalse)
'         Call World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")
      Else
         GetWellDataArray = "GetWellDataArray error, numeric ID " & numericID & " <= 0" & vbCrLf
         GetWellDataArray = GetWellDataArray & "WELL: " & lineArray(wellIdx) & "  does not translate to numeric id" 
         Call WriteToDebugFile("", GetWellDataArray, vbFalse)
'         Call World.Globals.PauseGenerator.BtnPromptUser(GetWellDataArray, Array("OK"), "OK")
         Set fso = Nothing
         Exit Function
      End If
      Call WriteToDebugFile("", resp, vbFalse)
   End If  ' if labware match

   ' DEBUG
   ' This fails if numericID<0
     resp = "GetWellDataArray" & vbCrLf
     resp = resp & "WELL: " & lineArray(wellIdx) & "  POS: " & numericID & "  Data " & dataArray(numericID)
'     Call  World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")
     
Loop
Set fso = nothing	
GetWellDataArray = dataArray
If Err.Number>0 Then
   GetWellDataArray = "GetWellDataArray error, 'GetWellDataArray = dataArray' file: " & vbCrLf & "Error= " & err.Description
   Call WriteToDebugFile("null", GetWellDataArray, vbFalse)
   Err.Clear
Else
   resp = "GetWellDataArray UBound=" &  UBound(dataArray) & vbTab & "VALUES:" & vbCrLf
   For ii=1 To UBound(dataArray)
      resp = resp & "[" & CStr(ii) & "]=" & dataArray(ii) & vbCrLf
   Next
   Call WriteToDebugFile("null", Replace(resp, cbCrLf, vbTab), vbFalse)
'   Call World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")
End If

End Function

'//////////////////////////////
' FUNCTION GetHeaderIndex
' Return the index of the header in the comma-separated string
' If no match is found, -1 is returned
' Ignores case
'/////////////////////////////
Function GetHeaderIndex(headerLine,myHeader)
   Dim headerArray, ii, header
   GetHeaderIndex = -1
   headerArray = Split(headerLine, ",")
   ii=0
   If LCase(myHeader)="null" Then
      GetHeaderIndex = -1
   ElseIf isNumeric(myHeader) Then
      GetHeaderIndex = CInt(myHeader)
   Else
      For Each header in headerArray 
         If LCase(myHeader) = LCase(header) Then
            GetHeaderIndex = ii
            Exit For
         End If
         ii = ii + 1
      Next
   End If
End Function

	
'//////////////////////////////
'FUNCTION IsFileOnServer
' Determine if the input file is on the network
' If it is not return vbFalse
' if it is, replace the mapped network drive with the 
' UNC path
'/////////////////////////////
Function IsFileOnServer(fileName)
Dim f, resp
Dim objWMIService, colDrives, objDrive
	
	IsFileOnServer = vbFalse
   Set objWMIService = GetObject("winmgmts:\\")
   Set colDrives = objWMIService.ExecQuery ("Select * From Win32_LogicalDisk Where DriveType = 4 ")
   f=LCase(fileName)
   
   For Each objDrive in colDrives
   ' DEBUG
    resp = "Drive = " & CStr(objDrive.DeviceID) & vbCrLf
    resp = resp & "Provider Name = " & CStr(objDrive.ProviderName) & vbCrLf
    resp = resp & "Device ID = " & CStr(objDrive.DeviceID) & vbCrLf
    resp = resp & "File Name = " & fileName & vbCrLf
'    resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")

    If InStr(f,LCase(CStr(objDrive.DeviceID)))>0 OR InStr(f,LCase(CStr(objDrive.ProviderName)))>0 Then
 	    f =  Replace(f, LCase(CStr(objDrive.DeviceID)), LCase(CStr(objDrive.ProviderName)))
       resp = "Incoming file name: " & fileName & vbCrLf
       resp = resp & "New File Name: " & f 
'       resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")
       IsFileOnServer = vbTrue
       Exit For
    End If   
   Next
   	
End Function

'//////////////////////////////
' FUNCTION ResolveFileName
' Returns the UNC File Name
'/////////////////////////////

Function ResolveFileName(fileName)
Dim foundFileOnNetwork, f, resp
Dim colDrives, objDrive

	foundFileOnNetwork = vbFalse 
   Set objWMIService = GetObject("winmgmts:\\")
   Set colDrives = objWMIService.ExecQuery ("Select * From Win32_LogicalDisk Where DriveType = 4 ")
   		
   ResolveFileName = fileName
   f = LCase(fileName)
   For Each objDrive in colDrives
	  ' DEBUG
  	   resp = "Drive = " & CStr(objDrive.DeviceID) & vbCrLf
  	   resp = resp & "Provider Name = " & CStr(objDrive.ProviderName) & vbCrLf
  	   resp = resp & "File Name = " & fileName & vbCrLf
'  	resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")

   	If(InStr(f,LCase(objDrive.DeviceID)) OR InStr(f,LCase(objDrive.ProviderName))) Then
	      foundFileOnNetwork = vbTrue    	
         ResolveFileName =  Replace(f, LCase(objDrive.DeviceID), LCase(objDrive.ProviderName))	
         Exit For
   	End If   
    Next
    ' DEBUG
'    If ( foundFileOnNetwork = vbFalse ) Then
'       resp = "Error: For LIMS to work, the file must be located on a server not the local PC." & vbCrLf
'       resp = resp & "File selected: " & fileName
'       resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")
'    Else
'       resp = "Result of resolveFileName" & vbCrLf
'       resp = resp & "File selected: " & fileName & vbCrLf
'       resp = resp & "Resolved FileName: " & ResolveFileName
'       resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")
'    End If

End Function

'////////////////////////////////////////
' SUB Write
' Takes a message as a string and writes it to the given fileName path.
' okToOverwriteFile and okToAppendFile are both booleans that denote if a file is allowed to be overwrriten or appended to.
'
' TODO:
' Prompt user for a file name if a new file is created
'///////////////////////////////////////

Sub Write(data, fileName, okToAppendFile, okToOverwriteFile)
  Dim append, overwrite, file, fso, resp, filePath

  Set fso = CreateObject("Scripting.FileSystemObject")
  append = okToAppendFile
  overwrite = okToOverwriteFile

  If fso.FileExists(fileName) AND okToOverwriteFile=vbFalse AND okToAppendFile=vbFalse Then
     resp = "The selected file already exists: " & vbCrLf
     resp = resp & fileName & vbCrLf & vbCrLf & "What should be done?"
     resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("Append", "Overwrite", "New File"), "Append")
     If resp="Append" then
        append = vbTrue
        overwrite = vbFalse
     End If
     If resp="Overwrite" then
        overwrite = vbTrue
        append = vbFalse
     End If
  End If

  If NOT (World.Globals.BrowserMan.Simulating) Then
     If resp = "New File" Then
        filePath = fso.GetBaseName(fileName)
        Set file = fso.CreateTextFile(filePath & "\_" & Now)
     ElseIf fso.FileExists(fileName) AND append = vbTrue Then
        Set file = fso.OpenTextFile(fileName, ForAppending)
     ElseIf fso.FileExists(fileName) AND overwrite = vbTrue Then
        Set file = fso.CreateTextFile(fileName, ForWriting)
     ElseIf NOT fso.FileExists(fileName) Then
        Set file = fso.CreateTextFile(fileName)
     End If
     file.Write(data)
     file.Close
  End If 
  Set fso = Nothing

End Sub

'////////////////////////////////////////////
' FUNCTION Read
' Takes a file as an input and returns the entire file as a string.
'////////////////////////////////////////////

Function Read(fileName)
  Dim fso, TextFile, line
  Set fso = CreateObject("Scripting.FileSystemObject")

  On Error Resume Next
  Set TextFile = fso.OpenTextFile(fileName, ForReading)
  If Err.Number=0 Then
     Do Until TextFile.AtEndOfStream 
	     line = line & TextFile.ReadLine &  vbCrLf
	     If Err.Number<>0 Then
	        Exit Do
	     End If
	   ''	World.Globals.PauseGenerator.BtnPromptUser line, Array("OK"), "OK"
	  Loop
	  If Err.Number=0 Then
	     Read = line
	  End If
  End If

  If Err.Number<>0 Then
     Read = "Error Reading File " & fileName & vbCrLf & Err.Description
     Err.Clear
  End If
  
End Function

'/////////////////////////////////////////////
' FUNCTION CreateWorklist
' Create a new file, FileName_LIMS and copy the text to it
'/////////////////////////////////////////////

Function CreateWorklist(fileName, text, suffix)
Dim tso, fso, worklistName,ii
  
  ' Find the extention
  For ii=Len(fileName) to Len(fileName)-10 Step -1
     If Mid(fileName, ii, 1) = "." Then
        Exit For
     End If
  Next

  worklistName = Mid(fileName,1,ii-1) & Replace(filename, ".", suffix & ".", ii, 1) 

' DEBUG
'  resp = "Original File: " & fileName & vbCrLf
'  resp = resp & "LIMS File: " & worklistName & vbCrLf
'  resp = resp & "Index of extension is " & CStr(ii)
'  resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")

   Set fso = CreateObject("Scripting.FileSystemObject")

   If fso.FileExists(worklistName) Then
      set tso = fso.GetFile(worklistName)
      tso.Delete
      Set tso = Nothing
   End If
  
' DEBUG
'  resp = World.Globals.PauseGenerator.BtnPromptUser("Writing File: " & worklistName, Array("OK"), "OK")
   
   Set tso = fso.CreateTextFile(worklistName, true, false)      
   tso.Write(text)

   tso.Close
   Set tso = Nothing
   Set fso = Nothing
   
   CreateWorklist = worklistName

End Function


'/////////////////////////////////////////////
' FUNCTION IsHeaderCorrect
' Compare the file header to the expected header
'/////////////////////////////////////////////
Function IsHeaderCorrect(fileHeader, expectedHeader)
DIM expectedNumberOfColumns, expectedColumnHeaders, fileHeaders
DIM errMsg, lastHeader, ii, errTitle
Dim expectedHeaderText

' expectedHeader = "S_Rack,S_Well,D_Rack,D_Well,S_Volume,TE_Rack,TE_Well,TE_Volume"
' expectedHeader = "S_Rack,S_Well,S_Barcode,S_Volume,TE_Rack,TE_Well,TE_Volume,Metaproject,D_Rack,D_Well,D_Barcode,D_Label"

 errTitle = "Error in IsHeaderCorrect:" & vbCrLf & "The worklist file headers do not match the expected headers:" & vbCrLf
 expectedColumnHeaders = Split(expectedHeader,",")
 expectedNumberOfColumns = UBound(expectedColumnHeaders)
 
' If header is not a match throw an error

     errMsg = ""
     fileHeaders = Split(fileHeader,",")
     If UBound(expectedColumnHeaders) <= UBound(fileHeaders) Then
        lastHeader = UBound(expectedColumnHeaders)
     Else
        lastHeader = UBound(fileHeaders)
     End If
     If UBound(expectedColumnHeaders) <> UBound(fileHeaders) Then
        errMsg = errMsg & "Expected Header: " & expectedHeaderText & vbCrLf
        errMsg = errMsg & "File's Header: " & fileHeader & vbCrLf
     End If
     For ii=0 To lastHeader    
        If Trim(LCase(fileHeaders(ii))) <> Trim(LCase(expectedColumnHeaders(ii))) Then
           errMsg = errMsg & "Column " & CStr(ii+1) & " header mismatch." & vbCrLf
           errMsg = errMsg & "Expected header " & expectedColumnHeaders(ii) & " but file has " & fileHeaders(ii) & vbCrLf            
        End If
     Next          
   
   If Len(errMsg)>0 Then
      IsHeaderCorrect = errTitle & errMsg
   Else
      IsHeaderCorrect = vbTrue
   End If

End Function


'/////////////////////////////////////////////
' FUNCTION InsertBarcodesIntoWorklist
' Need this for normalization, the file goes to lims and
' Exemplar worklist does not contain gen2 data
' Put the (tube) barcodes into the worklist
' worklist is a valid file
' wellHeader, bcHeader: numeric ==> Header index
'                        string ==> Header Caption
' barcodes:      array ==> tubes, find the associated well, 1...96                        
'             otherwise ==> all barcodes are the same
'
' TODO:
'   Add error handling, no file, can't find headers/indices, etc
'/////////////////////////////////////////////
Function InsertBarcodesIntoWorklist(worklist, wellHeader, bcHeader, myBarcodes, totalWells)
Dim resp,text, line, header, ii, vals, jj
DIM idxWell, idxBC, wlString, wlArray, well
Dim errMsg

InsertBarcodesIntoWorklist = "success"

  ' DEBUG
 '  resp = "Configuring  " & worklist
 '  resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")
   resp =""

wlString = Read(worklist)
wlArray = Split(wlString, vbCrLf)

' Grab the header
text = wlArray(0)   ' vbCrLf added later...

' Sort the header indices
' assume for now that the headers are valid
idxWell = GetHeaderIndex(wlArray(0),wellHeader)
idxBC = GetHeaderIndex(wlArray(0),bcHeader)

' If idxBC not found, throw error
If idxWell<0 Or idxBC<0 Then
   errMsg = "Error in vbScript Library, InsertBarcodesintoWorklist:" & vbCrLf
   errMsg = errMsg & "Could not find either " & wellHeader & " or " & bcHeader & " in the header," & vbCrLf 
   errMsg = errMsg & text 
   InsertBarcodesIntoWorklist = errMsg
   Exit Function
End If

'  DEBUG
   resp = "Worklist: " & worklist
   resp = resp & "Well IDX; " & CStr(idxWell) & vbCrLf
   resp = resp & "BC IDX; " & CStr(idxBc) & vbCrLf
 ''  resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")

For ii=1 To Ubound(wlArray)
   vals = Split(wlArray(ii),",")

  ' DEBUG
'   resp = "UBOUND(vals) = " & UBound(vals) & vbCrLf
'   resp = resp & "Well IDX: " & CStr(idxWell) & vbCrLf
'   resp = resp & "BC IDX: " & CStr(idxBC) & vbCrLf
'   If (Ubound(vals)>= idxWell) AND (Ubound(vals)>= idxBC) Then
'      resp = resp & "Well " & vals(idxWell) & " = " & well & vbCrLf
'     '' resp = resp & "Barcode " & bar & vbCrLf
'   Else
'      resp = resp & "Indices out of range, UBOUND(vals) = " & UBound(vals)
''      resp = resp & "Well idx " & CStr(idxWell) & vbCrLf
'      resp = resp & "BC IDX; " & CStr(idxBC) & vbCrLf
'   End If
'   resp = resp & wlArray(ii) & vbCrLf
'   For Each v in vals
'      resp = resp & "<> " & v & vbCrLf
'   Next
'   resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")

   If (Ubound(vals)>= idxWell) AND (Ubound(vals)>= idxBC) Then
      well = WellClass.GetNumericID( vals(idxWell), totalWells )
      
      If well>0 Then
         If IsArray(myBarcodes) Then
            If myBarcodes(well) <> "0" Then
               vals(idxBC) = myBarcodes(well)
            Else
               ' Worklist expects a tube but there's not one in the scanned barcode list
               If errMsg = "" Then
                  erMsg = "Error in InsertBarcodesIntoWorklist" & vbCrLf
                  errMsg = errMsg & "The barcodes entered do not match the barcodes required by the worklist." & vbCrLf
                  errMsg = errMsg & "The following wells in the worklist are missing a tube barcode:" & vbCrLf & vals(idxWell)
               Else
                  errMsg = errMsg & vbCrLf & vals(idxWell)
               End If
            End If
            
         Else
            vals(idxBC) = myBarcodes
         End If
         line = vals(0)
         For jj=1 To UBound(vals)
            line = line & "," & vals(jj)
         Next
         text = text & vbCrLf & line        ' vbCrLf is required after every line.
      End If
   End If
   
   '  DEBUG
      resp = "New Line: " & line
   '  resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")
Next
If Len(errMsg)<=0 Then
   CALL Write(text, worklist, vbFalse, vbTrue) 
   InsertBarcodesIntoWorklist = "success"
Else
   InsertBarcodesIntoWorklist = errMsg
End If
End Function
'/////////////////////////////////////////////
   
'/////////////////////////////////////////////
' Function InsertWellsIntoWorklist(worklist, wellHeader, newWellValue)
'
' Take the incoming worklist file and replace the wells with
' the input well.  Useful when labware does not match up with
' Exemplar worklist, i.e. troughs are only A1, but exemplar is treating it like 96-well
' Added ability to use an array of wells, each well is used sequentially
' Input newWellValue is either a well or an array of wells in alpha-numeric form
'/////////////////////////////////////////////
Function InsertWellsIntoWorklist(worklist, wellHeader, newWellValue)
Dim resp, text, line, header, ii, vals, jj, v
DIM idxWell,wlString, wlArray, well
Dim errMsg, newWellCtr, ele

InsertWellsIntoWorklist = "success"

 ' DEBUG
  resp = "InsertWellsIntoWorklist.... Configuring  " & worklist & vbCrLf
  resp = resp & "well header = " & wellHeader & vbCrLf
  If IsArray(newWellValue) Then
     For Each ele In newWellValue
        resp = resp & "new well value = " & ele & vbCrLf
     Next
  else
     resp = resp & "new well value = " & newWellValue & vbCrLf
  End if
'  Call World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")
  resp =""

wlString = Read(worklist)
If InStr(LCase(wlString), "error")>0 Then
   errMsg = "Error in vbScript Library, InsertWellsIntoWorklist reading worklist:" & vbCrLf & wlString
   InsertWellsIntoWorklist = errMsg
   Exit Function
End If
wlArray = Split(wlString, vbCrLf)
   
' Grab the header
text = wlArray(0)   ' vbCrLf added later...

' Sort the header index, if idxWell<0, throw error
idxWell = GetHeaderIndex(wlArray(0),wellHeader)
If idxWell<0 Then
   errMsg = "Error in vbScript Library, InsertWellsIntoWorklist:" & vbCrLf
   errMsg = errMsg & "Could not find " & wellHeader & " in the header," & vbCrLf 
   errMsg = errMsg & text 
   InsertWellsIntoWorklist = errMsg
   Exit Function
End If

'  DEBUG
   resp = "Worklist: " & worklist & vbCrLf & "Well IDX = " & CStr(idxWell) & vbCrLf
'   resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")

newWellCtr = 0
For ii=1 To Ubound(wlArray)
   vals = Split(wlArray(ii),",")

  ' DEBUG
   resp = "UBOUND(vals) = " & UBound(vals) & vbCrLf
   resp = resp & "Well IDX: " & CStr(idxWell) & vbCrLf
   If (Ubound(vals)>= idxWell) Then
      resp = resp & "Well " & vals(idxWell) & " = " & well & vbCrLf
   Else
      resp = resp & "Indices out of range, UBOUND(vals) = " & UBound(vals)
      resp = resp & " but Well idx = " & CStr(idxWell) & vbCrLf
   End If
   resp = resp & wlArray(ii) & vbCrLf
   For Each v in vals
      resp = resp & "<> " & v & vbCrLf
   Next
'''   Call World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")

   If (Ubound(vals)>= idxWell) Then
      If IsArray(newWellValue) Then
         vals(idxWell) = newWellValue(newWellCtr)        
         newWellCtr = newWellCtr + 1
         If newWellCtr>UBound(newWellValue) Then
            newWellCtr = 0
         End If
      Else
         vals(idxWell) = newWellValue
      End If
      line = vals(0)
      For jj=1 To UBound(vals)
         line = line & "," & vals(jj)
      Next
      text = text & vbCrLf & line        ' vbCrLf is required after every line.
   End If
   
   '  DEBUG
   '   Call World.Globals.PauseGenerator.BtnPromptUser("New Line: " & line, Array("OK"), "OK")
Next
If Len(errMsg)<=0 Then
   CALL Write(text, worklist, vbFalse, vbTrue)       
Else
   InsertWellsIntoWorklist = errMsg
End If
End Function
'/////////////////////////////////////////////   
   
'/////////////////////////////////////////////
' Function InsertLabwareNamesIntoWorklist(exFileName,bkFileName)
'
' Take the incoming Exemplar file and replace the barcodes with
' the labware name.  Used at Ligation, Normalization, and Pooling
' 
' exFineName: file including path of the Exemplar worklist
' bkFileName: file name and path of the biomek worklist used for
'              TFF step in the Biomek Method
'
' How to decide which headers?
' whats the header for ligation, normalization, pooling
'
' For each header where a barcode is replaced by labware name, 
' add the header string to the CASE loop.
'
'/////////////////////////////////////////////
Function InsertLabwareNamesIntoWorklist(exFileName,bkFileName)
DIM deckClass, resp
DIM line, text, vals, newText, name, ii
Dim header,exSrcBcHeader,exDstBcHeader, exMidHeader, exBlockerHeader, exProbeHeader, exTweenEbHeader
Dim headerList, bcIdx
'                 0                 1              2                      3                      4                   5                 6
''SourceRackBarcode,SourceTubeBarcode,SourcePosition,DestinationRackBarcode,DestinationTubeBarcode,DestinationPosition,SourceVolumeToUse"                       7                     8 
''SourceRackBarcode,SourceTubeBarcode,SourcePosition,DestinationRackBarcode,DestinationTubeBarcode,DestinationPosition,IndexAssignmentName,LigationMIDTrayBarcode,LigationIndexPosition                                 
'                                                                                                                                              6                             7                          8                        9
''SourceRackBarcode,SourceTubeBarcode,SourcePosition,DestinationRackBarcode,DestinationTubeBarcode,DestinationPosition,RackToPlateTransferVolume,RackToPlateBlockerTrayBarcode,RackToPlateBlockerPosition,RackToPlateBlockerVolume
Set deckClass = New Deck
Set headerList = CreateObject("Othros.VariantList")
 
exSrcBcHeader = "SourceRackBarcode"
exDstBcHeader = "DestinationRackBarcode"
exMidHeader = "LigationMIDTrayBarcode"   
' Skip for now until headers get worked out
''exTweenEbHeader = "TweenEBRack"
exBlockerHeader = "RackToPlateBlockerTrayBarcode"  
exProbeHeader = "HybPrepProbePlateBarcode"                         

InsertLabwareNamesIntoWorklist = "success"
resp = ""
text = Read(exFileName)
'World.Globals.PauseGenerator.BtnPromptUser "InsertLabwareNames...Read " & exFileName & vbCrLf & text, Array("OK"), "OK"
If InStr(LCase(text),"error")>0 Then
   InsertLabwareNamesIntoWorklist = "Error From InsertLabwareNames" & vbCrLf & text
   Exit Function
End If
newText = ""
For Each line In Split(text, vbCrLf)
   If line = "" Then
   ' Header, figure out the indices
   ElseIf newText="" Then
      newText = line    
      ii=0
      For Each header in Split(line,",")
         Select Case LCase(header)
            Case LCase(exSrcBcHeader)
               headerList.Add(ii)
            Case LCase(exDstBcHeader)
               headerList.Add(ii)
            Case LCase(exMidHeader)
               headerList.Add(ii)
            Case LCase(exBlockerHeader)
               headerList.Add(ii)
            Case LCase(exProbeHeader)
               headerList.Add(ii)
         End Select
         ii = ii + 1
      Next
  
   Else
' Can't just replace a barcode in the line since barcodes may be similar
      vals = Split(line, ",")
      For Each bcIdx In headerList
      
         resp = "bcIdx " & bcIdx & "  vals(bcIdx) " & vals(bcIdx)
     ''    resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")
                  
         If bcIdx<=UBound(vals) Then
           ' Sometimes there is no destination, only source & reagent
            name = ""
            If Len(Trim(vals(bcIdx)))>0 Then
               name = DeckClass.GetLabwareNameFromBarcode(vals(bcIdx)) 
            End If
            
            resp = "Replacing " & vals(bcIdx) & " with " & name
     ''       resp = World.Globals.PauseGenerator.BtnPromptUser(resp, Array("OK"), "OK")
            
            If InStr(LCase(name),"error")>0 Then      
               InsertLabwareNamesIntoWorklist = "Error From InsertLabwareNames" & vbCrLf & name
               Exit Function
            Else
               vals(bcIdx) = name
            End If
         End If
      Next 
   ' Put the line back together
      newText = newText & vbCrLf
      For ii=0 To Ubound(vals)
         If ii=0 Then
            newText = newText & vals(ii)
         Else
            newText = newText & "," & vals(ii)
         End If
      Next
   End If
Next
Call Write(newText, bkFileName, vbFalse, vbTrue)
End Function   

'/////////////////////////////////////////////

   
'/////////////////////////////////////////////
' Function InsertDataIntoWorklist(worklist,dataHeader,data)
'
' Take the incoming worklist and replace each cell in the dataHeader
' column with the given data.
' 
' worklist: file including path of the worklist to be modified
' dataHeader: index (string or number) of the data column
' data: data to replace in each cell, if data is an array
' indices are 0...n, n=number of lines-1 in worklist (skip the header)
' if data is a single value, then all the cells are over-written with the value
'
'/////////////////////////////////////////////
Function InsertDataIntoWorklist(worklist, dataHeader, data)
DIM resp, headerIdx, headers
DIM line, text, textArray, ii, jj, vals, newText
                       
On Error Resume Next
InsertDataIntoWorklist = "success"
resp = ""
text = Read(worklist)
'World.Globals.PauseGenerator.BtnPromptUser "InsertDataIntoWorklist...Read " & worklist & vbCrLf & text, Array("OK"), "OK"
If InStr(LCase(text),"error")>0 Then
   InsertDataIntoWorklist = "Error From InsertDataIntoWorklist" & vbCrLf & text
   Exit Function
Else
   textArray = Split(text, vbCrLf)   
End If

' If header index not found, throw an error
headerIdx = GetHeaderIndex(textArray(0),dataHeader) 
If headerIdx<0 Or headerIdx>UBound(Split(textArray(0),",")) Then
   errMsg = "Error in vbScript Library, InsertDataintoWorklist:" & vbCrLf
   errMsg = errMsg & "Could not find matching header from input dataHeader=" & dataHeader & " in the header," & vbCrLf 
   errMsg = errMsg & textArray(0) & vbCrLf & vbCrLf
   errMsg = errMsg & "Either no match was found in the headers or the index of the header exceeded the number of columns." 
   InsertDataIntoWorklist = errMsg
   Exit Function
End If      

'Insert data into each cell and rebuild the text file
newText = textArray(0)
For ii=1 To Ubound(textArray) 
   If Len(textArray(ii))<UBound(Split(textArray(0),",")) Then
   Else
      vals = Split(textArray(ii), ",")
      If IsArray(data) Then
         vals(headerIdx) = data(ii)
      Else
         vals(headerIdx) = data
      End If      
   End If
    
   ' Put the line back together
   newText = newText & vbCrLf
   For jj=0 To Ubound(vals)
      If jj=0 Then
         newText = newText & vals(jj)
      Else
         newText = newText & "," & vals(jj)
      End If
   Next
Next
Call Write(newText, worklist, vbFalse, vbTrue)
If Err.Number <> 0 Then
   InsertDataIntoWorklist = "Error in vbScript Library's InsertDataWorklist:" & vbCrLf & Err.Description
   Err.Clear
End If

End Function   

'/////////////////////////////////////////////
   
      
'/////////////////////////////////////////////
' Function CreateGen2WgsPoolingWorklist
'
' Take the incoming Exemplar Pooling File and convert it to the 
' Gen2 WGS pooling file, so it can be sent to Gen2 lims
'
' WGS/Midpooling
' LOCATION,sRack,SOURCE_BARCODE,dRack,dWell,volume,POOL_BARCODE	POOL_GROUP,POOL_LIB_TYPE,POOL_PERCENT_DECIMAL,COMMENTS
'
' Exemplar Pooling Worklist
' 0                 1                 2              3                      4                      5                   6               
' SourceRackBarcode,SourceTubeBarcode,SourcePosition,DestinationRackBarcode,DestinationTubeBarcode,DestinationPosition,SourceVolumeToUse
'
Function CreateGen2WgsPoolingWorklist(exFileName,g2FileName)
DIM deckClass, wellClass, resp
DIM exLine, exText, g2line, g2Text, exVals, file, exHeader, g2Header
Dim exSrcBcIdx,exSrcTubeBcIdx,exSrcWellIdx,exDstBcIdx,exDstTubeBcIdx,exDstWellIdx,exVolIdx
DIM g2SrcName,g2SrcWell,g2SrcTubeBc,g2Vol
DIM g2DstName,g2DstWell,g2DstTubeBc

Set wellClass = New Wells
Set deckClass = New Deck

exSrcBcIdx = 0
exSrcTubeBcIdx = 1
exSrcWellIdx = 2
exDstBcIdx = 3
exDstTubeBcIdx = 4
exDstWellIdx = 5
exVolIdx = 6

' LOCATION,sRack,SOURCE_BARCODE,dRack,dWell,volume,POOL_BARCODE	POOL_GROUP,POOL_LIB_TYPE,POOL_PERCENT_DECIMAL,COMMENTS
g2SrcName = "sRack"
g2SrcWell = "LOCATION"
g2SrcTubeBc = "SOURCE_BARCODE"
g2Vol = "volume"
g2DstName = "dRack"
g2dstWell = "dWell"
g2DstTubeBc = "POOL_BARCODE"

exHeader = "SourceRackBarcode,SourceTubeBarcode,SourcePosition,DestinationRackBarcode,DestinationTubeBarcode,DestinationPosition,SourceVolumeToUse"
g2Header = " LOCATION,sRack,SOURCE_BARCODE,dRack,dWell,volume,POOL_BARCODE,POOL_GROUP,POOL_LIB_TYPE,POOL_PERCENT_DECIMAL,COMMENTS"

resp = ""
g2Text = ""
exText = Read(exFileName)
' World.Globals.PauseGenerator.BtnPromptUser "Read " & exFileName & vbCrLf & exText, Array("OK"), "OK"
If InStr(LCase(exText),"error")>0 Then
   CreateGen2WgsPoolingWorklist = "Error From CreateGen2WgsPoolingWorklist" & vbCrLf & exText
   Exit Function
End If
For Each exLine In Split(exText, vbCrLf)

 'World.Globals.PauseGenerator.BtnPromptUser exLine, Array("OK"), "OK"

   If g2Text = "" Then
      resp = IsHeaderCorrect(exLine,exHeader)
      If InStr(LCase(resp),"error")>0 Then
         CreateGen2WgsPoolingWorklist = "Error From CreateGen2WgsPoolingWorklist" & vbCrLf & resp
         Exit Function
      End If
      g2Text = g2Header
      
   ElseIf (Len(exLine) >= exVolIdx) Then

    'World.Globals.PauseGenerator.BtnPromptUser exLine, Array("OK"), "OK"

      exVals  = Split(exLine, ",")
      g2Line = g2header

      g2Line = Replace(g2line,"POOL_GROUP,POOL_LIB_TYPE,POOL_PERCENT_DECIMAL,COMMENTS", " , , , ")

      g2Line = Replace(g2Line, g2SrcName, DeckClass.GetLabwareNameFromBarcode(exVals(exSrcBcIdx)) )
      g2Line = Replace(g2Line, g2SrcWell, exVals(exSrcWellIdx))
      g2Line = Replace(g2Line, g2SrcTubeBc, exVals(exSrcTubeBcIdx))

      g2Line = Replace(g2Line, g2Vol, exVals(exVolIdx))

      g2Line = Replace(g2Line, g2DstName, DeckClass.GetLabwareNameFromBarcode(exVals(exDstBcIdx)) )
      g2Line = Replace(g2Line, g2DstWell, exVals(exDstWellIdx))
      g2Line = Replace(g2Line, g2DstTubeBc, exVals(exDstTubeBcIdx))

      g2Text = g2Text & vbCrLf & g2line
      
   Else
   
   End If
Next
If resp = vbTrue Then
   resp = "Writing " & g2FileName & vbCrLf
   resp = resp & g2text
'   World.Globals.PauseGenerator.BtnPromptUser resp, Array("OK"), "OK"
   Call Write(g2Text, g2FileName, vbFalse, vbTrue)
   resp = vbTrue
Else
   resp = "Error in Create Gen2 WGS Pooling Worklist: " & vbCrLf & "NOT Writing " & g2FileName & vbCrLf
   resp = resp & g2text
'   World.Globals.PauseGenerator.BtnPromptUser , Array("OK"), "OK"
End If
CreateGen2WgsPoolingWorklist = resp

End Function

'/////////////////////////////////////////////

'/////////////////////////////////////////////
' Function CreateGen2NormalizationWorklist
'
' Take the incoming Exemplar Normalization File and convert it to the 
' Gen2 normalization file, so it can be sent to Gen2 lims

' Gen2 Sample
'LOCATION	Source_Rack	BARCODE	Destination_Rack	DEST_BARCODE	Source_Volume	Tween_EB_Rack	EB_Well		EB_Volume	Tween_Well	Tween_Volume	APPLICATION	LIBRARY_TYPE	LIBRARY_CONCENTRATION	LIBRARY_SIZE	COMMENTS
'LOCATION	Source_Rack	BARCODE	Destination_Rack	DEST_BARCODE	Source_Volume	Tween_EB_Rack	Tween_EB_Well	Tween_EB_Volume					FINAL_DILUTION	PLATFORM_APPLICATION	LIBRARY_TYPE	LIBRARY_CONCENTRATION	LIBRARY_SIZE	COMMENTS
'LOCATION	Source_Rack	BARCODE	Destination_Rack	DEST_BARCODE	Source_Volume	Tween_EB_Rack	EB_Well		EB_Volume	Tween_Well	Tween_Volume	APPLICATION	LIBRARY_PASS_CONCENTRATION (nM)	PLATFORM	LIBRARY_TYPE	LIBRARY_CONCENTRATION	LIBRARY_SIZE	COMMENTS
'LOCATION	Source_Rack	BARCODE	Destination_Rack	DEST_BARCODE	Source_Volume	EB_Rack		EB_Well		EB_Volume					FINAL_DILUTION	PLATFORM_APPLICATION	LIBRARY_TYPE	LIBRARY_CONCENTRATION	LIBRARY_SIZE	COMMENTS
'Location	Source_Rack	BARCODE	Destination_Rack	DEST_BARCODE	Source_Volume	EB_Rack		EB_Well		EB_Volume					APPLICATION	LIBRARY_PASS_CONCENTRATION (nM)	PLATFORM	LIBRARY_TYPE	LIBRARY_CONCENTRATION	LIBRARY_SIZE	COMMENTS
' Outlier... Ignore for now...
'sTray	sTrayBarcode	sWell	sTubeBarcode	sVol	dTray	dTrayBarcode	dWell	dTubeBarcode	Tween_EB_Rack	EB_Well	EB_Volume	Tween_Well	Tween_Volume
' Exemplar Sample
'SourceRackBarcode,SourceTubeBarcode,SourcePosition,DestinationRackBarcode,DestinationTubeBarcode,DestinationPosition,EBWell,SourceVolToUseAliq1,TweenEBRack,TweenVolume,EBVolume,TweenWell
'CON_RACK_KapaHyper_1681_C3_D3,TestSampleKapaHyperLib-1919,C3,DestAliquot,DestAliquotBC1,C3,{EBWell=A1, SourceVolToUseAliq1=50.0, TweenEBRack=TWEEBRack, TweenVolume=10.156408693885844, EBVolume=41.40767824497258, TweenWell=A3}
'
'Gen2 Files are not consistent, convert to just match the headers necessary for the Biomek worklist
'and copy the other columns
' TODO:
' Error if file not found??
' Header is wrong?
' How to vet the line? 
'///////////////////////////////////////
Function CreateGen2NormalizationWorklist(exFileName,g2FileName)
DIM deckClass, wellClass, resp
DIM exLine, exText, g2line, g2Text, exVals, file, exHeader, g2Header
Dim exSrcBcIdx,exSrcTubeBcIdx,exSrcWellIdx,exDstBcIdx,exDstTubeBcIdx,exDstWellIdx
DIM exEbWellIdx,exVolIdx,exTweenEbRackIdx,exTweenVolIdx,exEbVolIdx,exTweenWellIdx
DIM g2SrcName,g2SrcWell,g2SrcTubeBc,g2Vol
DIM g2DstName,g2DstTubeBc,g2TweenEbRackName,g2EbWell,g2EbVol,g2TweenWell
DIM g2TweenVol, g2EbRack, g2TweenRack

On Error Resume Next

Set wellClass = New Wells
Set deckClass = New Deck
                        
' Exemplar Header
           '0                 1                 2              3                      4                      5                   6                   7           8         9      10          11 
           'SourceRackBarcode,SourceTubeBarcode,SourcePosition,DestinationRackBarcode,DestinationTubeBarcode,DestinationPosition,SourceVolToUseAliq1,TweenEBRack,TweenWell,EBWell,TweenVolume,EBVolume
exHeader = "SourceRackBarcode,SourceTubeBarcode,SourcePosition,DestinationRackBarcode,DestinationTubeBarcode,DestinationPosition,EBWell,SourceVolToUseAliq1,TweenEBRack,TweenVolume,EBVolume,TweenWell"
exHeader = "SourceRackBarcode,SourceTubeBarcode,SourcePosition,DestinationRackBarcode,DestinationTubeBarcode,DestinationPosition,SourceVolToUseAliq1,TweenEBRack,TweenWell,EBWell,TweenVolume,EBVolume"

exSrcBcIdx = 0
exSrcTubeBcIdx = 1
exSrcWellIdx = 2
exDstBcIdx = 3
exDstTubeBcIdx = 4
exDstWellIdx = 5
exEbWellIdx = 9
exVolIdx = 6
exTweenEbRackIdx = 7
exTweenVolIdx = 10
exEbVolIdx = 11
exTweenWellIdx = 8 
      
' Gen2 Header
' NOTE Headers in ALL CAPS are required
'0          1           2        3                 4              5              6       7          8           9           10              11                          12                      13              14
'LOCATION   Source_Rack BARCODE  Destination_Rack  DEST_BARCODE   Source_Volume  EB_Rack Tween_Rack,EB_Well,EB_Volume,Tween_Well,Tween_Volume,APPLICATION,LIBRARY_PASS_CONCENTRATION (nM),PLATFORM,LIBRARY_TYPE,LIBRARY_CONCENTRATION,LIBRARY_SIZE,COMMENTS
'LOCATION	Source_Rack	BARCODE	Destination_Rack	DEST_BARCODE	Source_Volume	Tween_EB_Rack	Tween_EB_Well	Tween_EB_Volume	LIBRARY_PASS_CONCENTRATION (nM)	PLATFORM	APPLICATION	LIBRARY_TYPE	LIBRARY_CONCENTRATION	LIBRARY_SIZE	POST_PCR_CYCLE	COMMENTS
'library  'LOCATION	Source_Rack	BARCODE	Destination_Rack	DEST_BARCODE	Source_Volume	Tween_EB_Rack	EB_Well	EB_Volume	Tween_Well	Tween_Volume	APPLICATION	PLATFORM	LIBRARY_PASS_CONCENTRATION (nM)	LIBRARY_TYPE	LIBRARY_CONCENTRATION	LIBRARY_SIZE	COMMENTS 
'from Sam 'LOCATION	Source_Rack	BARCODE	Destination_Rack	DEST_BARCODE	Source_Volume	Tween_EB_Rack	Tween_EB_Well	Tween_EB_Volume	LIBRARY_PASS_CONCENTRATION (nM)	PLATFORM	APPLICATION	LIBRARY_TYPE	LIBRARY_CONCENTRATION	LIBRARY_SIZE	POST_PCR_CYCLE	COMMENTS
g2Header = "LOCATION,Source_Rack,BARCODE,Destination_Rack,DEST_BARCODE,Source_Volume,Tween_EB_Rack,EB_Well,EB_Volume,Tween_Well,Tween_Volume,APPLICATION,PLATFORM,LIBRARY_PASS_CONCENTRATION (nM),LIBRARY_TYPE,LIBRARY_CONCENTRATION,LIBRARY_SIZE,POST_PCR_CYCLE,COMMENTS"

g2SrcName = "Source_Rack"
'''''g2SrcBc = 1
g2SrcWell = "LOCATION"
g2SrcTubeBc = "BARCODE"
g2Vol = "Source_Volume"
'''''g2DstBc = "dTrayBarcode"
g2DstName = "Destination_Rack"
g2DstTubeBc = "DEST_BARCODE"
g2TweenEbRackName = "Tween_EB_Rack"
g2EbRack = "EB_Rack"
g2TweenRack = "Tween_Rack"
g2EbWell = "EB_Well"
g2EbVol = "EB_Volume"
g2TweenWell = "Tween_Well" 
g2TweenVol = "Tween_Volume"

' When file contains EB Rack and Tween Rack
'  g2Header = "Location,Source_Rack,Barcode,Destination_Rack,Dest_Barcode,Source_Volume,EB_Rack,Tween_Rack,EB_Well,EB_Volume,Tween_Well,Tween_Volume,APPLICATION,LIBRARY_PASS_CONCENTRATION (nM),PLATFORM,LIBRARY_TYPE,LIBRARY_CONCENTRATION,LIBRARY_SIZE,COMMENTS"
' GEN2 requires these headers:
'LOCATION,BARCODE,DEST_BARCODE,PLATFORM,APPLICATION,LIBRARY_TYPE,LIBRARY_CONCENTRATION,LIBRARY_SIZE,COMMENTS

resp = ""
g2Text = ""
exText = Read(exFileName)
For Each exLine In Split(exText, vbCrLf)

 'World.Globals.PauseGenerator.BtnPromptUser exLine, Array("OK"), "OK"

   If g2Text = "" Then
      On Error Resume Next
      resp = IsHeaderCorrect(exLine,exHeader)

      If Err.Number <> 0 Then
         resp = "Error in vbScript Library's CreateGen2NormalizationWorklist:" & vbCrLf & Err.Description
         Err.Clear
         Exit For
      ElseIf resp<> vbTrue Then
         resp = "Error in vbScript Library's CreateGen2NormalizationWorklist:" & vbCrLf & resp
         Err.Clear
         Exit For
      Else
         g2Text = g2Header
      End If
   ElseIf Len(exLine) >= exTweenWellIdx Then

    'World.Globals.PauseGenerator.BtnPromptUser exLine, Array("OK"), "OK"

      exVals  = Split(exLine, ",")
      g2Line = g2header

      g2Line = Replace(g2line,"APPLICATION,PLATFORM,LIBRARY_PASS_CONCENTRATION (nM),LIBRARY_TYPE,LIBRARY_CONCENTRATION,LIBRARY_SIZE,POST_PCR_CYCLE,COMMENTS", " , , , , , , , ")

      ' Replace DEST_BARCODES before BARCODES
      g2Line = Replace(g2Line, g2DstTubeBc, exVals(exDstTubeBcIdx))
      g2Line = Replace(g2Line, g2DstName, DeckClass.GetLabwareNameFromBarcode(exVals(exDstBcIdx)) )

      g2Line = Replace(g2Line, g2SrcName, DeckClass.GetLabwareNameFromBarcode(exVals(exSrcBcIdx)) )
      g2Line = Replace(g2Line, g2SrcWell, exVals(exSrcWellIdx))
      g2Line = Replace(g2Line, g2SrcTubeBc, exVals(exSrcTubeBcIdx))

      g2Line = Replace(g2Line, g2Vol, exVals(exVolIdx)) 

      g2Line = Replace(g2Line, g2TweenEbRackName, exVals(exTweenEbRackIdx)) 
      g2Line = Replace(g2Line, g2EbWell, exVals(exEbWellIdx))
      g2Line = Replace(g2Line, g2EbVol, exVals(exEbVolIdx))
      g2Line = Replace(g2Line, g2TweenWell, exVals(exTweenWellIdx))
      g2Line = Replace(g2Line, g2TweenVol, exVals(exTweenVolIdx))

      g2Text = g2Text & vbCrLf & g2Line
   End If
Next
If resp = vbTrue Then
   resp = "Writing " & g2FileName & vbCrLf
   resp = resp & g2text
 '  World.Globals.PauseGenerator.BtnPromptUser resp, Array("OK"), "OK"
   Call Write(g2Text, g2FileName, vbFalse, vbTrue)
   resp = vbTrue
Else
   resp = "Error in vbScript Library's Gen2NormalizationWorklist: " & vbCrLf & resp
   resp = resp & "Can't create gen2 worklist " & g2FileName & vbCrLf
   resp = resp & vbCrLf & g2text
'   World.Globals.PauseGenerator.BtnPromptUser resp, Array("OK"), "OK"
End If
CreateGen2NormalizationWorklist = resp

End Function

' =====================================================

End Class    ' end class TextFile

' =====================================================

Class Deck

'/////////////////////////////////////////////
' FUNCTION IsDeckPosition
' Given a position's name, iterate through the current deck layout and match it up with its position object. 
'/////////////////////////////////////////////

Function IsDeckPosition(pos)
Dim p, deckPos

  IsDeckPosition = vbFalse
  Set deckPos = World.Devices.Pipettor1.Deck.Positions
  For p=0 To deckPos.VariantDictionary.Count-1
     If UCase(pos) =  UCase(deckPos.VariantDictionary.Values(p).Name) Then
        IsDeckPosition = vbTrue
        Exit For
     End If
     
  'DEBUG
  'resp = "'" & pos &"'<>'" & deckPos.VariantDictionary.Values(p).Name & "'"
  'World.Globals.PauseGenerator.BtnPromptUser resp, Array("OK"), "OK"
  
  Next
End Function  
'/////////////////////////////////////////////

'/////////////////////////////////////////////
' FUNCTION GetAlp
' Given the labware barcode, return
' the alp the plate occupies on the deck
' Returns an error OR alp or vbFalse
'/////////////////////////////////////////////
Function GetAlp(barcode)
 Dim alpFound, alp, key, pos, alpBC, resp
 Dim fileClass
       
   Set fileClass = New TextFile
   On Error Resume Next
   alpFound = vbFalse
   If Len(Trim(barcode))<=0 Then
      GetAlp = "Error in GetAlp: Can't find ALP because barcode is undefined."
      Call fileClass.WriteToDebugFile("", GetAlp, vbFalse)
      Exit Function
   End If
   GetAlp = "Error in GetAlp: No labware found with "&_
            "barcode " & barcode
   For alp=0 To Positions.VariantDictionary.Count-1
     key = Positions.VariantDictionary.Keys(alp)
     Set pos = Positions.VariantDictionary.Get(key)
     ' Watch out for ALPs with no labware or labware properties etc
      alpBC = pos.Labware.Properties.Barcode
      resp = "GetAlp looking for " & barcode & " on " & key
      Call fileClass.WriteToDebugFile("", resp, vbFalse)
      If err.Number<>0 Then
         ' ignore empty ALPs  
         Err.Clear
      Else
       ' Barcodes may be concatenated, exemplarBarcode+gen2LimsBarcode
         If (alpBC = barcode) Or _
            ( (InStr(alpBC, barcode)>0) And (InStr(alpBC, "+")>0) ) Then
                 alpFound = vbTrue
                 GetAlp = key
                 Exit For
         End If
      End If
   Next
   
  'DEBUG
  resp = "Results From GetAlp: " & vbCrLf & "Alp with barcode " & barcode & " is " & GetAlp
  Call fileClass.WriteToDebugFile("", resp, vbFalse)
  
 End Function


' =====================================================
'////////////////////////////////////////
' Function GetLabwareNameFromBarcode
' Given the barcode of pievce of labware on the deck, 
' return the labware name.
'///////////////////////////////////////
Function GetLabwareNameFromBarcode(barcode)
   DIM alp, pos, key, alpBC, resp
  
   On Error Resume Next
   GetLabwareNameFromBarcode = "Error in GetLabwareNameFromBarcode: No labware name found with "&_
      "barcode " & barcode
   For alp=0 To Positions.VariantDictionary.Count-1
      key = Positions.VariantDictionary.Keys(alp)
      Set pos = Positions.VariantDictionary.Get(key)
     ' Watch out for ALPs with no labware or labware properties etc
      alpBC = pos.Labware.Properties.Barcode
      If err.Number<>0 Then
         ' ignore empty ALPs  
         Err.Clear
         alpBC = ""
      Else
         ' Barcodes may be concatenated, exemplarBarcode+gen2LimsBarcode
         resp = "GetLabwareNameFromBarcode:  Checking Alp " & alp & " with barcode " & alpBC & " for barcode " & barcode
         Call fileClass.WriteToDebugFile("", resp, vbFalse)
         If (alpBC = barcode) Or _
            ( (InStr(alpBC,barcode)>0) And (InStr(alpBC, "+")>0) ) Then
              GetLabwareNameFromBarcode = pos.Labware.Properties.Name
              Exit For
         End If
      End If
   Next

End Function
' =====================================================


End Class    ' end class Deck

' =====================================================


' =====================================================

Class Patterns

' =====================================================

Private header 
Private active

Private Sub Class_Initialize
  header = "Well,Active"
  active = 1
End Sub      


'/////////////////////////////////////////////
' Function GetPatternFromWorklist
' Read the pattern from a worklist
' Returns the pattern array
'/////////////////////////////////////////////
' =====================================================
Function GetPatternFromWorklist(fromFile, wellHeader, dataHeader)
	Dim fileClass, patternArray, well, text
	
	Set fileClass = New TextFiles
		
   patternArray = fileClass.GetWellDataArray(fromFile, wellHeader, dataHeader)
   If Not IsArray(patternArray) Then
      WritePattern = patternArray
      Exit Function
   End If
   
   ' Convert well data to 1 (use well) or 0 (skip well)
   For well=0 To UBound(patternArray)
      If CStr(patternArray(well)) <> "0" Then
         patternArray(well) = active
      End If
   Next
   
  GetPatternFromWorklist = patternArray
  
End Function
' =====================================================


'/////////////////////////////////////////////
' Function WritePattern
' Write the pattern to the file
' Returns the pattern array
'/////////////////////////////////////////////
' =====================================================
Function WritePatternToFile(toFile, patternArray)
	Dim fileClass,  well, text
	Dim resp, errMsg, myError
	
	Set fileClass = New TextFiles
	
   If Not IsArray(patternArray) Then
      WritePatternToFile = "Error from WritePatternToFile, input parameter, patternArray, is not an array."
      Exit Function
   End If
   
   Error=vbFalse
   text = header
   errMsg = "Error from WritePatterToFile, pattern array can only contain 1's and 0's."
   For well=1 To UBound(patternArray)
      If (CInt(patternArray(well))=0) OR (CInt(patternArray(well))=1) Then
         text = text & vbCrLf & CStr(well) & "," & CStr(patternArray(well))
      Else
         myError=vbTrue
         errMsg = errMsg & vbCrLf &  "Data in well " & CStr(well) & " is " & CStr(patternArray(well))
      End If
   Next

  If myError Then
     WritePatternToFile = errMsg
     Exit Function
  Else
     WritePatternToFile = fileClass.Write(text, toFile, vbFalse, vbTrue)
  End If

End Function
' =====================================================

'/////////////////////////////////////////////
' Function ShiftPattern
' Create a new pattern by shifting input pattern
' over to the next empty column
'
' ASSUMES:
' every-other column is empty
' last column is empty
'/////////////////////////////////////////////
' =====================================================
Function ShiftPattern(patternArray)
	Dim well, shiftArray
		
	ReDim shiftArray(UBound(patternArray))
	For well = 0 To UBound(shiftArray)
	   shiftArray(well) = 0
	Next
	For well = 0 To UBound(shiftArray)
	   If CInt(patternArray(well)) = active Then
	      If well<UBound(shiftArray) Then
	         shiftArray(well+1) = active
	      End If
	   End If
	Next
   ShiftPattern = shiftArray
  
End Function
' =====================================================

End Class    ' end class Patterns

' =====================================================