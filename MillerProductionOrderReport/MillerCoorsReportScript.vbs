' *****************************************************************************************
' * NAME:       process.vbs
' *
' * PURPOSE:    Read Miller Coors Pabst production report and generate usable file.
' *
' * AUTHOR:     Jeremy Vance
' *
' * DATE:       10:06 AM 6/2/2016
' *
' * DIRECTIONS: Everything goes into a c:\scripts\mcplant directory on computer
' *             Save source report to name pvso.txt in directory
' *             Open a command prompt and run wscript process.vbs to create output data
' *             Results are saved as final.csv which can be opened in Excel
' * 
' * CONCEPT:    (1) Loop through source file to remove any deadspace and extraneous data so
' *             there is a workable WIP (work in process) file. 
' *             (2) Run through dynamic WIP file to extract plant, product, and measures
' *             (3) Write results to a CSV which can be used for further analysis
' *****************************************************************************************

On Error Resume Next

InputFileName = "pvso.txt"
WorkFileName = "wip01.txt"
OutFileName = "final.csv"
SuppressZeroes = 1
SuppressOutput = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objInFile = objFSO.OpenTextFile(InputFileName, 1, True)
Set objWIPFile = objFSO.CreateTextFile(WorkFileName, 2)
Set objOutFile = objFSO.CreateTextFile(OutFileName, 2)

'*********************************************************
'Process File to remove any blank lines, page headers, and 
'separators to get to just the key data
'*********************************************************

LineCount = 0
SkipCount = 0
Do While objInFile.AtEndOfStream = False
    LineCount = LineCount + 1

    strLine = objInFile.ReadLine
    WriteFlag = 1

    if instr(strLine, "CBRP022") > 0 then
	writeFlag = 0
    end if

    if instr(strLine, "MILLER PRODUCTION AND") > 0 then
        writeFlag = 0
    end if

    if instr(strLine, "P A B S T   B R E W I") > 0 then
        writeFlag = 0
    end if

    if len(ltrim(rtrim(strLine))) = 0 then
       writeFlag = 0
    end if

    if instr(strLine,"0************************************************************************************************************************************") > 0 then
       writeFlag = 0
    end if

    if instr(strLIne,"BRAND TOTAL") > 0 then
       skipCount = 14

       for i = 1 to skipCount
           strLine = objInFile.ReadLine
       next       

       writeFlag = 0
    end if

    if instr(strLine,"                        -----------------------------------------------------------------------------------------------------------  ") > 0 then
	writeFlag = 0
    end if

    if instr(strLIne,"PLANT TOTAL") > 0 then
       skipCount = 14

       for i = 1 to skipCount
           strLine = objInFile.ReadLine
       next  
    
       writeFlag = 0     
    end if

    if instr(strLIne,"COMPANY TOTAL") > 0 then
       skipCount = 14

       for i = 1 to skipCount
           strLine = objInFile.ReadLine
       next 
       
       writeFlag = 0      
    end if

    if instr(strLine,"***** END OF REPORT ***** ") > 0 then
       writeFlag = 0
    end if
       
    if instr(strLine,"                        -----------------------------------------------------------------------------------------------------------  ") > 0 then
	writeFlag = 0
    end if

    if instr(strLine," ************************************************************************************************************************************") > 0 then
	writeFlag = 0
    end if

    if instr(strLine,"(BBLS)") > 1 then
        writeFlag = 0
    end if

    if instr(strLine, "TOTAL") > 1 then
       writeFlag = 0
    end if

    if WriteFlag = 1 then
       objWIPFile.WriteLine strLine
    end if
Loop
'*********************************************************
'Close Files
'*********************************************************

objWIPFile.Close
objInFile.Close

'*********************************************************
'Initialize Values for WIP File Processing
'*********************************************************

FirstLine = 0
OutputFlag = 0
PlantFlag = 0
ProductFlag = 0
LineCount = 0
OutputCount = 0
UOMFlag = 0

prevPlantName = ""
prevPlantCode = ""

dim sweek(20)
dim pweek(20)
dim vweek(20)

for i = 1 to 20
  sweek(i) = ""
  pweek(i) = 0
  vweek(i) = ""
next

foundweek = -1

'*********************************************************
' Make File Header for Final Outfile.
'*********************************************************

outputLine = "LineCount,OutLineCount,PlantCode,PlantName,ProductCode,ProductName,UOM,WEEK,UNITS"             
'wscript.echo outputLine
objOutFile.WriteLine outputLine

'*********************************************************
' Run through Work In Process File to Extract Key Data to 
' a Recordset Format for outuput to CSV
'*********************************************************

Set objInFile = objFSO.OpenTextFile(WorkFileName, 1)
Do While objInFile.AtEndOfStream = False
    strLine = objInFile.ReadLine
    LineCount = LineCount + 1

	
'*********************************************************
    'Identify Plant Section and extract plant code
    'and product codes.  Also build out Week Array
    'structure
'*********************************************************

    if instr(strLine,"PLANT:") > 1 then
    'extract plant
	plantFlag = 1
	plantCode = right(left(strLine,len(" PLANT: ")+2),2)
	plantName = replace(ltrim(rtrim(right(strLine,len(strLine)-11))),",","")

        'skip a line to extract product
        tmpLine = strLine
        strLine = objInFile.ReadLine
		
		'edit by erik stryshak
		'accounts for duplicate plant line that isn't the same as the first line
		'or if there is more than one duplicate line
		'loopCount will cause the program to exit if the while loop goes more than
		'8 times, the loop gets stuck here if the last line of the file is a plant name
		loopCount = 0
		While instr(strLine,"PLANT:") > 1
		
			If (loopCount > 8) Then
				'close files
					objInFile.Close
					objOutFile.Close
				'unassign
					set objOutFile = Nothing
					set objWIPFile = Nothing
					set objInFile = Nothing
					set objFSO = Nothing	
				'quit the program
					wscript.quit
				ElseIf (tmpLine = strLine) Then
					strLine = objInFile.ReadLine
				Else
					plantCode = right(left(strLine,len(" PLANT: ")+2),2)
					plantName = replace(ltrim(rtrim(right(strLine,len(strLine)-11))),",","")
					strLine = objInFile.ReadLine
			End If
			loopCount = loopCount + 1
			
		Wend
		
			'store this plant's information for next iteration
			prevPlantCode = plantCode
			prevPlantName = plantName
			
			'get product and week information
			productFlag = 1
			productName = ltrim(rtrim(strLine))
			productCode = ltrim(rtrim(left(strLine,8)))
			productName = right(productName,len(productName)-len(productCode)-1)

			
			'skip to next line to begin extracting weeks
			strLine = objInFile.ReadLine
			if(instr(strLine,"INVENTORY") = 0) then
				for i = 1 to len(strLine)
					x = mid(strLine,i,1)
					if x <> " " then
					   foundweek = foundweek + 1
					   sweek(foundweek) = mid(strLine,i,8)
					   pweek(foundweek) = i

					   i = i + 8
					end if
				next
			end if
'*********************************************************
'Read Inventory Units
'*********************************************************
    ElseIf instr(strLine,"INVENTORY(UNITS)") > 0 then
       uomflag = 1
       uom = "Inventory Units"

         if foundweek > -1 then
               for i = 0 to foundweek
                 tstring = mid(strLine,pweek(i),8)
                    tstring = ltrim(rtrim(tstring))
                    tstring = replace(tstring,",","")
                    if len(tstring) = 0 then
                          tstring = "0"
                    end if
                    vweek(i) = tstring
               next
         end if

 '*********************************************************
   'Read Prod Plan Units
'*********************************************************

    ElseIf instr(strLine,"PROD PLAN(UNITS)") > 0 then
       uomflag = 1
       uom = "Prod Plan Units"
         if foundweek > -1 then
               for i = 0 to foundweek
                 tstring = mid(strLine,pweek(i),8)
                    tstring = ltrim(rtrim(tstring))
                    tstring = replace(tstring,",","")
                    if len(tstring) = 0 then
                          tstring = "0"
                    end if
                    vweek(i) = tstring
               next
         end if

'*********************************************************
    'Read Truck Orders Units
'*********************************************************

    ElseIf instr(strLine,"TRUCK ORDERS(UNITS)") > 0 then
       uomflag = 1
       uom = "Truck Order Units"
         if foundweek > -1 then
               for i = 0 to foundweek
                 tstring = mid(strLine,pweek(i),8)
                    tstring = ltrim(rtrim(tstring))
                    tstring = replace(tstring,",","")
                    if len(tstring) = 0 then
                          tstring = "0"
                    end if
                    vweek(i) = tstring
               next
         end if

 '*********************************************************
   'Read Rail Orders Units
'*********************************************************

    ElseIf instr(strLine,"RAIL ORDERS(UNITS)") > 0 then
       uomflag = 1
       uom = "Rail Order Units"
         if foundweek > -1 then
               for i = 0 to foundweek
                 tstring = mid(strLine,pweek(i),8)
                    tstring = ltrim(rtrim(tstring))
                    tstring = replace(tstring,",","")
                    if len(tstring) = 0 then
                          tstring = "0"
                    end if
                    vweek(i) = tstring
               next
         end if

'*********************************************************
    'Read Held Orders Units
'*********************************************************
    ElseIf instr(strLine,"HELD ORDERS(UNITS)") > 0 then
       uomflag = 1
       uom = "Held Order Units"
         if foundweek > -1 then
               for i = 0 to foundweek
                 tstring = mid(strLine,pweek(i),8)
                    tstring = ltrim(rtrim(tstring))
                    tstring = replace(tstring,",","")
                    if len(tstring) = 0 then
                          tstring = "0"
                    end if
                    vweek(i) = tstring
               next
         end if
	'condition if the plant code and name was not given
	Else
		plantFlag = 1
		plantName = prevPlantName
		plantCode = prevPlantCode
		
		'get product and week information
        productFlag = 1
        productName = ltrim(rtrim(strLine))
        productCode = ltrim(rtrim(left(strLine,8)))
        productName = right(productName,len(productName)-len(productCode)-1)
		
        'skip to next line to begin extracting weeks
        strLine = objInFile.ReadLine
		if(instr(strLine,"INVENTORY") = 0) then
			for i = 1 to len(strLine)
				x = mid(strLine,i,1)
				if x <> " " then
				   foundweek = foundweek + 1

				   sweek(foundweek) = mid(strLine,i,8)
				   pweek(foundweek) = i

				   i = i + 8
				end if
			next
		end if
	end if

'*********************************************************
    'Determine if a Line Write is needed
'*********************************************************
    
	if plantFlag = 1 then
        if productFlag = 1 then
			if UOMFlag = 1 then
				OutputFlag = 1
			end if
		end if
    end if

'*********************************************************
    'Write out line if needed
 '*********************************************************
   if OutputFlag = 1 then
            OutputCount = OutputCount + 1
            OutputFlag = 0
            uomflag = 1

            if foundweek > -1 then
                for i = 0 to foundweek
                     outputLine = cstr(LineCount) + "," + cstr(OutputCount) + "," + plantCode + "," + plantName + "," + productCode + "," + ProductName + "," + UOM + "," + sweek(i) + "," + vweek(i) 

                     if suppresszeroes = 1 then
                         if vweek(i) <> "0" then
                            if suppressoutput = 0 then
                            	Wscript.Echo outputLine
                            end if
                            objOutFile.WriteLine outputLine
                         end if
                     else
                         if suppressoutput = 0 then
                            Wscript.Echo outputLine
                         end if
                         objOutFile.WriteLine outputLine
                     end if                
                     
                next
            end if	    
    end if

 '*********************************************************
   'Determine if Plant and Product Flags to be Reset
 '*********************************************************
   if instr(strLine,"HELD ORDERS(UNITS)") > 0 then
       PlantFlag = 0
       ProductFlag = 0   

       PlantFlag = 0
       ProductFlag = 0
       UOMFlag = 0
       NextLine = 0

       for i = 1 to 20
          sweek(i) = ""
          pweek(i) = 0
          vweek(i) = ""
       next

       foundweek = -1
    end if
Loop

'*********************************************************
'Close Files
'*********************************************************
objInFile.Close
objOutFile.Close

'*********************************************************
'Release Resources
'*********************************************************
set objOutFile = Nothing
set objWIPFile = Nothing
set objInFile = Nothing
set objFSO = Nothing
wscript.quit