Option Explicit
Const INPUT_FILE_1 = "c:\works\App\file1.txt"
Const INPUT_FILE_2 = "c:\works\App\file2.txt"
Const OUTPUT_FILE = "c:\works\App\file.log"
Const ForReading = 1
Const ForWriting = 2
Dim dicData1, dicData2, strKey, fso, ts, dicbal1, dicbal2
Dim filInput1, filInput2, filResults, strLine
Set fso = CreateObject("Scripting.FileSystemObject")
Set filInput1 = _
  fso.OpenTextFile(INPUT_FILE_1, ForReading)
Set filInput2 = _
  fso.OpenTextFile(INPUT_FILE_2, ForReading)
Set filResults = _
  fso.OpenTextFile(OUTPUT_FILE, ForWriting, True)
  filResults.WriteLine strKey & _
                  "********************** Processing started *********************"
  filResults.WriteLine strKey &_
                   "File                    "& INPUT_FILE_1 
  filResults.WriteLine strKey & _
                   "compare with the file     "& INPUT_FILE_2
 Set dicData1 = CreateObject("Scripting.Dictionary")
 While Not filInput1.AtEndOfStream
      strLine = filInput1.ReadLine
      strLine = Replace(strLine, """" , "")
      dicbal1 = Split(strLine,",")
      dicData1.Add dicbal1(2)&"_"& dicbal1(3),dicbal1(4)&dicbal1(5)&dicbal1(6)&dicbal1(7)&dicbal1(8)&dicbal1(9)&dicbal1(10)&dicbal1(11)&dicbal1(12) 
 Wend
filInput1.Close
 Set dicData2 = CreateObject("Scripting.Dictionary")
 While Not filInput2.AtEndOfStream
      strLine = filInput2.ReadLine
      strLine = Replace(strLine, """" , "")
      dicbal2 =  Split(strLine,",")
      dicData2.Add dicbal2(2)&"_"& dicbal2(3),dicbal2(4)&dicbal2(5)&dicbal2(6)&dicbal2(7)&dicbal2(8)&dicbal2(9)&dicbal2(10)&dicbal2(11)&dicbal2(12)

 Wend
filInput2.Close
 
For Each strKey In dicData1
      If Not dicData2.Exists(strKey) Then
            filResults.WriteLine strKey & ": Account balance data changed " & INPUT_FILE_1
		If dicbal2(0) <> dicbal1 (0) Then  filResults.WriteLine dicbal1(0)& "--------" & dicbal2 (0) End If
		If dicbal2(1) <> dicbal1 (1) Then  filResults.WriteLine dicbal1(1)& "--------" & dicbal2 (1) End If
		If dicbal2(2) <> dicbal1 (2) Then  filResults.WriteLine dicbal1(2) & "--------" & dicbal2 (2) End If
		If dicbal2(3) <> dicbal1 (3) Then  filResults.WriteLine dicbal1(3) & "--------" & dicbal2 (3) End If
		If dicbal2(4) <> dicbal1 (4) Then  filResults.WriteLine dicbal1(4) & "--------" & dicbal2 (4) End If
        If dicbal2(5) <> dicbal1 (5) Then  filResults.WriteLine dicbal1(5) & "--------" & dicbal2 (5) End if
        If dicbal2(6) <> dicbal1 (6) Then  filResults.WriteLine dicbal1(6) & "--------" & dicbal2 (6) End if
        If dicbal2(7) <> dicbal1 (7) Then  filResults.WriteLine dicbal1(7) & "--------" & dicbal2 (7) End if
        If dicbal2(8) <> dicbal1 (8) Then  filResults.WriteLine dicbal1(8) & "--------" & dicbal2 (8) End if
        If dicbal2(9) <> dicbal1 (9) Then  filResults.WriteLine dicbal1(9) & "--------" & dicbal2 (9) End if
        If dicbal2(10) <> dicbal1 (10) Then  filResults.WriteLine dicbal1(10) & "--------" & dicbal2 (10) End if
        If dicbal2(11) <> dicbal1 (11) Then  filResults.WriteLine dicbal1(11) & "--------" & dicbal2 (11) End if
        If dicbal2(12) <> dicbal1 (12) Then  filResults.WriteLine dicbal1(12) & "--------" & dicbal2 (12) End if
        If dicbal2(13) <> dicbal1 (13) Then  filResults.WriteLine dicbal1(13) & "--------" & dicbal2 (13) End if
        If dicbal2(14) <> dicbal1 (14) Then  filResults.WriteLine dicbal1(14) & "--------" & dicbal2 (14) End if
        If dicbal2(15) <> dicbal1 (15) Then  filResults.WriteLine dicbal1(15) & "--------" & dicbal2 (15) End If           

      Else
            If dicData2.Item(strKey) <> dicData1.Item(strKey) Then
                filResults.WriteLine strKey & _
                  ": Account balance data changed "& INPUT_FILE_2
		If dicbal2(0) <> dicbal1 (0) Then  filResults.WriteLine dicbal2(0)& "--------" & dicbal1 (0) End If
		If dicbal2(1) <> dicbal1 (1) Then  filResults.WriteLine dicbal2(1)& "--------" & dicbal1 (1) End If
		If dicbal2(2) <> dicbal1 (2) Then  filResults.WriteLine dicbal2(2)& "--------" & dicbal1 (2) End If
		If dicbal2(3) <> dicbal1 (3) Then  filResults.WriteLine dicbal2(3) & "--------" & dicbal1 (3) End If
		If dicbal2(4) <> dicbal1 (4) Then  filResults.WriteLine dicbal2(4) & "--------" & dicbal1 (4) End If
        If dicbal2(5) <> dicbal1 (5) Then  filResults.WriteLine dicbal2(5) & "--------" & dicbal1 (5) End if
        If dicbal2(6) <> dicbal1 (6) Then  filResults.WriteLine dicbal2(6) & "--------" & dicbal1 (6) End if
        If dicbal2(7) <> dicbal1 (7) Then  filResults.WriteLine dicbal2(7) & "--------" & dicbal1 (7) End if
        If dicbal2(8) <> dicbal1 (8) Then  filResults.WriteLine dicbal2(8) & "--------" & dicbal1 (8) End if
        If dicbal2(9) <> dicbal1 (9) Then  filResults.WriteLine dicbal2(9) & "--------" & dicbal1 (9) End if
        If dicbal2(10) <> dicbal1 (10) Then  filResults.WriteLine dicbal2(10) & "--------" & dicbal1 (10) End if
        If dicbal2(11) <> dicbal1 (11) Then  filResults.WriteLine dicbal2(11) & "--------" & dicbal1 (11) End if
        If dicbal2(12) <> dicbal1 (12) Then  filResults.WriteLine dicbal2(12) & "--------" & dicbal1 (12) End if
        If dicbal2(13) <> dicbal1 (13) Then  filResults.WriteLine dicbal2(13) & "--------" & dicbal1 (13) End if
        If dicbal2(14) <> dicbal1 (14) Then  filResults.WriteLine dicbal2(14) & "--------" & dicbal1 (14) End if
        If dicbal2(15) <> dicbal1 (15) Then  filResults.WriteLine dicbal2(15) & "--------" & dicbal1 (15) End If           
             End If
      End If
Next
For Each strKey In dicData2
      If Not dicData1.Exists(strKey) Then
            filResults.WriteLine "Balance account " & strKey & " missing in file " & INPUT_FILE_1
      End If
Next
 
filResults.WriteLine "Number of lines in file 1" & _
  "(" & dicData1.Count & ") Number of lines in file 2 (" & dicData2.Count & ")."

filResults.WriteLine "*********************** Parsing is over **********************" 

filResults.Close


 