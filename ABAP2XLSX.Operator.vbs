
'-Begin-----------------------------------------------------------------
'-
'- Hint: Swap string $TMP to your package manually
'-
'-----------------------------------------------------------------------

  '-Directives----------------------------------------------------------
    Option Explicit

  '-Constants-----------------------------------------------------------
    Const ForReading = 1
    Const ForWriting = 2

  '-Variables-----------------------------------------------------------
    Dim FileName, FSO, CSVFile, Line, OldName, NewName, XMLFile, XML
    Dim Lines, i, LogFile

  '-Main----------------------------------------------------------------
    FileName = "ABAP2XLSX_V_7_0_5.nugg"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If IsObject(FSO) Then
      Set CSVFile = FSO.OpenTextFile(FileName & ".csv", ForReading)
      Set XMLFile = FSO.OpenTextFile(FileName, ForReading)
      Set LogFile = FSO.OpenTextFile(FileName & ".log", ForWriting, True)
      XML = XMLFile.ReadAll
      XMLFile.Close
      CSVFile.ReadLine
      Do
        Line = Split(CSVFile.ReadLine, ";")
        OldName = Line(2)
        NewName = Line(4)
        If InStr(UCase(XML), "_" & UCase(OldName)) <> 0 Then
          Lines = Split(XML, Chr(10))
          For i = 0 To UBound(Lines) - 1
            If InStr(UCase(Lines(i)), "_" & UCase(OldName)) = 0 Then
              Lines(i) = Replace(Lines(i), OldName, NewName, 1, -1, 1)
            Else
              LogFile.WriteLine "Check line " & CStr(i + 1)
            End If
          Next
          XML = ""
          For i = 0 To UBound(Lines) - 1
            XML = XML & Lines(i) & Chr(10)
          Next
          LogFile.WriteLine OldName & " > " & NewName
        Else
          XML = Replace(XML, OldName, NewName, 1, -1, 1)
          LogFile.WriteLine OldName & " > " & NewName
        End If
      Loop Until CSVFile.AtEndOfStream
      CSVFile.Close
      LogFile.Close
      Set XMLFile = FSO.OpenTextFile(FileName & ".new", ForWriting, True)
      XMLFile.WriteLine XML
      XMLFile.Close
      Set FSO = Nothing
    End If

'-End-------------------------------------------------------------------
