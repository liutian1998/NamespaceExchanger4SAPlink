
'-Begin-----------------------------------------------------------------

  '-Directives----------------------------------------------------------
    Option Explicit

  '-Constants-----------------------------------------------------------
    Const ForReading = 1

  '-Variables-----------------------------------------------------------
    Dim FileName, FSO, File, i, LenLine, OldLenLine, Zeile

  '-Main----------------------------------------------------------------
    FileName = "ABAP2XLSX_V_7_0_5.nugg"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If IsObject(FSO) Then
      Set File = FSO.OpenTextFile(FileName, ForReading)
      Do Until File.AtEndOfStream
        i = i + 1
        LenLine = Len(File.ReadLine)
        If LenLine > OldLenLine Then
          OldLenLine = LenLine
          Zeile = i
        End If
      Loop 
      File.Close
      Set FSO = Nothing
      MsgBox FileName & vbCrLf & vbCrLf & "Longest line is " & _
        CStr(Zeile) & " with " & OldLenLine & " characters", vbOkOnly
    End If

'-End-------------------------------------------------------------------
