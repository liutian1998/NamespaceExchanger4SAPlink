
'-Begin-----------------------------------------------------------------

  '-Directives----------------------------------------------------------
    Option Explicit

  '-Variables-----------------------------------------------------------
    Dim xmlDoc, xmlRoot, xmlNodes, xmlNode
    Dim Lines()
    Dim PrefixOld, PrefixNew, NodeName, NewNodeName, Signal, ClassName
    Dim FileName, i, Text, FSO, File

  '-Function CorrectNodeName--------------------------------------------
    Function CorrectNodeName(NodeName)

      '-Variables-------------------------------------------------------
        Dim NewNodeName

      NewNodeName = Replace(UCase(NodeName), "_EXCEL_", "_")
      NewNodeName = Replace(UCase(NewNodeName), "_WORKSHEET_", "_WKS_")
      NewNodeName = Replace(UCase(NewNodeName), "_CONVERTER", "_CONV")
      NewNodeName = Replace(UCase(NewNodeName), "_CONDITIONAL", "_COND")
      NewNodeName = Replace(UCase(NewNodeName), "_CONDITION", "_COND")
      NewNodeName = Replace(UCase(NewNodeName), "_COMPONENT", "_COMP")
      NewNodeName = Replace(UCase(NewNodeName), "_PROTECTION", "_PROT")
      NewNodeName = Replace(UCase(NewNodeName), "MAPPING", "MAP")
      If UCase(Right(NewNodeName, 14)) = "_NUMBER_FORMAT" Then
        NewNodeName = Left(NewNodeName, Len(NewNodeName) - 14) + "_NUMFMT"
      End If

      CorrectNodeName = NewNodeName

    End Function

  '-Main----------------------------------------------------------------
    PrefixOld = "Z"
    PrefixNew = "/MNS/MPA_"

    Set xmlDoc = CreateObject("MSXML.DOMDocument")
    FileName = "ABAP2XLSX_V_7_0_3.nugg"
    If IsObject(xmlDoc) Then
      xmlDoc.Async = False
      If xmlDoc.Load(FileName) Then
        Set xmlRoot = xmlDoc.documentElement
        Set xmlNodes = xmlRoot.childNodes
        ReDim Preserve Lines(1)
        Lines(0) = "Line;Type;OldName;ClassName;NewName;Length;Signal"
        For Each xmlNode In xmlNodes

          ReDim Preserve Lines(UBound(Lines) + 1)
          ClassName = ""

          Select Case UCase(xmlNode.nodeName)

            Case "CLAS"
              NodeName = xmlNode.getAttribute("CLSNAME")
              If UCase(Left(NodeName, 4)) = PrefixOld + "CL_" Then
                NewNodeName = PrefixNew + "CL_" + Right(NodeName, _
                  Len(NodeName) - Len(PrefixOLd) - 3)
              ElseIf UCase(Left(NodeName, 4)) = PrefixOld + "CX_" Then
                NewNodeName = Left(PrefixNew, 5) + "CX_" + _
                  Right(PrefixNew, Len(PrefixNew) - 5) + _
                  Right(NodeName, Len(NodeName) - 4)
              End If
              NewNodeName = CorrectNodeName(NewNodeName)
              If Len(NewNodeName) > 30 Then
                Signal = "1"
              Else
                Signal = "0"
              End If

            Case "DOMA"
              NodeName = xmlNode.getAttribute("DOMNAME")
              NewNodeName = PrefixNew + Right(NodeName, _
                Len(NodeName) - Len(PrefixOld))
              NewNodeName = CorrectNodeName(NewNodeName)
              If Len(NewNodeName) > 30 Then
                Signal = "1"
              Else
                Signal = "0"
              End If

            Case "DTEL"
              NodeName = xmlNode.getAttribute("ROLLNAME")
              NewNodeName = PrefixNew + Right(NodeName, _
                Len(NodeName) - Len(PrefixOld))
              NewNodeName = CorrectNodeName(NewNodeName)
              If Len(NewNodeName) > 30 Then
                Signal = "1"
              Else
                Signal = "0"
              End If

            Case "INTF"
              NodeName = xmlNode.getAttribute("CLSNAME")
              If UCase(Left(NodeName, 4)) = PrefixOld + "IF_" Then
                NewNodeName = PrefixNew + "IF_" + Right(NodeName, _
                  Len(NodeName) - Len(PrefixOld) - 3)
              End If
              NewNodeName = CorrectNodeName(NewNodeName)
              If Len(NewNodeName) > 30 Then
                Signal = "1"
              Else
                Signal = "0"
              End If

            Case "MSAG"
              NodeName = xmlNode.getAttribute("ARBGB")
              NewNodeName = PrefixNew + Right(NodeName, _
                Len(NodeName) - Len(PrefixOld))
              NewNodeName = CorrectNodeName(NewNodeName)
              If Len(NewNodeName) > 20 Then
                Signal = "1"
              Else
                Signal = "0"
              End If

            Case "PROG"
              NodeName = xmlNode.getAttribute("NAME")
              NewNodeName = PrefixNew + Right(NodeName, _
                Len(NodeName) - Len(PrefixOld))
              NewNodeName = CorrectNodeName(NewNodeName)
              If Len(NewNodeName) > 30 Then
                Signal = "1"
              Else
                Signal = "0"
              End If

            Case "TABL"
              NodeName = xmlNode.getAttribute("TABNAME")
              NewNodeName = PrefixNew + Right(NodeName, _
                Len(NodeName) - Len(PrefixOld))
              NewNodeName = CorrectNodeName(NewNodeName)
              Select Case UCase(xmlNode.getAttribute("TABCLASS"))
                Case "INTTAB"
                  ClassName = "INTTAB"
                  If Len(NewNodeName) > 30 Then
                    Signal = "1"
                  Else
                    Signal = "0"
                  End If
                Case "TRANSP"
                  ClassName = "TANSP"
                  If Len(NewNodeName) > 16 Then
                    Signal = "1"
                  Else
                    Signal = "0"
                  End If
                
              End Select

            Case "TTYP"
              NodeName = xmlNode.getAttribute("TYPENAME")
              NewNodeName = PrefixNew + Right(NodeName, _
                Len(NodeName) - Len(PrefixOld))
              NewNodeName = CorrectNodeName(NewNodeName)
              If Len(NewNodeName) > 30 Then
                Signal = "1"
              Else
                Signal = "0"
              End If

            Case "XSLT"
              NodeName = xmlNode.getAttribute("XSLTDESC")
              NewNodeName = PrefixNew + Right(NodeName, _
                Len(NodeName) - Len(PrefixOld))
              NewNodeName = CorrectNodeName(NewNodeName)
              If Len(NewNodeName) > 30 Then
                Signal = "1"
              Else
                Signal = "0"
              End If

          End Select

          Lines(UBound(Lines)) = CStr(UBound(Lines) - 1) & ";" & _
            UCase(xmlNode.nodeName) & ";" & NodeName & ";" & _
            ClassName & ";" & NewNodeName & ";" & _
            CStr(Len(NewNodeName)) & ";" & Signal

        Next

      End If
      Set xmlDoc = Nothing
    End If

    Text = Lines(0) & vbCrLf
    For i = UBound(Lines) To 2 Step -1
      Text = Text & Lines(i) & vbCrLf
    Next

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If IsObject(FSO) Then
      Set File = FSO.CreateTextFile(FileName & ".csv", True)
      File.Write Text
      File.Close
    End If

'-End-------------------------------------------------------------------
