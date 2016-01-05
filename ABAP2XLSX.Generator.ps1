
#-Begin-----------------------------------------------------------------

  #-Load assemblies-----------------------------------------------------
    [Reflection.Assembly]::LoadWithPartialName("'Microsoft.VisualBasic") > $Null

  #-Function Correct-NodeName-------------------------------------------
    Function Correct-NodeName([String] $NodeName) {

      $NewNodeName = $NodeName.ToUpper() -replace "_EXCEL_", "_"
      $NewNodeName = $NewNodeName.ToUpper() -replace "_WORKSHEET_", "_WKS_"
      $NewNodeName = $NewNodeName.ToUpper() -replace "_CONVERTER", "_CONV"
      $NewNodeName = $NewNodeName.ToUpper() -replace "_CONDITIONAL", "_COND"
      $NewNodeName = $NewNodeName.ToUpper() -replace "_CONDITION", "_COND"
      $NewNodeName = $NewNodeName.ToUpper() -replace "_COMPONENT", "_COMP"
      $NewNodeName = $NewNodeName.ToUpper() -replace "_PROTECTION", "_PROT"
      $NewNodeName = $NewNodeName.ToUpper() -replace "MAPPING", "MAP"

      If ($NewNodeName.Length -gt 14) {
        If ($NewNodeName.ToUpper().Substring($NewNodeName.Length - 14, 14) -eq "_NUMBER_FORMAT") {
          $NewNodeName = $NewNodeName.ToUpper().Substring(0, $NewNodeName.Length - 14) + "_NUMFMT"
        }
      }

      Return $NewNodeName

    }

  #-Sub Main------------------------------------------------------------
    Function Main () {

      $PrefixOld = "Z"
      $PrefixNew = "/MNS/MPA_"
      $FileName = "ABAP2XLSX_V_7_0_5.nugg"

      $xmlDoc = [XML] (Get-Content -Path $FileName)
      If ($xmlDoc -eq $Null) { 
        Break
      }

      $xmlRoot = $xmlDoc.SelectSingleNode("/*")
      $xmlNodes = $xmlRoot.ChildNodes
      [String[]]$Lines = @()
      $Lines += "Line;Type;OldName;ClassName;NewName;Length;Signal"

      ForEach ($xmlNode In $xmlNodes) {

        $ClassName = $Null

        Switch ($xmlNode.LocalName.ToUpper()) {
          "CLAS" {
            $NodeName = $xmlNode.CLSNAME
            If ($NodeName.Substring(0, 4).ToUpper() -eq $PrefixOld + "CL_") {
              $NewNodeName = $PrefixNew + "CL_" + $NodeName.Substring($PrefixOld.Length + 3, 
                $NodeName.Length - ($PrefixOld.Length + 3))
            }
            ElseIf ($NodeName.Substring(0, 4).ToUpper() -eq $PrefixOld + "CX_") {
              $NewNodeName = $PrefixNew.Substring(0, 5) + "CX_" + 
                $PrefixNew.Substring($PrefixNew.Length - 5, 4) + "_" + 
                $NodeName.Substring($PrefixOld.Length + 3, 
                $NodeName.Length - ($PrefixOld.Length + 3))
            }
            $NewNodeName = Correct-NodeName $NewNodeName
            If ($NewNodeName.Length -gt 30) {
              $Signal = 1
            }
            Else {
              $Signal = 0
            }
          }

          "DOMA" {
            $NodeName = $xmlNode.DOMNAME
            $NewNodeName = $PrefixNew + $NodeName.Substring($PrefixOld.Length,
              $NodeName.Length - $PrefixOld.Length)
            $NewNodeName = Correct-NodeName $NewNodeName
            If ($NewNodeName.Length -gt 30) {
              $Signal = 1
            }
            Else {
              $Signal = 0
            }
          }

          "DTEL" {
            $NodeName = $xmlNode.ROLLNAME
            $NewNodeName = $PrefixNew + $NodeName.Substring($PrefixOld.Length,
              $NodeName.Length - $PrefixOld.Length)
            $NewNodeName = Correct-NodeName $NewNodeName
            If ($NewNodeName.Length -gt 30) {
              $Signal = 1
            }
            Else {
              $Signal = 0
            }
          }

          "INTF" {
            $NodeName = $xmlNode.CLSNAME
            If ($NodeName.Substring(0, 4).ToUpper() -eq $PrefixOld + "IF_") {
              $NewNodeName = $PrefixNew + "IF_" + 
                $NodeName.Substring($PrefixOld.Length + 3, 
                $NodeName.Length - ($PrefixOld.Length + 3))
            }
            $NewNodeName = Correct-NodeName $NewNodeName
            If ($NewNodeName.Length -gt 30) {
              $Signal = 1
            }
            Else {
              $Signal = 0
            }
          }

          "MSAG" {
            $NodeName = $xmlNode.ARBGB
            $NewNodeName = $PrefixNew + $NodeName.Substring($PrefixOld.Length,
              $NodeName.Length - $PrefixOld.Length)
            $NewNodeName = Correct-NodeName $NewNodeName
            If ($NewNodeName.Length -gt 20) {
              $Signal = 1
            }
            Else {
              $Signal = 0
            }
          }
          
          "PROG" {
            $NodeName = $xmlNode.NAME
            $NewNodeName = Correct-NodeName $NodeName
            If ($NewNodeName.Length -gt 30) {
              $Signal = 1
            }
            Else {
              $Signal = 0
            }
          }

          "TABL" {
            $NodeName = $xmlNode.TABNAME
            $NewNodeName = $PrefixNew + $NodeName.Substring($PrefixOld.Length, 
              $NodeName.Length - $PrefixOld.Length)
            $NewNodeName = Correct-NodeName $NewNodeName
            Switch ($xmlNode.TABCLASS.ToUpper()) {
              "INTTAB" {
                $ClassName = "INTTAB"
                If ($NewNodeName.Length -gt 30) {
                  $Signal = 1
                }
                Else {
                  $Signal = 0
                }
              }
              "TRANSP" {
                $ClassName = "TRANSP"
                If ($NewNodeName.Length -gt 16) {
                  $Signal = 1
                }
                Else {
                  $Signal = 0
                }
              }
            }
          }

          "TTYP" {
            $NodeName = $xmlNode.TYPENAME
            $NewNodeName = $PrefixNew + $NodeName.Substring($PrefixOld.Length,
              $NodeName.Length - $PrefixOld.Length)
            $NewNodeName = Correct-NodeName $NewNodeName
            If ($NewNodeName.Length -gt 30) {
              $Signal = 1
            }
            Else {
              $Signal = 0
            }
          }

          "XSLT" {
            $NodeName = $xmlNode.XSLTDESC
            $NewNodeName = $PrefixNew + $NodeName.Substring($PrefixOld.Length,
              $NodeName.Length - $PrefixOld.Length)
            $NewNodeName = Correct-NodeName $NewNodeName
            If ($NewNodeName.Length -gt 30) {
              $Signal = 1
            }
            Else {
              $Signal = 0
            }
          }

        }

        $Lines += [String]($Lines.Length) + ";" + $xmlNode.LocalName.ToUpper() +
          ";" + $NodeName + ";" + $ClassName + ";" + $NewNodeName +
          ";" + $NewNodeName.Length.ToString() + ";" + $Signal

      }

      [String[]]$Text = $Lines[0]
      For ($i = $Lines.Length - 1; $i -ge 1; $i--) {
        $Text += $Lines[$i]
      }

      $FileName = $FileName + ".csv"
      $Text | Out-File $FileName -Encoding ascii

    }

  #-Main----------------------------------------------------------------
    Main

#-End-------------------------------------------------------------------
