
#-Begin-----------------------------------------------------------------

  #-Load assemblies-----------------------------------------------------
    [Reflection.Assembly]::LoadWithPartialName("'Microsoft.VisualBasic") > $Null

  #-Function Correct-ProgName-------------------------------------------
    Function Correct-ProgName ([String] $NodeName) {

      If ($NodeName.Length -ge 12) {
        If ($NodeName.Substring(0, 12) -eq "ZDEMO_TECHED") {
          If ([Int]$NodeName.Substring(12, $NodeName.Length - 12) -lt 10) {
            $NewNodeName = "/GKV/CA03DT0" + $NodeName.Substring(12, $NodeName.Length - 12)
          }
          Else {
            $NewNodeName = "/GKV/CA03DT" + $NodeName.Substring(12, $NodeName.Length - 12)
          }
        }
      }

      If ($NodeName.Length -ge 11) {
        If ($NodeName.Substring(0, 11) -eq "ZDEMO_EXCEL") {
          $IsNum = [Microsoft.VisualBasic.Information]::IsNumeric($NodeName.Substring(11, $NodeName.Length - 11))
          If ($IsNum -eq $True) {
            If ([Int]$NodeName.Substring(11, $NodeName.Length - 11) -lt 10) {
              $NewNodeName = "/GKV/CA03DE0" + $NodeName.Substring(11, $NodeName.Length - 11)
            }
            Else {
              $NewNodeName = "/GKV/CA03DE" + $NodeName.Substring(11, $NodeName.Length - 11)
            }
          }
        }
      }

      Switch ($NodeName) {
        "ZDEMO_EXCEL_OUTPUTOPT_INCL" {
          $NewNodeName = "/GKV/CA03_DEMO_OUTPUTOPT_INCL"
        }
        "ZDEMO_EXCEL" {
          $NewNodeName = "/GKV/CA03DE00"
        }
        "ZDEMO_CALENDAR_CLASSES" {
          $NewNodeName = "/GKV/CA03_DEMO_CALENDAR_CLAZES"
        }
        "ZDEMO_CALENDAR" {
          $NewNodeName = "/GKV/CA03DCAL"
        }
        "ZANGRY_BIRDS" {
          $NewNodeName = "/GKV/CA03ANBI"
        }
        "ZABAP2XLSX_DEMO_SHOW" {
          $NewNodeName = "/GKV/CA03ABDS"
        }
      }

      Return $NewNodeName
    
    }

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
      $PrefixNew = "/GKV/CA03_"
      $FileName = "C:\Schnell\Pool\FV_Aerzte\Excel\Community3\ABAP2XLSX_V_7_0_5.nugg"

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
            $NewNodeName = Correct-ProgName $NodeName
            If ($NewNodeName.Length -gt 13) {
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
