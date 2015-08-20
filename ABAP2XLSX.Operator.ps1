
#-Begin-----------------------------------------------------------------
#-
#- Hint: Swap string $TMP to your package manually
#-
#-----------------------------------------------------------------------

  $XMLFileName = "ABAP2XLSX_V_7_0_5.nugg"
  $XMLFileNameNew = $XMLFileName + ".new"
  $CSVFileName = $XMLFileName + ".csv"
  $CSVFile = Import-Csv -Path $CSVFileName -Delimiter ";"
  [String]$XMLFile = (Get-Content -Path $XMLFileName) -Join "'r'n"
  $CSVFile | ForEach-Object {
    $OldName = $_.OldName; $NewName = $_.NewName
    Write-Host $OldName " > " $NewName
    if ($XMLFile.ToUpper().Contains("_" + $OldName.ToUpper()) -eq $True) {
      [String[]]$XML = $XMLFile -Split "'r'n"
      For($i = 0; $i -le $XML.Count - 1; $i++) {
        if ($XML[$i].ToUpper().Contains("_" + $OldName.ToUpper()) -eq $False) {
          $XML[$i] = $XML[$i] -ireplace $OldName, $NewName
        }
        else {
          Write-Host "Check line " $($i + 1)
        }
      }
      $XMLFile = $XML -Join "'r'n"
    }
    else {
      $XMLFile = $XMLFile -ireplace $OldName, $NewName
    }
  }
  Set-Content -Path $XMLFileNameNew -Value $XMLFile.Replace("'r'n", "$([char]0x0D)$([char]0x0A)")

#-End-------------------------------------------------------------------
