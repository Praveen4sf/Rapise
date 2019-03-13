$filePath = https://github.com/Praveen4sf/Rapise/RapiseDeployment/PkgFolder/Praveen/Rapise_Config.xlsx

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true

if (test-path $filePath) #...Reading data from Excel
{
	$wb = $xl.Workbooks.Open($filePath)
	$ws = $xl.WorkSheets.Item("sheet1")
	$countVal = $ws.usedrange.rows.count
	
	for($i=2;$i -le $countVal;$i++)
    {
	    $aVal = ""
	    $bVal = ""
	    $cVal = ""
	    $dVal = ""
		
		#...To make sure the Rapise Engine path and Rapise Test File path in the Input file
		
	    if($ws.Range("A"+$i).value2 -ne $Null)
	    {
		    $aVal =$ws.Range("A"+$i).value2.Trim()
			
		    if($ws.Range("B"+$i).value2 -ne $Null)
		    {
			    $bVal =$ws.Range("B"+$i).value2.Trim()

			    if(($aVal  -ne "") -or ($bVal  -ne ""))
			    {
				    
				    write-host $aVal

				    write-host $bVal

				    $cVal = $ws.Range("C"+$i).value2
				    if($cVal -ne $Null)
				    {
					    $cVal = $cVal.Trim()
				    }
				    write-host $cVal

				    $dVal = $ws.Range("D"+$i).value2
				    if($dVal -ne $Null)
				    {
					    $dVal = $dVal.Trim()
				    }
				    write-host $dVal
					

					#...Fetching the path of Test Script folder
					
					$path = Split-Path $bVal
					#$bVal.Substring(0,$bVal.LastIndexOf("\"))
					
					write-host "***Path="$path

			    }else
                {
                 Write-Host "Rapise Engine Path/Rapise Test File Path is Not Present in the Record No: "($i-1) " of the File: "$filePath
                }
		    }else
            {
                 Write-Host "Rapise Test File Path is Not Present in the Record No: "($i-1) " of the File: "$filePath
            }
	    }else
        {
             Write-Host "Rapise Engine Path is Not Present in the Record No: "($i-1) " of the File: "$filePath
        }
    }
	  
	
}else 
{
    write-host "FilePath: " $filePath + " NOT found"
}
$wb.close()
$xl.quit()

