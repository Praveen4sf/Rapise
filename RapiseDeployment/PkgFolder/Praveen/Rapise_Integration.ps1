#...Note: Change the File/Folder paths according to your Environment
#...Reading Rapise Engine Path and Test Script inputs from Excel File

$filePath = dir Rapise_Config.xlsx

$csvContents = @()  #...Object Creation for CSV file

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
<#
					If($path -ne "")
                    {
					    $text = 
					    'If /I "%Processor_Architecture%" NEQ "x86" (
					    %SystemRoot%\SysWoW64\WindowsPowerShell\v1.0\powershell.exe /C %0
					    goto :eof
					    )
					    pushd %~dp0
					    cscript "'+$aVal+'" "'+$bVal+'" "-eval:g_testSetParams={userName:'''+$cVal+''', password:'''+$dVal+'''};"
					    popd'
						
						$CMDFile = $bVal.Substring(($bVal.LastIndexof("\")+1),($bVal.LastIndexof(".")-($bVal.LastIndexof("\")+1)))
						$filename = $path+'\'+$CMDFile+'_Execute.cmd'
						
					    #$filename = $path+'\Execute.cmd'
					    $text | Set-Content $filename  #...Creating .cmd file for executing Test Scripts
                       
					    cmd /c $filename #...Executing .cmd file
                         
                        
					    #...Reading the output of Test (TRP file)

					    $TRPfile = dir $path *.trp | sort LastWriteTime | select -last 1 | select name
                        If ($TRPfile -ne $null)
                        {
                            $TRPfile = $TRPfile -split "="
					        $TRPfile = $TRPfile[1]
					        $TRPfile = $TRPfile -replace '}',''
					        write-host "***TRP file: "$TRPfile

					        $TRPfilePath = $path+"\"+$TRPfile
					        write-host "***TRP FilePath: "$TRPfilePath

					        $NL = "`r`n" #...To add new Line to the log variable
					        $Reason = ""
					        $Status = ""
					        $Summary = ""
                            $LWT = dir $path *.trp | sort LastWriteTime | select -last 1 
                            $LWT = $LWT.lastwritetime

					        $log = "File Name = "+$TRPfile
                            $log=$log + $NL + "Last Write Time = "+$LWT

                            $text1 =""
                            $text2 =""

					        #...Searching for 'Fail' word in TRP file for to know the Result of Test
                            $text1 = Get-Content $TRPfilePath | Select-String -pattern "\bFail\b"

                            if($text1.length -ne 0)
                            {
	                            $Status = "FAIL"
						
	                            $array1 = $text1[0] -split """"
	                            $index1 = [array]::Indexof($array1, " name=")
	                            $Reason = $array1[$index1+1]
						
	                            $log=$log + $NL + "Status = "+$Status
	                            $log=$log +$NL+"Failure Reason = " + $Reason
                            }else
                            {
	                            #...Searching for 'Pass' word in TRP file for to know the Result of Test
	                            $text1 = Get-Content $TRPfilePath | Select-String -pattern "\bPass\b"

	                            if($text1.length -ne 0)
	                            {
		                            $Status = "PASS"
		                            $log=$log + $NL + "Status = "+$Status
	                            }else{
		                            $Status = "No Status"
		                            $log=$log + $NL + "Status = "+$Status
	                            }
                            }

                            #...Searching for 'Failed' word in TRP file for Summary
                            $text2 = Get-Content $TRPfilePath | Select-String -pattern "\bFailed\b"
                            if($text2.length -ne 0)
                            {
	                            $array2 = $text2[0] -split """"
	                            $index2 = [array]::Indexof($array2, " comment=")
	                            $Summary = $array2[$index2+1]
						
	                            $log=$log + $NL + "Summary = " + $Summary
                            }else
                            {
	                            $Summary = "No Summary for this file / Test Doesn't Executed"
	                            $log=$log + $NL + "Summary = " + $Summary
                            }
                            write-host $log

					
					        #...Appending log values to CSV file row 
					        $row = New-Object System.Object 
					        $row | Add-Member -MemberType NoteProperty -Name "File Name" -Value $TRPfile
                            $row | Add-Member -MemberType NoteProperty -Name "Last Write Time" -Value $LWT
					        $row | Add-Member -MemberType NoteProperty -Name "Status" -Value $Status 
					        $row | Add-Member -MemberType NoteProperty -Name "Failure Reason" -Value $Reason 
					        $row | Add-Member -MemberType NoteProperty -Name "Summary" -Value $Summary


					        $csvContents += $row 
					        $csvContents | Export-CSV -Path RapiseResult.csv -NoTypeInformation  -Force			
					    }else
                        {
                            Write-Host "TRP File Not Found at "$path
                        }#>
                    }else
                    {
                        Write-Host "Something Wrong with Test Script Path, Please check the Inputs files"
                    }
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

