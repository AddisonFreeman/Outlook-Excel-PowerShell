#PowerShell v2.0
#An application to parse an Outlook inbox for specific attachments
#	and save their data in a predefined excel spreadsheet

$excelPath = "C:\Users\addis\GitHub\Outlook-Excel-PowerShell\end\dest.xlsx"
$tempDirectory = "C:\Users\addis\GitHub\Outlook-Excel-PowerShell\end" #dont include an ending \ at the end of that line
$subjectTitle = "0123"
$attachmentName = "attachment"
$i = 1
#$columnNamesFlag = False

Add-Type -assembly "Microsoft.Office.Interop.Outlook"
$o = New-Object -comobject outlook.application
$n = $o.GetNamespace("MAPI")
$inbox = $n.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

$excel = New-Object -Com Excel.Application
$workbook = $Excel.Workbooks.Open($excelPath)
$excelPage = "Sheet1"
$ws = $Workbook.worksheets | where-object {$_.Name -eq $excelPage}

foreach ($item in $inbox.items) {
	If ( ($item.subject -match $subjectTitle) ) {
		foreach ($attachment in $item.attachments) {
			If ($attachment.FileName -match $attachmentName) {
				$tempAttachFile = "$tempDirectory\attach$i.txt"
				$attachment.SaveAsFile($tempAttachFile)
				$contents = Get-Content $tempAttachFile
				$j = 1
				$contents | ForEach-Object { 
					#if the string is date or time, write string, else check for double at end of line
					if ( $_ -match "Date") {
						$ws.Cells.item($j,$i) = $_.Split(":")[-1]
						$j++
					} else {
						if ($_ -match "Time") {
							$ws.Cells.item($j,$i) = Get-Date ($_.Split(":")[1] + " :" + $_.Split(":")[-1]) -format t
							$j++
						} else {
							[double]$lineDouble = $null
							[double]::TryParse($_.Split(":")[-1], [ref]$lineDouble) 
							if ($lineDouble -is [Double]) {
								if($lineDouble -eq 0) {
									$ws.Cells.item($j,$i) = $lineDouble; 
									$j++;
								} else {
									$ws.Cells.item($j,$i) = $lineDouble.ToString("#.##");
									$j++;
								}
							}
						}						
					}
				}
				Write-Host $j, $i
				$i++
				Remove-Item $tempAttachFile
			}
		}
	}
}
$workbook.SaveAs($excelPath)
$workbook.Close($true)
$excel.quit()
