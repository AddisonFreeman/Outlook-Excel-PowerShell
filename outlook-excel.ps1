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
				#Write Each line, $j, from attached file to column $i of excel spreadsheet $sheetNum
				$j = 1
				$ws.Cells.item($j,$i) = $item.SentOn
				$j++
				$contents | ForEach-Object { 
					[double]$lineDouble = $null
					[double]::TryParse($_.Split(":")[-1], [ref]$lineDouble)
					if ($lineDouble) {
						$ws.Cells.item($j,$i) = $lineDouble.ToString("#.##");; 
						$j++;
					} else {
						$ws.Cells.item($j,$i) = $_
						$j++
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
