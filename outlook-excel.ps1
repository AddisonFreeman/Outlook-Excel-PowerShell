#PowerShell v2.0
#An application to parse an Outlook inbox for specific attachments
#	and save their data in a predefined excel spreadsheet

$excelPath = "C:\Users\addis\GitHub\Outlook-Excel-PowerShell\end\dest.xlsx"
$tempDirectory = "C:\Users\addis\GitHub\Outlook-Excel-PowerShell\end" #dont include an ending \ at the end of that line
$subjectTitle = "0123"
$attachment = "attachment"
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

$inbox.items | foreach {
	If ( ($_.subject -match $subjectTitle) ) {
		$_.attachments | foreach {
			If ($_.FileName -match $attachment) {
				$tempAttachFile = "$tempDirectory\attach$i.txt"
				$_.SaveAsFile($tempAttachFile)
				$contents = Get-Content $tempAttachFile
				#Write Each line, $j, from attached file to column $i of excel spreadsheet $sheetNum
				$j = 1
				#if !ColumnNamesFlag, write left of each line to first column and the data to the second, only take data after that
				$contents | ForEach-Object { $ws.Cells.item($j,$i) = $_.Split(":")[-1]; $j++;  }
				Write-Host $j, $i, $sheetNum
				$i++
				Remove-Item $tempAttachFile
			}
		}
	}
}
$workbook.SaveAs($excelPath)
$workbook.Close($true)
$excel.quit()
