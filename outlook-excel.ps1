#PowerShell v2.0
#An application to parse an Outlook inbox for specific attachments
#	and save their data in a predefined excel spreadsheet
#
#	now?: connect to inbox, count number of emails

#TODO make below paths optional parameters
$srcPath = "\src\attachment-12-34.txt"
#$emailSubjectMatch = '^TestSubjectName[0-9]*$'
#$attachmentMatch = [regex] '^attachment-[0-9][0-9]-[0-9][0-9].txt$'
$excelPath = "\end\dest.xlsx"
$excelPage = "Sheet1"

Add-Type -assembly "Microsoft.Office.Interop.Outlook"
$o = New-Object -comobject outlook.application
$n = $o.GetNamespace("MAPI")
$inbox = $n.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

#$excel = New-Object -Com Excel.Application
#$workbook = $Excel.Workbooks.Open($excelPath)
#$ws = $Workbook.worksheets | where-object {$_.Name -eq $excelPage}

$inbox.items | foreach {
	If ($_.subject -match '0123') {
		$_.attachments | foreach {
			If ($_.FileName -match 'attachment') {
				Write-Host $_.FileName
			}
		}
	}
}
#$n.Folders.Item('test.addisonfreeman@gmail.com').Folders.Item('Inbox')
 
 
 #Below: testing writing to an excel file
# $inbox.items | foreach {
#	If ($_.Subject -match "attachment*") {
#	$_attachments | foreach {
		
		#$content = Get-Content ($_)[4..6]
		#Write-Host "ok, good"
#	}
#}
			#$ws.Cells.item(3,1) = $content[0] #Row 3, Col 1
#			$ws.Cells.item(4,1) = $content[1] #Row 4, Col 1
#			$ws.Cells.item(5,1) = $content[2] #Row 5, Col 1
#			$workbook.Close($true)
#			$excel.quit()
#		}
#	}
#}
