$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open((Get-Item "test_toc_table.docx").FullName)
$doc.Repaginate()
$doc.Fields.Update()
Write-Output "DOC OPENED OK. Pages: $($doc.ComputeStatistics(2))"
Write-Output "Table cell 1,2: $($doc.Tables.Item(1).Cell(1,2).Range.Text.Trim())"
$doc.Close($false)
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
