$ErrorActionPreference = "Stop"
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $filePath = (Get-Item "Telecom_Job_Preparation_Guide_Book.docx").FullName
    $doc = $word.Documents.Open($filePath, [Type]::Missing, $false)
    $doc.Repaginate()
    $doc.Fields.Update()
    
    $pages = $doc.ComputeStatistics(2)
    $words = $doc.ComputeStatistics(0)
    Write-Output "TOTAL_PAGES:$pages"
    Write-Output "TOTAL_WORDS:$words"
    
    $results = @{}
    foreach ($b in $doc.Bookmarks) {
        $name = $b.Name
        $page = $b.Range.Information(3) # 3 = wdActiveEndPageNumber
        $results[$name] = $page
    }
    
    # Save the updated fields
    $doc.Save()
    $doc.Close($false)
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    
    $results.GetEnumerator() | Sort-Object Value | ForEach-Object {
        Write-Output "$($_.Key):$($_.Value)"
    }
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
