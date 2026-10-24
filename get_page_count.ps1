$ErrorActionPreference = "Stop"
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $filePath = (Get-Item "Telecom_Job_Preparation_Guide_Book.docx").FullName
    $doc = $word.Documents.Open($filePath, [Type]::Missing, $true)
    
    # Repaginate to ensure accurate page count
    $doc.Repaginate()
    
    # wdStatisticPages = 2, wdStatisticWords = 0, wdStatisticCharacters = 3, wdStatisticParagraphs = 4
    $pages = $doc.ComputeStatistics(2)
    $words = $doc.ComputeStatistics(0)
    $chars = $doc.ComputeStatistics(3)
    $paragraphs = $doc.ComputeStatistics(4)
    
    Write-Output "PAGES:$pages"
    Write-Output "WORDS:$words"
    Write-Output "CHARS:$chars"
    Write-Output "PARAGRAPHS:$paragraphs"
    
    $doc.Close($false)
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
