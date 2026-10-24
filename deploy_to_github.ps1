# deploy_to_github.ps1
# Script to regenerate the web eBook and push updates to GitHub

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   টেলিকম জব প্রিপারেশন গাইড: ওয়েব ই-বুক সিঙ্ক ও ডেপ্লয়মেন্ট" -ForegroundColor Yellow
Write-Host "==========================================================" -ForegroundColor Cyan

# Step 1: Re-generate web eBook
Write-Host "`n[১/৩] ওয়েব সংস্করণ (docs/) রি-জেনারেট করা হচ্ছে..." -ForegroundColor Green
python -X utf8 export_to_web_ebook.py
if ($LASTEXITCODE -ne 0) {
    Write-Host "ত্রুটি: ওয়েব সংস্করণ তৈরি করতে সমস্যা হয়েছে!" -ForegroundColor Red
    exit 1
}

# Step 2: Git status and stage
Write-Host "`n[২/৩] গিট স্ট্যাটাস চেক এবং ফাইল স্টেজ করা হচ্ছে..." -ForegroundColor Green
if (-not (Test-Path ".git")) {
    Write-Host "গিট রিপোজিটরি ইনিশিয়ালাইজ করা হচ্ছে..." -ForegroundColor Cyan
    git init
    git branch -M main
}

git add .
$status = git status --porcelain
if (-not $status) {
    Write-Host "কোনো নতুন পরিবর্তন পাওয়া যায়নি। সবকিছু আপ-টু-ডেট আছে।" -ForegroundColor Yellow
} else {
    $commitMsg = "Update book content and interactive web eBook: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    git commit -m $commitMsg
    Write-Host "`nকমিট সম্পন্ন হয়েছে: $commitMsg" -ForegroundColor Green

    # Step 3: Git push
    Write-Host "`n[৩/৩] গিটহাবে পুশ করার চেষ্টা করা হচ্ছে..." -ForegroundColor Green
    $remotes = git remote
    if ($remotes -contains "origin") {
        git push origin main
        Write-Host "`nসফলভাবে গিটহাবে পুশ হয়েছে এবং GitHub Pages সিঙ্ক হয়েছে!" -ForegroundColor Cyan
        Write-Host "লাইভ ই-বুক লিঙ্ক: https://abmannan1761.github.io/Telecom-Job-Preparation-Guide/" -ForegroundColor Yellow
    } else {
        Write-Host "`nরিমোট রিপোজিটরি সেট করা নেই।" -ForegroundColor Yellow
        Write-Host "অনুগ্রহ করে একবার নিচের কমান্ডটি চালান:" -ForegroundColor White
        Write-Host "git remote add origin https://github.com/AbMannan1761/Telecom-Job-Preparation-Guide.git" -ForegroundColor Cyan
        Write-Host "git push -u origin main" -ForegroundColor Cyan
    }
}

Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host "সম্পন্ন হয়েছে!" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Cyan
