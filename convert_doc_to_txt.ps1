$docPath = Read-Host "Path to documents?"
$files = Get-ChildItem $docPath -filter "*.doc"
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatText")
$word = new-object -comobject word.application
#Time to process the files...
foreach ($file in $files) {

    #Write the output, change file name, and wrap up.

    Write-Output "Processing: $($file.FullName)"
    $doc = $word.Documents.Open($file.FullName)
    $fileName = $file.BaseName + '.txt'
    $doc.SaveAs("$docpath\$fileName",[ref]$saveFormat)
    Write-Output "File saved as $docPath\$fileName"
    $doc.Close()

}

$word.Quit()