Add-Type -AssemblyName System.Windows.Forms

$ascii = @"
 ________     ___    ___ ________  ________   ________  ________     
|\   __  \   |\  \  /  /|\   __  \|\   ___  \|\   __  \|\   __  \    
\ \  \|\  \  \ \  \/  / | \  \|\  \ \  \\ \  \ \  \|\  \ \  \|\  \   
 \ \   __  \  \ \    / / \ \   __  \ \  \\ \  \ \   __  \ \   _  _\  
  \ \  \ \  \  /     \/   \ \  \ \  \ \  \\ \  \ \  \ \  \ \  \\  \| 
   \ \__\ \__\/  /\   \    \ \__\ \__\ \__\\ \__\ \__\ \__\ \__\\ _\ 
    \|__|\|__/__/ /\ __\    \|__|\|__|\|__| \|__|\|__|\|__|\|__|\|__|
             |__|/ \|__|                                             
                                                                     
                                        PRESENT

  ____                           _     _               _            __ 
 / ___| _ __  _ __ ___  __ _  __| |___| |__   ___  ___| |_    ___  / _|
 \___ \| '_ \| '__/ _ \/ _` |/ _` / __| '_ \ / _ \/ _ \ __|  / _ \| |_ 
  ___) | |_) | | |  __/ (_| | (_| \__ \ | | |  __/  __/ |_  | (_) |  _|
 |____/| .__/|_|  \___|\__,_|\__,_|___/_| |_|\___|\___|\__|  \___/|_|  
       |_|                                                             
"@

Write-Host $ascii -ForegroundColor Green

$ascii = @"
  ____   ____ ____            ____ _____  ___   _   _____ _____    _    __  __ 
 |  _ \ / ___/ ___|          |  _ \_ _\ \/ / \ | | |_   _| ____|  / \  |  \/  |
 | |_) | |  _\___ \   _____  | |_) | | \  /|  \| |   | | |  _|   / _ \ | |\/| |
 |  _ <| |_| |___) | |_____| |  __/| | /  \| |\  |   | | | |___ / ___ \| |  | |
 |_| \_\\____|____/          |_|  |___/_/\_\_| \_|   |_| |_____/_/   \_\_|  |_|
                                                                               
"@

Write-Host $ascii -ForegroundColor Red

Write-Host "Script for download created by MIKCPU / Script per il download creato da MIKCPU" -ForegroundColor White

$FileId = "1yeir31wUMQXb9cWuI6mkrZiZPdWvGy_T"
$DownloadUrl = "https://drive.google.com/uc?export=download&id=$FileId"

$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Seleziona la cartella dove salvare il file / Select folder to save the file"
$folderBrowser.ShowNewFolderButton = $true

if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $selectedPath = $folderBrowser.SelectedPath
    $LocalFile = Join-Path -Path $selectedPath -ChildPath "RGS Downloads.ods"
    $TempFile = "$LocalFile.new"

    Invoke-WebRequest -Uri $DownloadUrl -OutFile $TempFile -UseBasicParsing

    if (-Not (Test-Path $LocalFile)) {
        Move-Item -Path $TempFile -Destination $LocalFile
    }
    else {
        $originalSize = (Get-Item $LocalFile).Length
        $newSize = (Get-Item $TempFile).Length

        if ($originalSize -ne $newSize) {
            Move-Item -Path $TempFile -Destination $LocalFile -Force
        }
        else {
            Remove-Item $TempFile
        }
    }

    Write-Host "Download completed successfully! / Download completato con successo!" -ForegroundColor Cyan
    Write-Host "Press any key to open the file... / Premi un tasto qualsiasi per aprire il file..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)

    Start-Process $LocalFile
}
else {
    Write-Host "Operation cancelled by user. / Operazione annullata dall'utente." -ForegroundColor Red
}
