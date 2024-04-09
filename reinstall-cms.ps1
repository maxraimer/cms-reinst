# Перевірка, чи запущено скрипт від адміна
function TestAdminRights {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    return $isAdmin
}

if (-not (TestAdminRights)) {
    Write-Host ">> This script must be run as admin." -BackgroundColor Red -ForegroundColor White
    Pause
}

# Функція конвертації кирилиці в юнікод
function ConvertTo-Encoding ([string]$From, [string]$To){
    Begin{
        $encFrom = [System.Text.Encoding]::GetEncoding($from)
        $encTo = [System.Text.Encoding]::GetEncoding($to)
    }
    Process{
        $bytes = $encTo.GetBytes($_)
        $bytes = [System.Text.Encoding]::Convert($encFrom, $encTo, $bytes)
        $encTo.GetString($bytes)
    }
}

# Функція перевірки наявності файлів
function CheckFile($path, $fileName) {
    if (Test-Path (Join-Path $path $fileName)) {
        Write-Host ">> $fileName was found" -ForegroundColor Green
        return $true
    } else {
        Write-Host ">> $fileName not found" -ForegroundColor Red
        return $false
    }
}




# Посилання на інсталятор нового CMS та шлях до старого на локальному компі
$cmsInstallerURL = "https://github.com/maxraimer/cms-reinst/raw/main/setup.exe"
$shortcutIconURL = "https://github.com/maxraimer/cms-reinst/raw/main/cms.ico"
$oldCMSPath = "C:\Program Files (x86)\CMS"
$polyvisionCMSPath = "C:\Program Files (x86)\Polyvision\CMS"
$tempDir = "D:\Temp"





Write-Host "                            " -BackgroundColor DarkBlue
Write-Host ">> Executing script started." -BackgroundColor DarkBlue -ForegroundColor White
Write-Host "                            " -BackgroundColor DarkBlue





# Етап 1. Перевірка наявності старого CMS
#   Тут ми перевіряємо, чи існує папка старого CMS, чи існує CMS.exe та чи є необхідні нам 4 файли XML
#   Крім того, ми створюємо тимчасову папку, в яку будемо скидувати все нам необхідне

Start-Sleep -Milliseconds 1000
Write-Host ""
Write-Host ">> Starting old CMS check proccess..." -BackgroundColor DarkBlue -ForegroundColor White

# Перевіряємо наявність тимчасової папки, і якщо її нема - створюємо
if (-not (Test-Path $tempDir -PathType Container)) {
    mkdir $tempDir | Out-Null
    Write-Host ">> $tempDir created." -ForegroundColor Green
} else {
    Write-Host ">> $tempDir is already exist." -ForegroundColor Yellow
}

# Визначаємо булеві змінні
$isOldCMSInstalled = $false
$dataExist = $false
$devGroupExist = $false
$planTemplateExist = $false
$usersExist = $false
$xmlCopiedToTemp = $false

# Перевіряємо наявність папки CMS і, якщо вона існує, необхідних файлів
if (Test-Path $oldCMSPath -PathType Container) {
    Write-Host ">> CMS directory was found." -ForegroundColor Green

    # Перевіряємо наявність екзешніка
    $isOldCMSInstalled = CheckFile $oldCMSPath "CMS.exe"

    # Перевіряємо наявність папки XML і необхідних нам файлів
    if (Test-Path (Join-Path $oldCMSPath "XML\")) {
        Write-Host ">> XML directory was found." -ForegroundColor Green

        $dataExist = CheckFile $oldCMSPath "XML\Data.xml"
        $devGroupExist = CheckFile $oldCMSPath "XML\DevGroup.xml"
        $planTemplateExist = CheckFile $oldCMSPath "XML\PlanTemplate.xml"
        $usersExist = CheckFile $oldCMSPath "XML\users.xml"
    } else {
        Write-Host ">> CMS/XML directory not found." -ForegroundColor Red
    }
} else {
    Write-Host ">> CMS directory not found." -ForegroundColor Red
}

Write-Host ">> Old CMS check finished." -BackgroundColor DarkGreen -ForegroundColor White





# Етап 2. Перевіряємо, чи встановлено новий Polyvision/CMS

Start-Sleep -Milliseconds 1000
Write-Host ""
Write-Host ">> Starting Polyvision/CMS check proccess..." -BackgroundColor DarkBlue -ForegroundColor White

# Визначаємо булеві змінні
$isPolyvisionCMSInstalled = $false

# Перевіряємо наявність папки Polyvision/CMS і, якщо вона існує, необхідних файлів
if (Test-Path $polyvisionCMSPath -PathType Container) {
    Write-Host ">> Polyvision/CMS directory was found." -ForegroundColor Green

    # Перевіряємо наявність екзешніка
    $isPolyvisionCMSInstalled = CheckFile $polyvisionCMSPath "CMS.exe"
} else {
    Write-Host ">> Polyvision/CMS directory not found." -ForegroundColor Red
}

Write-Host ">> Polyvision/CMS check finished." -BackgroundColor DarkGreen -ForegroundColor White





# Етап 3. Копіювання XML файлів та завантаження інсталятора Polyvision/CMS
#   Тут ми копіюємо XML файли, якщо вони є, і завантажуємо інсталятор в тимчасову папку

if ($isOldCMSInstalled) {
    Start-Sleep -Milliseconds 1000
    Write-Host ""
    Write-Host ">> Starting XML files copying proccess..." -BackgroundColor DarkBlue -ForegroundColor White

    # Копіюємо XML файли в тимчасову папку, якщо вони існують
    if ($dataExist -and $devGroupExist -and $planTemplateExist -and $usersExist) {

        $xmlFiles = Get-ChildItem -Path (Join-Path $oldCMSPath "XML") -File
        $totalFiles = $xmlFiles.Count
        $copiedFiles = 0

        foreach ($file in $xmlFiles) {
            $copiedFiles++
            Write-Host ">> Copying file $copiedFiles of $($totalFiles): $($file.Name)" -ForegroundColor Yellow
            Copy-Item -Path $file.FullName -Destination $tempDir -Force
            Start-Sleep -Milliseconds 500
        }

        Write-Host ">> Copying finished." -BackgroundColor DarkGreen -ForegroundColor White
        $xmlCopiedToTemp = $true
    } else {
        Write-Host ">> Some (or all) XML files are missed. Copying proccess skipped." -BackgroundColor DarkYellow -ForegroundColor White
    }
} else {
    Start-Sleep -Milliseconds 1000
    Write-Host ""
    Write-Host ">> Old CMS is not installed. Copying proccess skipped." -BackgroundColor DarkBlue -ForegroundColor White
}

if (-not $isPolyvisionCMSInstalled) {
    Write-Host ""
    Write-Host ">> Starting Polyvision/CMS download proccess..." -BackgroundColor DarkBlue -ForegroundColor White
    
    if (-not (CheckFile $tempDir "setup.exe")) {
        try {
            # Загрузка файла
            Write-Host ">> Downloading Polyvision/CMS installer..." -ForegroundColor Yellow
            Invoke-WebRequest -Uri $cmsInstallerURL -OutFile (Join-Path $tempDir "\setup.exe")
            Write-Host ">> Polyvision/CMS downloaded to $tempDir" -ForegroundColor Green
        } catch {
            Write-Host ">> Error downloading Polyvision/CMS installer. Aborting script." -ForegroundColor Red
            Pause
        }

        Write-Host ">> Polyvision/CMS installer download finished." -BackgroundColor DarkGreen -ForegroundColor White
    } else {
        Write-Host ">> Polyvision/CMS installer is already donwloaded." -BackgroundColor DarkGreen -ForegroundColor White
    }    
} else {
    Start-Sleep -Milliseconds 1000
    Write-Host ""
    Write-Host ">> Polyvision/CMS is already installed. Download skipped." -BackgroundColor DarkBlue -ForegroundColor White
}





# Етап 4. Встановлення Polyvision/CMS та видалення старого CMS
#   В цьому етапі, якщо Polyvision/CMS не встановлено, ми його встановлюємо
#   Крім того, ми копіюємо збережені файли XML до нового CMS 

if (-not $isPolyvisionCMSInstalled) {
    Start-Sleep -Milliseconds 1000
    Write-Host ""
    Write-Host ">> Starting Polyvision/CMS installation proccess..." -BackgroundColor DarkBlue -ForegroundColor White
    
    $setupPath = Join-Path $tempDir "\setup.exe"
    $arguments = "/SILENT"

    Start-Process -FilePath $setupPath -ArgumentList $arguments -Wait
    Write-Host ">> Polyvision/CMS installed." -ForegroundColor Green
    
    if ($xmlCopiedToTemp) {
        $xmlFiles = Get-ChildItem -Path $tempDir -Filter "*.xml" -File
        $totalFiles = $xmlFiles.Count
        $copiedFiles = 0

        foreach ($file in $xmlFiles) {
            $copiedFiles++
            Write-Host ">> Copying file $copiedFiles of $($totalFiles): $($file.Name)" -ForegroundColor Yellow
            Copy-Item -Path $file.FullName -Destination (Join-Path $polyvisionCMSPath "\XML") -Force
            Start-Sleep -Milliseconds 500
        }
        Write-Host ">> XML files copying finished." -ForegroundColor Green
        Write-Host ">> Polyvision/CMS is installed and configured." -BackgroundColor DarkGreen -ForegroundColor White
    } else {
        Write-Host ""
        Write-Host ">> No XML files found. You will need to set up Polyvision/CMS manually." -BackgroundColor DarkYellow -ForegroundColor White
    }
} else {
    Start-Sleep -Milliseconds 1000
    Write-Host ""
    Write-Host ">> Polyvision/CMS is already installed. Installation skipped." -BackgroundColor DarkBlue -ForegroundColor White
}





# Етап 5. Створення Bat-файлу та ярлика на робочому столі
#   В цьому етапі ми створюємо Bat-файл для запуску CMS без введення паролю та створюємо ярлик на робочому столі

Start-Sleep -Milliseconds 1000
Write-Host ""
Write-Host ">> Starting BAT-file creation proccess..." -BackgroundColor DarkBlue -ForegroundColor White

$batFilePath = Join-Path $polyvisionCMSPath "\CMS.bat"
if (Test-Path $batFilePath -PathType Leaf) {
    Write-Host ">> BAT-file is already exist. Going on..."  -ForegroundColor Yellow
} else {
    $code = 'cmd /min /C "set __COMPAT_LAYER=RUNASINVOKER && start "" "' + $polyvisionCMSPath + '\CMS.exe""'
    # Створюємо файл
    Set-Content -Path $batFilePath -Value $code
    Write-Host ">> BAT-file created." -ForegroundColor Green
}

Write-Host ">> Creating shortcut..."  -ForegroundColor Yellow
# Створення ярлика
function CreateShortcut {
    param (
        [string]$TargetPath,
        [string]$ShortcutPath,
        [string]$IconPath
    )

    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $TargetPath
    $Shortcut.IconLocation = $IconPath
    $Shortcut.Save()
}

# Шлях для збереження ярлика на робочому столі
$shortcutPath = "C:\Users\kassir\Desktop\КАМЕРЫ.lnk" | ConvertTo-Encoding "UTF-8" "windows-1251"

try {
    Write-Host ">> Downloading icon..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $shortcutIconURL -OutFile (Join-Path $tempDir "\cms.ico")
    Write-Host ">> Icon downloaded to $tempDir" -ForegroundColor Green
} catch {
    Write-Host ">> Error downloading icon. Aborting script." -ForegroundColor Red
    Pause
}

# Створення ярлика

try {
    CreateShortcut -TargetPath $batFilePath -ShortcutPath $shortcutPath -IconPath (Join-Path $tempDir "\cms.ico")
    Write-Host ">> Shortcut created on the Desktop." -ForegroundColor Green
} catch {
    Write-Host ">> Error creating shortcut." -ForegroundColor Red
}
Write-Host ">> BAT-file and shortcut created." -BackgroundColor DarkGreen -ForegroundColor White





# Етап 6. "Прибирання за собою" та видалення старого CMS 
#   В цьому етапі ми видаляємо старий CMS, видаляємо папку Temp

Start-Sleep -Milliseconds 1000
Write-Host ""
Write-Host ">> Starting cleaning proccess..." -BackgroundColor DarkBlue -ForegroundColor White

if ($isOldCMSInstalled) {
    Write-Host ">> Removing old CMS..." -ForegroundColor Yellow
    Remove-Item -Path $oldCMSPath -Recurse
    Write-Host ">> Old CMS removed." -ForegroundColor Green
} else {
    Write-Host ">> No old CMS. Skipping removing." -ForegroundColor Green
}

Write-Host ">> Removing temporary directory..." -ForegroundColor Yellow
Remove-Item -Path $tempDir -Recurse
Write-Host ">> Temporary directory removed." -ForegroundColor Green

Write-Host ">> Cleaning done." -BackgroundColor DarkGreen -ForegroundColor White

Start-Sleep -Milliseconds 1000
Write-Host ""
Write-Host "                                     " -BackgroundColor DarkGreen
Write-Host ">> The script has finished executing." -BackgroundColor DarkGreen -ForegroundColor White
Write-Host "                                     " -BackgroundColor DarkGreen