# Перевірка, чи запущено скрипт від адміна
function TestAdminRights {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    return $isAdmin
}

if (-not (TestAdminRights)) {
    Write-Host ">> This script must be run as admin." -ForegroundColor Red
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
$cmsInstallerURL = "https://github.com/aspektyoyo/pk/raw/main/Setup.exe"
$oldCMSPath = "C:\Program Files (x86)\CMS"
$polyvisionCMSPath = "C:\Program Files (x86)\Polyvision\CMS"
$tempDir = "D:\Temp"





# Етап 1. Перевірка наявності старого CMS
#   Тут ми перевіряємо, чи існує папка старого CMS, чи існує CMS.exe та чи є необхідні нам 4 файли XML
#   Крім того, ми створюємо тимчасову папку, в яку будемо скидувати все нам необхідне

Start-Sleep -Milliseconds 1000
Write-Host ""
Write-Host ">> Starting old CMS check proccess" -BackgroundColor DarkMagenta -ForegroundColor White

# Перевіряємо наявність тимчасової папки, і якщо її нема - створюємо
if (-not (Test-Path $tempDir -PathType Container)) {
    New-Item -Path $tempDir -ItemType Directory
    Write-Host ">> $tempDir created" -ForegroundColor Green
} else {
    Write-Host ">> $tempDir is already exist" -ForegroundColor Yellow
}

# Визначаємо булеві змінні
$isOldCMSInstalled = $false
$dataExist = $false
$devGroupExist = $false
$planTemplateExist = $false
$usersExist = $false

# Перевіряємо наявність папки CMS і, якщо вона існує, необхідних файлів
if (Test-Path $oldCMSPath -PathType Container) {
    Write-Host ">> CMS directory was found" -ForegroundColor Green

    # Перевіряємо наявність екзешніка
    $isOldCMSInstalled = CheckFile $oldCMSPath "CMS.exe"

    # Перевіряємо наявність папки XML і необхідних нам файлів
    if (Test-Path (Join-Path $oldCMSPath "XML\")) {
        Write-Host ">> XML directory was found" -ForegroundColor Green

        $dataExist = CheckFile $oldCMSPath "XML\Data.xml"
        $devGroupExist = CheckFile $oldCMSPath "XML\DevGroup.xml"
        $planTemplateExist = CheckFile $oldCMSPath "XML\PlanTemplate.xml"
        $usersExist = CheckFile $oldCMSPath "XML\users.xml"
    } else {
        Write-Host ">> CMS/XML directory not found" -ForegroundColor Red
    }
} else {
    Write-Host ">> CMS directory not found" -ForegroundColor Red
}

Write-Host ">> Old CMS check finished." -BackgroundColor DarkGreen -ForegroundColor White





# Етап 2. Перевіряємо, чи встановлено новий Polyvision/CMS

Start-Sleep -Milliseconds 1000
Write-Host ""
Write-Host ">> Starting Polyvision/CMS check proccess" -BackgroundColor DarkMagenta -ForegroundColor White

# Визначаємо булеві змінні
$isPolyvisionCMSInstalled = $false
$polyvisionDataExist = $false
$polyvisionDevGroupExist = $false
$polyvisionPlanTemplateExist = $false
$polyvisionUsersExist = $false

# Перевіряємо наявність папки Polyvision/CMS і, якщо вона існує, необхідних файлів
if (Test-Path $polyvisionCMSPath -PathType Container) {
    Write-Host ">> Polyvision/CMS directory was found" -ForegroundColor Green

    # Перевіряємо наявність екзешніка
    $isPolyvisionCMSInstalled = CheckFile $polyvisionCMSPath "CMS.exe"

    # Перевіряємо наявність папки XML і необхідних нам файлів
    if (Test-Path (Join-Path $polyvisionCMSPath "XML\")) {
        Write-Host ">> XML directory was found" -ForegroundColor Green

        $polyvisionDataExist = CheckFile $polyvisionCMSPath "XML\Data.xml"
        $polyvisionDevGroupExist = CheckFile $polyvisionCMSPath "XML\DevGroup.xml"
        $polyvisionPlanTemplateExist = CheckFile $polyvisionCMSPath "XML\PlanTemplate.xml"
        $polyvisionUsersExist = CheckFile $polyvisionCMSPath "XML\users.xml"
    } else {
        Write-Host ">> Polyvision/CMS/XML directory not found" -ForegroundColor Red
    }
} else {
    Write-Host ">> Polyvision/CMS directory not found" -ForegroundColor Red
}

Write-Host ">> Polyvision/CMS check finished." -BackgroundColor DarkGreen -ForegroundColor White






# Етап 3. Копіювання XML файлів та завантаження інсталятора Polyvision/CMS
#   Тут ми копіюємо XML файли, якщо вони є, і завантажуємо інсталятор в тимчасову папку

if ($isOldCMSInstalled) {
    Start-Sleep -Milliseconds 1000
    Write-Host ""
    Write-Host ">> Starting XML files copying proccess" -BackgroundColor DarkMagenta -ForegroundColor White

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
    } else {
        Write-Host ">> Some (or all) XML files are missed. Copying proccess skipped." -BackgroundColor DarkYellow -ForegroundColor White
    }
} else {
    Write-Host ""
    Write-Host ">> Old CMS is not installed. Copying proccess skipped." -BackgroundColor DarkYellow -ForegroundColor White
}

if (-not $isPolyvisionCMSInstalled) {
    Write-Host ""
    Write-Host ">> Starting Polyvision/CMS download proccess" -BackgroundColor DarkMagenta -ForegroundColor White
    
    try {
        # Загрузка файла
        Write-Host ">> Downloading Polyvision/CMS installer..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $cmsInstallerURL -OutFile (Join-Path $tempDir "\setup.exe")
        Write-Host ">> Polyvision/CMS downloaded to $tempDir" -ForegroundColor Green
    } catch {
        Write-Host ">> Error downloading Polyvision/CMS installer. Aborting script." -ForegroundColor Red
    }
    
    Write-Host ">> Polyvision/CMS download finished." -BackgroundColor DarkGreen -ForegroundColor White
} else {
    Write-Host ""
    Write-Host ">> Polyvision/CMS is already installed. Download skipped." -BackgroundColor DarkYellow -ForegroundColor White
}
