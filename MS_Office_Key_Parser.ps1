# Функция для вывода текста с заданным цветом
$url = "https://t.me/+ZipppRAXqghhYjUy"
$dev = "@prodby_expill"
function Write-LogoText {
    param (
        [string]$text,
        [string]$color
    )
    Write-Host $text
    Write-Host ""  # Перевод строки и сброс цвета
}

# Логотип в красном цвете
function Show-Logo {
    Write-Host @"


                         __    __     ______        ______     ______   ______   __     ______     ______        __  __     ______     __  __        ______   ______     ______     ______     ______     ______    
                        /\ "-./  \   /\  ___\      /\  __ \   /\  ___\ /\  ___\ /\ \   /\  ___\   /\  ___\      /\ \/ /    /\  ___\   /\ \_\ \      /\  == \ /\  __ \   /\  == \   /\  ___\   /\  ___\   /\  == \   
                        \ \ \-./\ \  \ \___  \     \ \ \/\ \  \ \  __\ \ \  __\ \ \ \  \ \ \____  \ \  __\      \ \  _"-.  \ \  __\   \ \____ \     \ \  _-/ \ \  __ \  \ \  __<   \ \___  \  \ \  __\   \ \  __<   
                         \ \_\ \ \_\  \/\_____\     \ \_____\  \ \_\    \ \_\    \ \_\  \ \_____\  \ \_____\     \ \_\ \_\  \ \_____\  \/\_____\     \ \_\    \ \_\ \_\  \ \_\ \_\  \/\_____\  \ \_____\  \ \_\ \_\ 
                          \/_/  \/_/   \/_____/      \/_____/   \/_/     \/_/     \/_/   \/_____/   \/_____/      \/_/\/_/   \/_____/   \/_____/      \/_/     \/_/\/_/   \/_/ /_/   \/_____/   \/_____/   \/_/ /_/ 
                                                                                                                                                                                                
    
                            

                                                                            Dev_tg: $dev                                            Telegram: $url
"@ -ForegroundColor Cyan

} 

# Функция для парсинга и показа ключа
function ParseAndShowKey {
    Write-Host "Поиск ключей активации MS Office..." -ForegroundColor Yellow

    $officeKey = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\16.0\Registration\*" -ErrorAction SilentlyContinue |
        Select-Object -ExpandProperty DigitalProductID -ErrorAction SilentlyContinue |
        ForEach-Object {
            [System.Text.Encoding]::Unicode.GetString($_[52..67]) -replace '\x00'
        }

    if ($officeKey) {
        Write-Host "Ключ активации MS Office: $officeKey" -ForegroundColor Green
    } else {
        Write-Host "Ключ не найден или MS Office не установлен." -ForegroundColor Red
        Start-Sleep -Seconds 2
        Clear-Host
        return
    }
}

# Функция для парсинга и сохранения ключа в файл
function ParseKeyToTxt {
    $documentsPath = [Environment]::GetFolderPath("MyDocuments")
    $filePath = Join-Path $documentsPath "OfficeKey.txt"

    Write-Host "Сохранение ключа в файл $filePath..." -ForegroundColor Yellow

    $officeKey = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\16.0\Registration\*" -ErrorAction SilentlyContinue |
        Select-Object -ExpandProperty DigitalProductID -ErrorAction SilentlyContinue |
        ForEach-Object {
            [System.Text.Encoding]::Unicode.GetString($_[52..67]) -replace '\x00'
        }

    if ($officeKey) {
        Set-Content -Path $filePath -Value "Ключ активации MS Office: $officeKey"
        Write-Host "Ключ успешно сохранён!" -ForegroundColor Green
    } else {
        Write-Host "Ключ не найден или MS Office не установлен." -ForegroundColor Red
        Start-Sleep -Seconds 2
        Clear-Host
        return
    }
}

# Функция для поиска OSPP.VBS
function Find-OSPP {
    Write-Host "Поиск файла OSPP.VBS..." -ForegroundColor Yellow

    # Системные папки для поиска
    $possiblePaths = @(
        "C:\Program Files\Microsoft Office\Office16\OSPP.VBS",
        "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS",
        "C:\Program Files\Microsoft Office\Office15\OSPP.VBS",
        "C:\Program Files (x86)\Microsoft Office\Office15\OSPP.VBS"
    )

    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            return $path
        }
    }

    Write-Host "Файл OSPP.VBS не найден." -ForegroundColor Red
    return $null
}

# Функция для выполнения OSPP.VBS с выводом информации
function Run-OSPP {
    Write-Host "Запуск OSPP.VBS в тихом режиме..." -ForegroundColor Yellow
    Write-Host ""
    $osppPath = Find-OSPP

    if ($osppPath) {
        $output = & cscript.exe $osppPath /dstatus
        $last5Chars = $output | Select-String -Pattern "Last 5 characters of installed product key:" | ForEach-Object { $_.Line }
        $licenseType = $output | Select-String -Pattern "LICENSE NAME:" | ForEach-Object { $_.Line }
        if ($last5Chars, $licenseType) {
            Write-Host ""
            Write-Host $licenseType -ForegroundColor DarkCyan
            Write-Host ""
            Write-Host $last5Chars -ForegroundColor Green
        } else {
            Write-Host "Не удалось найти информацию о ключе." -ForegroundColor Red
        }
    } else {
        Write-Host "Не найден файл OSPP.VBS. Убедитесь, что Office установлен." -ForegroundColor Red
        Start-Sleep -Seconds 2
        Clear-Host
    }
}

# Функция для вывода меню
function Show-Menu {
    Write-Host ""
    Write-Host "                                                                                                  ╔══════════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "                                                                                                  ║             Выберите функцию:            ║" -ForegroundColor Yellow
    Write-Host "                                                                                                  ╠══════════════════════════════════════════╣" -ForegroundColor Yellow
    Write-Host "                                                                                                  ║                                          ║" -ForegroundColor Yellow
    Write-Host "                                                                                                  ║       1. Показать ключи активации        ║" -ForegroundColor Yellow
    Write-Host "                                                                                                  ║                                          ║" -ForegroundColor Yellow
    Write-Host "                                                                                                  ║       2. Сохранить ключи в файл          ║" -ForegroundColor Yellow
    Write-Host "                                                                                                  ║                                          ║" -ForegroundColor Yellow
    Write-Host "                                                                                                  ║       3. Выполнить скрипт OSPP.VBS       ║" -ForegroundColor Yellow
    Write-Host "                                                                                                  ║                                          ║" -ForegroundColor Yellow
    Write-Host "                                                                                                  ╠══════════════════════════════════════════╣" -ForegroundColor Yellow
    Write-Host "                                                                                                  ║                99. Выход                 ║" -ForegroundColor Yellow
    Write-Host "                                                                                                  ╚══════════════════════════════════════════╝" -ForegroundColor Yellow
}

function Clear-Display {
      Clear-Host
    
}

# Основной блок
do {
    Show-Logo
    Show-Menu
    $choice = Read-Host "Введите номер функции"
    Clear-Display

    switch ($choice) {
        "1" { ParseAndShowKey }
        "2" { ParseKeyToTxt }
        "3" { Run-OSPP; Start-Sleep -Seconds 2 }
        "99" { Write-Host "Выход..."; Start-Sleep -Seconds 1.5 ; break }
        default {
            Write-Host "Неверный ввод, попробуйте снова." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Clear-Host
        }
    }

    Write-Host ""
} 
while ($choice -ne "99")