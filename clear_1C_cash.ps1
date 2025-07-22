# Функция для проверки прав администратора
function Test-IsAdmin {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Проверяем, запущен ли скрипт с правами администратора
if (-not (Test-IsAdmin)) {
    Write-Host "Скрипт не запущен с правами администратора. Перезапускаем с повышенными правами..."
    
    # Получаем полный путь к текущему скрипту
    $scriptPath = $MyInvocation.MyCommand.Definition
    
    # Запускаем скрипт с повышенными правами
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
    exit
}

#requires -Version 5.1

function Clear-1CCache {
    param(
        [string]$ProfilePath,
        [string]$Computer
    )
    
    $cachePaths = @(
        "$ProfilePath\AppData\Local\1C\1Cv8",
        "$ProfilePath\AppData\Roaming\1C\1Cv8"
    )
    
    $pattern = '^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$'
    $deleted = $false
    
    foreach ($path in $cachePaths) {
        if (-not (Test-Path $path)) { 
            Write-Host "Каталог 1Cv8 не существует: $path" -ForegroundColor DarkGray
            continue
        }
        
        try {
            $folders = Get-ChildItem $path -Directory | 
                       Where-Object { $_.Name -match $pattern }
            
            if (-not $folders) {
                Write-Host "Не найдено кэш-папок (GUID) в: $path" -ForegroundColor DarkGray
                continue
            }
            
            foreach ($folder in $folders) {
                try {
                    $folderPath = $folder.FullName
                    
                    # Удаление через файловую систему
                    Remove-Item $folderPath -Recurse -Force -ErrorAction Stop
                    
                    Write-Host "Удалено: $folderPath" -ForegroundColor Green
                    $deleted = $true
                }
                catch {
                    Write-Host "Ошибка удаления: $($folder.FullName)" -ForegroundColor Red
                    Write-Host "Причина: $($_.Exception.Message)" -ForegroundColor DarkRed
                }
            }
        }
        catch {
            Write-Host "Ошибка доступа к $path : $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    if (-not $deleted) {
        Write-Host "Не найдено кэш-папок для удаления" -ForegroundColor Cyan
    }
}

# Главное меню
while ($true) {
    $input = Read-Host "`nВведите имя ПК (0=очистить локально, Enter=выход)"
    
    if ([string]::IsNullOrWhiteSpace($input)) { 
        Write-Host "Выход..." -ForegroundColor Cyan
        exit 
    }

    if ($input -eq "0") {
        Write-Host "`n=== Очистка ЛОКАЛЬНОГО кэша 1С ===" -ForegroundColor Cyan
        
        # Проверка запущенных процессов 1С
        $processes = Get-Process -Name "1cv8*" -ErrorAction SilentlyContinue
        if ($processes) {
            Write-Host "ВНИМАНИЕ: Обнаружены запущенные процессы 1С!" -ForegroundColor Red
            Write-Host "Список процессов:"
            $processes | Format-Table Id, Path -AutoSize
            Write-Host "Для корректной очистки закройте все сеансы 1С." -ForegroundColor Yellow
            continue
        }
        
        Clear-1CCache -ProfilePath $env:USERPROFILE
        Write-Host "`nЛокальная очистка завершена!" -ForegroundColor Green
    }
    else {
        $pcName = $input.Trim()
        Write-Host "`n=== Очистка кэша 1С на $pcName ===" -ForegroundColor Cyan
        
        # Проверка доступности
        if (-not (Test-Connection $pcName -Count 1 -Quiet)) {
            Write-Host "ОШИБКА: Компьютер $pcName недоступен!" -ForegroundColor Red
            Write-Host "Проверьте сетевое подключение и имя ПК" -ForegroundColor Yellow
            continue
        }
        
        # Проверка доступа к C$
        $adminShare = "\\$pcName\c$"
        if (-not (Test-Path $adminShare)) {
            Write-Host "ОШИБКА: Нет доступа к $adminShare" -ForegroundColor Red
            Write-Host "Убедитесь что:`n- Общий доступ C$ включен`n- У вас есть права администратора" -ForegroundColor Yellow
            continue
        }
        
        try {
            # Получение профилей через WMI
            Write-Host "Получение списка пользователей..." -ForegroundColor DarkCyan
            $profiles = Get-WmiObject -Class Win32_UserProfile -ComputerName $pcName -ErrorAction Stop |
                        Where-Object { 
                            $_.Special -eq $false -and 
                            $_.LocalPath -notmatch 'systemprofile|service|default' -and
                            $_.Loaded -eq $false
                        }
            
            if (-not $profiles) {
                Write-Host "Не найдено профилей для очистки" -ForegroundColor Yellow
                continue
            }
            
            Write-Host "Найдено профилей: $($profiles.Count)" -ForegroundColor DarkCyan
            
            # Обработка профилей через сетевой путь
            foreach ($profile in $profiles) {
                $userPath = $profile.LocalPath
                
                # Правильное формирование сетевого пути
                $drive = $userPath.Substring(0, 1)
                $pathWithoutDrive = $userPath.Substring(3)
                $networkPath = "\\$pcName\$drive`$$pathWithoutDrive"
                
                Write-Host "`nОбработка профиля: $userPath" -ForegroundColor Magenta
                Clear-1CCache -ProfilePath $networkPath
            }
            
            Write-Host "`nУдаленная очистка завершена!" -ForegroundColor Green
        }
        catch {
            Write-Host "`nКРИТИЧЕСКАЯ ОШИБКА: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Проверьте: `n- Права администратора `n- Доступ к WMI `n- Брандмауэр" -ForegroundColor Yellow
        }
    }
}