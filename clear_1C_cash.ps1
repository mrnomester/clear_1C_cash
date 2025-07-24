<#
.SYNOPSIS
    Скрипт для очистки кэша 1С на локальных и удаленных компьютерах
.DESCRIPTION
    Поддерживает очистку через административную шару, WMI и PowerShell Remoting
    Использует методы в порядке приоритета
.VERSION
    5.2
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$ComputerName,
    
    [switch]$AutoMode,
    
    [switch]$Force,
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath
)

#region Инициализация и настройки
$scriptVersion = "5.2"
$defaultLogPath = "$env:TEMP\clear_1c_cache.log"
$maxLogSize = 1MB
$global:ErrorActionPreference = "Stop"

# Определение пути для логов
if (-not $LogPath) {
    $LogPath = $defaultLogPath
}

# Инициализация логов
function Initialize-Log {
    param([string]$OperationType = "Локальная")
    
    try {
        if (Test-Path $LogPath) {
            $logSize = (Get-Item $LogPath).Length
            if ($logSize -gt $maxLogSize) {
                Remove-Item $LogPath -Force -ErrorAction Stop
            } else {
                return
            }
        }
        
        $header = @"
====================================================================
=            ЛОГ ОЧИСТКИ КЭША 1С (v$scriptVersion)                =
=            Дата: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")       =
=            Операция: $OperationType                             =
=            Пользователь: $env:USERNAME                          =
=            Компьютер: $env:COMPUTERNAME                         =
====================================================================

"@
        $header | Out-File -FilePath $LogPath -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-Console "Ошибка инициализации лога" -Level "ERROR"
    }
}

# Функция записи в лог
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $logEntry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    
    try {
        $logEntry | Out-File -FilePath $LogPath -Encoding UTF8 -Append -ErrorAction Stop
    }
    catch {
        Write-Console "Ошибка записи в лог" -Level "ERROR"
    }
}

# Функция вывода в консоль
function Write-Console {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $color = switch ($Level) {
        "SUCCESS" { "Green" }
        "ERROR"   { "Red" }
        "WARNING" { "Yellow" }
        default   { "Cyan" }
    }
    
    $prefix = switch ($Level) {
        "SUCCESS" { "[УСПЕХ] " }
        "ERROR"   { "[ОШИБКА] " }
        "WARNING" { "[ВНИМАНИЕ] " }
        default   { "[ИНФО] " }
    }
    
    if (-not $AutoMode -or $Level -in @("ERROR", "WARNING")) {
        Write-Host "$prefix$Message" -ForegroundColor $color
    }
}

# Инициализация логов
Initialize-Log -OperationType "Локальная"
Write-Log "=== НАЧАЛО НОВОЙ ОПЕРАЦИИ ===" -Level "INFO"
#endregion

#region Вспомогательные функции
function Test-IsAdmin {
    try {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        Write-Console "Ошибка проверки прав администратора" -Level "ERROR"
        return $false
    }
}

function Test-ComputerAvailability {
    param([string]$Computer)
    
    try {
        $result = Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction Stop
        if (-not $result) {
            Write-Console "Компьютер $Computer недоступен" -Level "ERROR"
            Write-Log "Компьютер $Computer недоступен" -Level "ERROR"
        }
        return $result
    }
    catch {
        Write-Console "Ошибка проверки доступности компьютера $Computer" -Level "ERROR"
        Write-Log "Ошибка проверки доступности компьютера ${Computer}: $_" -Level "ERROR"
        return $false
    }
}

function Test-AdminShare {
    param([string]$Computer)
    
    try {
        $result = Test-Path "\\$Computer\c$\Windows" -ErrorAction Stop
        if (-not $result) {
            Write-Log "Административная шара на компьютере $Computer недоступна" -Level "WARNING"
        }
        return $result
    }
    catch {
        Write-Log "Ошибка проверки административной шары на ${Computer}: $_" -Level "ERROR"
        return $false
    }
}

function Test-WMI {
    param([string]$Computer)
    
    try {
        $null = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer -ErrorAction Stop
        return $true
    }
    catch {
        Write-Log "Ошибка проверки WMI на ${Computer}: $_" -Level "ERROR"
        return $false
    }
}

function Test-PSRemoting {
    param([string]$Computer)
    
    try {
        $null = Invoke-Command -ComputerName $Computer -ScriptBlock { $true } -ErrorAction Stop
        return $true
    }
    catch {
        Write-Log "Ошибка проверки PSRemoting на ${Computer}: $_" -Level "ERROR"
        return $false
    }
}

function Test-1CProcesses {
    param([string]$Computer = $env:COMPUTERNAME)
    
    try {
        if ($Computer -eq $env:COMPUTERNAME) {
            $processes = Get-Process -Name "1cv8*" -ErrorAction SilentlyContinue
        }
        else {
            $processes = Get-WmiObject -Query "SELECT * FROM Win32_Process WHERE Name LIKE '1cv8%'" -ComputerName $Computer -ErrorAction SilentlyContinue
        }

        if ($processes) {
            $count = $processes.Count
            Write-Console "Обнаружены запущенные процессы 1С ($count)" -Level "WARNING"
            Write-Log "Обнаружены запущенные процессы 1С ($count) на $Computer" -Level "WARNING"
            return $true
        }
        return $false
    }
    catch {
        Write-Console "Ошибка проверки процессов 1С" -Level "ERROR"
        Write-Log "Ошибка проверки процессов 1С на ${Computer}: $_" -Level "ERROR"
        return $false
    }
}
#endregion

#region Функции очистки кэша
function Clear-1CCache {
    param(
        [string]$ProfilePath,
        [string]$Computer = $env:COMPUTERNAME
    )
    
    $cachePaths = @(
        "$ProfilePath\AppData\Local\1C\1Cv8",
        "$ProfilePath\AppData\Roaming\1C\1Cv8"
    )
    
    $pattern = '^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$'
    $deleted = $false
    $totalDeleted = 0
    
    foreach ($path in $cachePaths) {
        if (-not (Test-Path -LiteralPath $path)) {
            Write-Log "Путь $path не найден" -Level "INFO"
            continue
        }
        
        try {
            $folders = Get-ChildItem -LiteralPath $path -Directory -ErrorAction Stop | 
                       Where-Object { $_.Name -match $pattern }
            
            $count = $folders.Count
            if ($count -gt 0) {
                Write-Log "Найдено $count кэш-папок в $path" -Level "INFO"
            } else {
                Write-Log "Кэш-папки не найдены в $path" -Level "INFO"
            }
            
            foreach ($folder in $folders) {
                try {
                    $folderPath = $folder.FullName
                    
                    if ($PSCmdlet.ShouldProcess($folderPath, "Удаление кэш-папки")) {
                        Remove-Item -LiteralPath $folderPath -Recurse -Force -ErrorAction Stop
                        Write-Log "Удалено: $folderPath" -Level "SUCCESS"
                        $deleted = $true
                        $totalDeleted++
                    }
                }
                catch {
                    Write-Log "Ошибка удаления $($folder.FullName): $($_.Exception.Message)" -Level "ERROR"
                }
            }
        }
        catch {
            Write-Log "Ошибка доступа к $path : $($_.Exception.Message)" -Level "ERROR"
        }
    }
    
    if ($totalDeleted -gt 0) {
        Write-Log "Всего удалено кэш-папок: $totalDeleted" -Level "SUCCESS"
    }
    
    return $deleted
}

function Clear-ViaAdminShare {
    param([string]$Computer)
    
    try {
        Write-Console "Используем метод административной шары" -Level "INFO"
        Write-Log "Попытка очистки через административную шару на $Computer" -Level "INFO"
        
        $usersPath = "\\$Computer\c$\Users"
        $userFolders = Get-ChildItem $usersPath -Directory | 
                      Where-Object { $_.Name -notmatch 'default|public|Администратор|Administrator' }

        Write-Log "Найдено профилей пользователей: $($userFolders.Count)" -Level "INFO"
        
        $anyDeleted = $false
        foreach ($userFolder in $userFolders) {
            if (Clear-1CCache -ProfilePath $userFolder.FullName -Computer $Computer) {
                $anyDeleted = $true
            }
        }
        
        if ($anyDeleted) {
            Write-Console "Кэш успешно очищен" -Level "SUCCESS"
        } else {
            Write-Console "Не найдено кэш-папок для удаления" -Level "WARNING"
        }
        return $anyDeleted
    }
    catch {
        Write-Console "Ошибка при использовании административной шары: $($_.Exception.Message)" -Level "ERROR"
        Write-Log "Ошибка при использовании административной шары: $_" -Level "ERROR"
        return $false
    }
}

function Clear-ViaWMI {
    param([string]$Computer)
    
    try {
        Write-Console "Используем метод WMI" -Level "INFO"
        Write-Log "Попытка очистки через WMI на $Computer" -Level "INFO"
        
        $profiles = Get-WmiObject -Class Win32_UserProfile -ComputerName $Computer -ErrorAction Stop |
                    Where-Object { 
                        $_.Special -eq $false -and 
                        $_.LocalPath -notmatch 'systemprofile|service|default' -and
                        $_.Loaded -eq $false
                    }

        Write-Log "Найдено профилей через WMI: $($profiles.Count)" -Level "INFO"
        
        $anyDeleted = $false
        foreach ($profile in $profiles) {
            $userPath = $profile.LocalPath
            $drive = $userPath.Substring(0, 1)
            $pathWithoutDrive = $userPath.Substring(3)
            $profilePath = "\\$Computer\$drive`$$pathWithoutDrive"
            
            if (Clear-1CCache -ProfilePath $profilePath -Computer $Computer) {
                $anyDeleted = $true
            }
        }
        
        if ($anyDeleted) {
            Write-Console "Кэш успешно очищен" -Level "SUCCESS"
            # Запись в Event Log
            try {
                $eventMessage = "Кэш 1С успешно очищен на компьютере $Computer"
                $null = New-EventLog -LogName "Application" -Source "1CCacheCleaner" -ErrorAction SilentlyContinue
                Write-EventLog -LogName "Application" -Source "1CCacheCleaner" -EventId 1001 -EntryType Information -Message $eventMessage
                Write-Log "Запись в Event Log создана" -Level "INFO"
            }
            catch {
                Write-Log "Не удалось записать в Event Log: $_" -Level "WARNING"
            }
        } else {
            Write-Console "Не найдено кэш-папок для удаления" -Level "WARNING"
        }
        return $anyDeleted
    }
    catch {
        Write-Console "Ошибка при использовании WMI: $($_.Exception.Message)" -Level "ERROR"
        Write-Log "Ошибка при использовании WMI: $_" -Level "ERROR"
        return $false
    }
}

function Clear-ViaPSRemoting {
    param([string]$Computer)
    
    try {
        Write-Console "Используем метод PowerShell Remoting" -Level "INFO"
        Write-Log "Попытка очистки через PSRemoting на $Computer" -Level "INFO"
        
        $scriptBlock = {
            $cachePaths = @(
                "$env:USERPROFILE\AppData\Local\1C\1Cv8",
                "$env:USERPROFILE\AppData\Roaming\1C\1Cv8"
            )
            
            $pattern = '^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$'
            $totalDeleted = 0
            
            foreach ($path in $cachePaths) {
                if (Test-Path -LiteralPath $path) {
                    $folders = Get-ChildItem -LiteralPath $path -Directory | 
                               Where-Object { $_.Name -match $pattern }
                    
                    foreach ($folder in $folders) {
                        try {
                            Remove-Item -LiteralPath $folder.FullName -Recurse -Force -ErrorAction Stop
                            $totalDeleted++
                        }
                        catch {
                            Write-Error "Ошибка удаления $($folder.FullName): $_"
                        }
                    }
                }
            }
            
            return $totalDeleted
        }
        
        $result = Invoke-Command -ComputerName $Computer -ScriptBlock $scriptBlock -ErrorAction Stop
        
        if ($result -gt 0) {
            Write-Console "Удалено $result кэш-папок через PSRemoting" -Level "SUCCESS"
            Write-Log "Удалено $result кэш-папок через PSRemoting на $Computer" -Level "SUCCESS"
            return $true
        } else {
            Write-Console "Не найдено кэш-папок для удаления через PSRemoting" -Level "WARNING"
            Write-Log "Не найдено кэш-папок для удаления через PSRemoting на $Computer" -Level "WARNING"
            return $false
        }
    }
    catch {
        Write-Console "Ошибка при использовании PSRemoting: $($_.Exception.Message)" -Level "ERROR"
        Write-Log "Ошибка при использовании PSRemoting: $_" -Level "ERROR"
        return $false
    }
}

function Clear-Remote1CCache {
    param([string]$Computer)
    
    Write-Console "Очистка кэша 1С на $Computer" -Level "INFO"
    Write-Log "=== Начало очистки кэша на $Computer ===" -Level "INFO"
    
    if (-not (Test-ComputerAvailability -Computer $Computer)) {
        return
    }
    
    if (Test-1CProcesses -Computer $Computer) {
        if ($Force) {
            Write-Console "Принудительная очистка, несмотря на запущенные процессы 1С" -Level "WARNING"
            Write-Log "Принудительная очистка при запущенных процессах 1С" -Level "WARNING"
        } else {
            if (-not $AutoMode) {
                $choice = Read-Host "Продолжить очистку? (y/n)"
                if ($choice -ne 'y') {
                    Write-Console "Отмена операции" -Level "WARNING"
                    Write-Log "Пользователь отменил операцию из-за запущенных процессов 1С" -Level "WARNING"
                    return
                }
            } else {
                Write-Console "Обнаружены запущенные процессы 1С, пропускаем $Computer (используйте -Force для принудительной очистки)" -Level "WARNING"
                Write-Log "Обнаружены запущенные процессы 1С, пропускаем $Computer" -Level "WARNING"
                return
            }
        }
    }
    
    # Проверяем методы в порядке приоритета
    if (Test-AdminShare -Computer $Computer) {
        Clear-ViaAdminShare -Computer $Computer
        return
    }
    
    if (Test-WMI -Computer $Computer) {
        Clear-ViaWMI -Computer $Computer
        return
    }
    
    if (Test-PSRemoting -Computer $Computer) {
        Clear-ViaPSRemoting -Computer $Computer
        return
    }
    
    Write-Console "Нет доступных методов для очистки" -Level "ERROR"
    Write-Log "Не найдено доступных методов для очистки на $Computer" -Level "ERROR"
}

function Clear-Local1CCache {
    Write-Console "Очистка локального кэша 1С" -Level "INFO"
    Write-Log "=== Начало локальной очистки кэша ===" -Level "INFO"
    
    if (Test-1CProcesses) {
        if ($Force) {
            Write-Console "Принудительная очистка, несмотря на запущенные процессы 1С" -Level "WARNING"
            Write-Log "Принудительная очистка при запущенных процессах 1С" -Level "WARNING"
        } else {
            if (-not $AutoMode) {
                $choice = Read-Host "Продолжить очистку? (y/n)"
                if ($choice -ne 'y') {
                    Write-Console "Отмена операции" -Level "WARNING"
                    Write-Log "Пользователь отменил операцию из-за запущенных процессов 1С" -Level "WARNING"
                    return
                }
            } else {
                Write-Console "Обнаружены запущенные процессы 1С, пропускаем очистку (используйте -Force для принудительной очистки)" -Level "WARNING"
                Write-Log "Обнаружены запущенные процессы 1С, пропускаем очистку" -Level "WARNING"
                return
            }
        }
    }
    
    if (Clear-1CCache -ProfilePath $env:USERPROFILE) {
        Write-Console "Локальная очистка завершена" -Level "SUCCESS"
    } else {
        Write-Console "Не найдено кэш-папок для удаления" -Level "WARNING"
    }
}
#endregion

#region Главное меню и интерфейс
function Show-Menu {
    if ($AutoMode) { return }
    
    Clear-Host
    Write-Host "====================================================================" -ForegroundColor Cyan
    Write-Host "=            УТИЛИТА ОЧИСТКИ КЭША 1С (v$scriptVersion)            =" -ForegroundColor Cyan
    Write-Host "====================================================================" -ForegroundColor Cyan
    Write-Host "= 1. Очистить кэш 1С на ЛОКАЛЬНОМ компьютере                       =" -ForegroundColor Cyan
    Write-Host "= 2. Очистить кэш 1С на УДАЛЕННОМ компьютере                       =" -ForegroundColor Cyan
    Write-Host "= 3. Просмотреть лог                                               =" -ForegroundColor Cyan
    Write-Host "=                                                                  =" -ForegroundColor Cyan
    Write-Host "= 0. ВЫХОД                                                         =" -ForegroundColor Cyan
    Write-Host "====================================================================" -ForegroundColor Cyan
}
#endregion

#region Точка входа
if (-not (Test-IsAdmin)) {
    Write-Console "Требуются права администратора. Перезапуск..." -Level "WARNING"
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`"" -Verb RunAs
    exit
}

if (-not $AutoMode -and -not $ComputerName -and -not $Force) {
    # Режим с меню
    while ($true) {
        Show-Menu
        $choice = Read-Host "Выберите действие"
        
        switch ($choice) {
            "1" { Clear-Local1CCache }
            "2" { 
                $computer = Read-Host "Введите имя или IP-адрес компьютера"
                Clear-Remote1CCache -Computer $computer 
            }
            "3" { 
                if (Test-Path $LogPath) {
                    try {
                        notepad $LogPath
                    }
                    catch {
                        Write-Console "Не удалось открыть лог" -Level "ERROR"
                    }
                }
                else {
                    Write-Console "Лог не найден" -Level "WARNING"
                }
            }
            "0" { exit }
            default { 
                Write-Console "Неверный выбор!" -Level "ERROR"
            }
        }
        
        if ($choice -in "1","2") {
            Write-Host "`nНажмите любую клавишу для продолжения..."
            [void][System.Console]::ReadKey($true)
        }
    }
}
else {
    # Прямой режим (для автоматизации)
    if ($ComputerName) {
        Clear-Remote1CCache -Computer $ComputerName
    }
    else {
        Clear-Local1CCache
    }
}
#endregion