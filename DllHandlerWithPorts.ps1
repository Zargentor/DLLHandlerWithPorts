# Скрытие консольного окна через WinAPI
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);'
[Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0)

# Создание GUI-формы
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Основная форма приложения
$Form = New-Object system.Windows.Forms.Form
$Form.ClientSize = New-Object System.Drawing.Point(600, 300)
$Form.text = "Регистрация компонент для просмотра сеансов"
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable

# Таблица для отображения результатов
$DataGridView2 = New-Object system.Windows.Forms.DataGridView
$DataGridView2.width = 585
$DataGridView2.height = 250
$DataGridView2.location = New-Object System.Drawing.Point(8, 49)
$DataGridView2.ColumnCount = 2
$DataGridView2.Columns[0].Name = "Найденные radmin"  # Колонка с путями к компонентам
$DataGridView2.Columns[0].Width = 300
$DataGridView2.Columns[1].Name = "Порты"            # Колонка с портами служб
$DataGridView2.Columns[1].Width = 200
$DataGridView2.Dock = [System.Windows.Forms.DockStyle]::Bottom

# Кнопка запуска поиска
$Button2 = New-Object system.Windows.Forms.Button
$Button2.text = "Вывести список существующих radmin"
$Button2.width = 350
$Button2.height = 30
$Button2.location = New-Object System.Drawing.Point(7, 10)
$Button2.Font = New-Object System.Drawing.Font('HP Simplified', 13)

# Добавление элементов управления на форму
$Form.controls.AddRange(@($DataGridView2, $Button2))

# Обработчик клика по кнопке поиска
$Button2.Add_MouseClick({ 
    $DataGridView2.Rows.Clear()      # Очистка предыдущих результатов
    SearchAllDll                     # Запуск основного поиска
})

# Обработчик двойного клика по записи
$DataGridView2.Add_CellDoubleClick({ 
    if ($DataGridView2.SelectedCells[0].Value.ToString() -like "*radmin*") {
        # Регистрация DLL с правами администратора
        Start-Process -FilePath "C:\Windows\System32\regsvr32.exe" -ArgumentList "`"$($DataGridView2.SelectedCells[0].Value.ToString())`"" -Verb runas
    } else {
        # Запуск CMD-скрипта с правами администратора
        RegisterRadmin -Path $DataGridView2.SelectedCells[0].Value.ToString()
    }
})

# Функция запуска файлов с повышенными привилегиями
function RegisterRadmin {
    param([Parameter(Mandatory)][string]$Path)
    Start-Process -FilePath $Path -Verb runas
}

# Основная функция поиска компонентов
function SearchAllDll {
    $pathx64 = "C:\Program Files\1cv8"  # Стандартный путь установки 1С
    
    # Фильтрация нерелевантных каталогов
    foreach ($Platformpath in Get-ChildItem -Path $pathx64 | Where-Object { 
        $_.Name -notlike "srvinfo*" -and $_.Name -notlike "common" -and $_.Name -notlike "conf" } | Select-Object -ExpandProperty FullName) {
        
        # Особый случай для версии 8.3.9.2170
        if ($Platformpath -notlike "*8.3.9.2170*") {
            # Поиск CMD-скриптов регистрации
            foreach ($binPath in Get-ChildItem -Path $Platformpath | Where-Object { $_.Name -like "bin" } | Select-Object -ExpandProperty FullName) {
                $RegFilesPaths = Get-ChildItem -Path $binPath | Where-Object { $_.Name -like "RegMSC.cmd" } | Select-Object -ExpandProperty FullName
                $ports = Get-PortsForPlatform -PlatformPath $Platformpath
                if ($RegFilesPaths -ne $null) {
                    $DataGridView2.Rows.Add($RegFilesPaths, $ports)  # Добавление в таблицу
                }
            }
        } else {
            # Поиск DLL-файлов для конкретной версии
            foreach ($binPath in Get-ChildItem -Path $Platformpath | Where-Object { $_.Name -like "bin" } | Select-Object -ExpandProperty FullName) {
                $RegFilesPaths = Get-ChildItem -Path $binPath | Where-Object { $_.Name -like "radmin.dll" } | Select-Object -ExpandProperty FullName
                $ports = Get-PortsForPlatform -PlatformPath $Platformpath
                if ($RegFilesPaths -ne $null) {
                    $DataGridView2.Rows.Add($RegFilesPaths, $ports)  # Добавление в таблицу
                }   
            }
        }
    }
}

# Функция поиска портов служб 1С
function Get-PortsForPlatform {
    param([Parameter(Mandatory)][string]$PlatformPath)
    
    # Экранирование спецсимволов для WMI-запроса
    $escapedPath = $PlatformPath -replace '\\', '\\' -replace '%', '%%'
    
    # Поиск служб связанных с платформой
    $services = Get-WmiObject -Query @"
    SELECT * FROM Win32_Service 
    WHERE (Name LIKE '1C:Enterprise%' OR Name LIKE '%Server Agent%')
    AND PathName LIKE '%$escapedPath%'
"@
    
    # Парсинг портов из параметров запуска
    $ports = @()
    foreach ($service in $services) {
        if ($service.PathName -match "-port\s+(\d+)") {
            $ports += $matches[1]
        }
    }
    return ($ports -join ", ")  # Возврат списка портов через запятую
}

# Запуск главного окна приложения
[void]$Form.ShowDialog()
