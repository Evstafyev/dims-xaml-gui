# DIMS XAML GUI
Dynamic Inventory Managements System based on PowerShell with XAML GUI
![](gui_.png)
## Назначение
Скрипт используется для быстрого сбора инветарных данных о хостах окружения. Для управления реализован графический интерфейс на базе XAML. 
## Конфигурация
Конфигурация скрипта расположена в файле config.xml  
В составе конфигурационного .xml файла представлены следующие параметры:  
1. Узел Form:  
  - FormPath, путь к файлу разметки формы XAML -  mainForm.xaml
  - LoaderPath, путь к файлу .ps скрипта - загрузчика формы, - loader.ps1
2. Узел exportPath:  
  - tempPath, путь к каталогу для хранения отчетов
  - logPath, путь к кататлогу для хранения логов
3. Узел adSearchBase:
  - wksDn, DistinguishedName в текущей конфигурации Active Directory, из которого должны бьть получены данные о записях рабочих станций для сканирования
  - usrDn, DistinguishedName в текущей конфигурации Active Directory, из которого должны быть получены данные о записях пользователей
## Что потребуется?
### Модули
**PDFTools** - используется для экспорта файлов отчетов в .pdf
### Переменные окружения
Для импорта конфигурации необходимо создать переменную окружения DIMS со ссылкой на целевой каталог.
## Как работает?
### Графический интерфейс
Основные элементы:
1. Комбобокс wksCbox - содержит имена доступных на момент сканирования хостов
2. Область NIC - содержит сведения об основном сетевом адаптере хоста:
  - "ICMPv4 Status", статус подключения 
  - "Connection", тип подключения 
  - "IPv4 address", текущий IP адрес 
  - "DHCP status", статус DHCP 
  - "Lease obtained", дату получения аренды 
  - "Gateway", адрес шлюза 
  - "MAC", MAC-адрес сетевого адаптера
  - "Model", модель сетевого адаптера
3. Область Active Directory:
  - "USR DN", откуда получен список имен пользователей
  - "WKS DN", откуда получены список рабочих станций
  - "Enabled WKS", рабочие станции в статусе Enabled
  - "Disabled WKS", рабочие станции в статусе Disabled
  - "Enabled users", УЗ пользователей в статусе Enabled
  - "Disabled users", УЗ пользователей в статусе Disabled
  - "WKS", всего доступно рабочих станций
  - "Users", всего доступно УЗ пользователей
  - "UP", всего рабочих станций в статусе UP
  - "Down", всего рабочих станций в статусе Down
4. Область Workstation, содержит основную информаци о рабочей станции:
  - "Model", модель системного блока (класс Win32_ComputerSystem)
  - "Serial", серийный номер (класс Win32_BIOS)
  - "OS", вресия операционной системы (класс Win32_OperatingSystem_
  - "CPU", модель ЦПУ (класс Win32_Processor)
  - "Cores", количество ядер ЦПУ (класс Win32_Processor)
  - "MOBO", серийный номер материнской платы (класс Win32_Baseboard)
  - "RAM", суммарный объем оперативной памяти (класс Win32_PhysicalMemory)
  - "TYPE", тип памяти - определен по частоте (класс Win32_PhysicalMemory)
  - "BANKS", суммарное количество планок памяти (класс Win32_PhysicalMemory)
5. Кнопка Scan - выполняет запуск базового сканирования, с заполнением вышеописанных областей интерфейса.
6. Кнопка USB - выполняет запуск сканирования USB с выводом отчета в GridView
7. Кнопка HDD - выполняет запуск сканирования HDD с выводом отчета в GridView
8. Кнопка RAM - выполняет запуск сканирования RAM с выводом отчета в GridView 
9. Кнопка DISPLAY - выполняет запуск сканирования DISPLAY с выводом отчета в GridView 
10. Кнопка Printers - выполняет запуск сканирования Printers с выводом отчета в GridView
11. Лейбл ProcData - содержит данные о владельце текущего .ps процесса (procOwner), выделенный объем памяти (procMemory), идентиифкатор текущего процесса (PID) и  
уникальный идентификатор сессии (sessionId).  
### Коротко
1. Скрипт получает конфигурацию из файла config.xml
2. Выполняется сборка XAML интерфейса внутри функции Get-Xaml, из файла источника
3. Из целевого DN получаем список рабочих станций в состоянии Enabled для последующего сканирования
4. Из полученного списка, на основании проверки ICMP, формируется итоговый - из доступных на момент проверки рабочих станций.
5. На основании данных из пункта 4 заполняется ComboBox $wksCbox
6. Пользователь выбирает целевой хост из выпадающего списика и запускает процесс базового сканирования нажатием кнопки "Scan". 
7. Также пользователь может инициализировать дополнительные запросы путем нажатия кнопок "USB", "HDD", "RAM", "Display", "Printers". Запросы выполняются независимо от базового сканирования.
8. По результатам выполнения сканирования из пунтка 6 пользователь может выполнить эспорт данных, отраженных в графическом интрефейсе, в PDF отчет, нажатием кнопки "PDF".   
### Дополнительные функции
Также доступны 5 вариантов дополнительных запросов c выводом результатов в GridView: USB, HDD, RAM, Display, Printers:
1. **USB** - отчет о подключенных USB устройствах
2. **HDD** - отчет о подключенных HDD устройствах
3. **RAM** - отчет о модулях памяти
4. **Display** - отчет о используемых мониторах
5. **Printers** - отчет о подключнных принтерах 
## Что внутри?
Получение данных о состоянии хостов реализовано посредством выполнения WMI запросов к классам Win32. Запросы выполняются синхронно. В данной реализации, множественные асинхронные запросы не поддерживаются. Пример основных типов запросов из конструкции switch функции Get-PcData ниже:
```powershell
        switch($type){

        "CS" {$data = gwmi Win32_ComputerSystem -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "CPU" {$data = gwmi Win32_Processor -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "BIOS" {$data = gwmi Win32_BIOS -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "OS" {$data = gwmi Win32_OperatingSystem -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "MOBO" {$data = gwmi Win32_Baseboard -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "RAM" {$data = gwmi Win32_PhysicalMemory -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "USB" {$data = gwmi Win32_UsbDevice -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "HDD" {$data = gwmi Win32_DiskDrive -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "NIC" {$data = gwmi Win32_NetworkAdapterConfiguration -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "PRINTER" {$data = gwmi Win32_Printer -cn $computername -ErrorAction Stop -ErrorVariable $err}
        "DISPLAY" {$data = gwmi WmiMonitorID -Namespace root\wmi -cn $computername -ErrorAction Stop -ErrorVariable $err}

        }
```
## Как запустить?
```powershell
powershell.exe .\dims-gui-xaml.ps1
```
