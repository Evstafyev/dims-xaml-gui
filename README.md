# DIMS XAML GUI
Dynamic Inventory Managements System based on PowerShell with XAML GUI
## Для чего нужно?
Скрипт используется для быстрого сбора инветарных данных о хостах окружения. Для управления реализован графический интерфейс на базе XAML. 
## Где взять конфигурацию ?
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
## Что потребуется дополнительно?
### Модули
**PDFTools** - используется для экспорта файлов отчетов в .pdf
### Переменные окружения
Для импорта конфигурации необходимо создать переменную окружения DIMS со ссылкой на целевой каталог.
## Как это работает?
### Графический интерфейс
Основные элементы:
1. wksCbox - комбобокс, содержащий имена доступных на момент сканирования хостов
2. Область NIC - содержит сведения об основном сетевом адаптере хоста:
  - статус подключения "ICMPv4 Status"
  - тип подключения "Connection"
  - текущий IP адрес "IPv4 address"
  - статус DHCP "DHCP status"
  - дату получения аренды "Lease obtained"
  - адрес шлюза "Gateway"
  - MAC-адрес
  - модель "Model"
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
### Коротко
1. Скрипт получает конфигурацию из файла config.xml
2. Выполняется сборка XAML интерфейса внутри функции Get-Xaml, из файла источника
3. Из целевого DN получаем список рабочих станций в состоянии Enabled для последующего сканирования
4. Из полученного списка, на основании проверки ICMP, формируется итоговый - из доступных на момент проверки рабочих станций.
5. На основании данных из пункта 4 заполняется ComboBox $wksCbox
6. Пользователь выбирает целевой хост из выпадающего списика и запускает процесс сканирования нажатием кнопки "Scan".  
### Дополнительные функции
Также доступны 5 вариантов дополнительных запросов c выводом результатов в GridView: USB, HDD, RAM, Display, Printers.
## Что внутри?
Получение данных о состоянии хостов реализовано посредством выполнения WMI запросов. Запросы выполняются синхронно. В данной реализации, множественные асинхронные запросы не поддерживаются.
## А можно чуть-чуть поподробнее? Описание графического интерфейса.
