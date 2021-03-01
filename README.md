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
Короткое описание алгоритма работы:
1. Скрипт получает конфигурацию из файла config.xml
2. Выполняется сборка XAML интерфейса внутри функции Get-Xaml, из файла источника
3. Из целевого DN получаем список рабочих станций в состоянии Enabled для последующего сканирования
4. Из полученного списка, на основании проверки ICMP, формируется итоговый - из доступных на момент проверки рабочих станций.
5. На основании данных из пункта 4 заполняется ComboBox $wksCbox, из которого впоследстии производится выборка хостов для проверки
## Что внутри?
Получение данных о состоянии хостов реализовано за счет выполнения WMI запросов.
## Асинхронное выполнение?
Невозможно. В данной реализации все действия выполняютяс последовательно:
1. Пользователь выбирает из выпадающего списка
## А можно чуть-чуть поподробнее? Описание графического интерфейса.
