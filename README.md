# Добро пожаловать в репозиторий VBA demo

В данном репозитории представлен комплексный скрипт, который осуществляет загрузку файлов из S3 бакета, обработку данных файлов с помощью скриптов Power Query (выборка нужных данных и приведение датасетов к единой структуре) и формирование сводного файла с результатами обработки. Доступ к бакету в S3 является публичным, поэтому работу скрипта проверить может любой желающий до тех пор, пока остается действующий лимит.

Скрипт управляется и настраивается через визуальную форму, которая вызывается из файла vacancy_processor.xlsm. В указанном файле также присутствуют скрытые листы (их можно открыть при желании), которые служат бэкендом для скрипта - хранят справочную информацию и переменные настроек скрипта. Для работы скрипта нужны все файлы, расположенные в корне репозитория, за исключением файла gen_url.py, который представляет собой нескомпилированную версию python-скрипта для извлечения ссылок на объекты в бакете S3.

Исходный код файлов можно изучить в файлах модулей в директории Sources, код скриптов Power Query - путем открытия файлов с расширением *.odc* и уже открытием непосредственно окна редактирования запроса.

Наличия установленного Python не обязательно, поскольку скрипт для получения ссылок скомпилирован в exe файл. Имейте ввиду, что на старых версиях Windows он может не заработать!

Хочу подчеркнуть, что данный проект не имел своей целью создать "шедевр" кодинга от мира языков VBA и M, а только продемонстрировать уровень владения данными инструментами на уровне комплексных скриптов. На мой взгляд, прибегать к использованию VBA для реализации комплексных ETL проектов нужно только от безысходности, поскольку язык в настоящее время морально и функционально устарел.

Для неопытных юзеров git вот [ссылка](https://github.com/DanchukIvan/vba_demo/archive/refs/heads/main.zip) на скачивание файла.
