Надстройка AutocadUpLoad для excel позволяет скачать краткую ведомость из коммерческого предложения для экспорта в Автокад.

Надстройка состоит из: 

-стандартного класса надстройки, в котором агрегируются переменные активного приложения, рабочей книги и рабочего листа. Он также содержит методы, подписанные на стандартные события листа.

-класса панели, который содержит методы, подписанные на события элементов интерфейса, и вызываемые ими приватные методы.

-файла «Commands», который содержит небольшую иерархию классов-функторов, выделяющих на листе ячейки, задействованные в программе. Форматирование этих ячеек можно сбросить до прежнего состояния.

-файла "Loaders"с классом, в котором инкапсулирована логика создания excel-файла для выгрузки. Ведомость состоит из 3 столбцов: артикулы, названия и картинки.

При выгрузке ведомости также создаётся версия ведосомти в PDF, содержащая только картинки.
