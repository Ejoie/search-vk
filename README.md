# search-vk

Предыдущая программа формирует отчет в формате excel, где система обходит в своей базе всех авторов, 
чьи работы защищались в указанный период (по данным ВАК). И, если у автора нет связи с профилем не elibrary, 
запрашивает на последнем данные о публикациях всех однофамильцах автора и сохраняет.
Итоговый файл excel хранит данные о найденных людях с сайта elibrary.

Данный скрипт обрабатывает полученный файл, выполняя поиск в вк людей, найденных с помощью предыдущей программы.
Из предположения, что люди могут менять город проживания со временем, сначала поиск выполняется в указанном городе, 
где проходила защита работы, а затем выполняется поиск без привязки к городу.
По итогу отчет формируется в формате html с ссылками на найденные аккаунты в соц. сети ВК.
