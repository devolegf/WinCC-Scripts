/*
        (с) 2015 http://www.proasutp.com

        Функция-пример записи значения в ячейку таблицы Excel.

        Вход:
                Нет
        Выход:
                Нет

*/

#include "apdefap.h"

void WriteValueToExcelTable()
{  
    // указатели на объекты...

    __object* pExcel = NULL;
    __object* pWorkbooks = NULL;
    __object* pActiveWorkbook = NULL;
    __object* pCells = NULL;
   
    #define CELL_VALUE 2.34L;   // значение, которое будем записывать в ячейку таблицы Excel   
    #define TABLE_ROW 2         // номер строки, в нашем случае это строка 2
    #define TABLE_COL 1         // номер столбца, в нашем случае это столбец 1 или столбец А

    // Создаем объект приложения независимо от версии Excel по его ProgID...

    pExcel = __object_create("Excel.Application");
    if(NULL == pExcel) {
        printf
("Ошибка создания приложения Excel.Application!\r\n");
        return;
        }

    // Получаем указатель на книгу...
       
    pWorkbooks = pExcel->Workbooks;
    if(NULL == pWorkbooks) {
        __object_delete(pExcel);
        printf
("Ошибка получения указателя на Workbooks!\r\n");
        return;
        }

    // Делает окно Excel невидимым и открываем предварительно созданный в Excel тестовый файл 

    pExcel->Visible = 0;
    pWorkbooks->Open ("c:\\sample.xlsx");

    // Пишем тестовое значение в ячейку активной рабочей книги в активную таблицу в ячейку A2

    pCells = pExcel->Cells (TABLE_ROW, TABLE_COL);
    pCells->Value = CELL_VALUE;
    __object_delete (pCells);

    // Получаем указатель на активную рабочую книгу и сохраняем измененный файл

    pActiveWorkbook = pExcel->ActiveWorkbook;

    if(NULL == pActiveWorkbook) {
        __object_delete(pWorkbooks);
        __object_delete(pExcel);
        printf
("Ошибка получения указателя на pActiveWorkbook!\r\n");
        return;
        }

    pActiveWorkbook->Save();
    __object_delete(pActiveWorkbook);

    // Завершаем работу приложения и уничтожаем объекты
   
    pWorkbooks->Close();
    pExcel->Quit();
   
    __object_delete(pWorkbooks);
    __object_delete(pExcel);

    return;
}