/*
    (с) 2016 http://www.proasutp.com

    Функция-пример добавления новой записи в базу данных пользовательского архива (User Archive) WinCC.

    Вход:
        Нет.
    Выход:
        Нет.
*/

#include "apdefap.h"

void UserArchiveWritingValues()
{
    #define ARCHIVE_NAME "UserArchiveName"

    #define ARCHIVE_ID_FIELD_NUMBER 0
    #define ARCHIVE_FLOAT_FIELD_NUMBER 1
    #define ARCHIVE_STRING_FIELD_NUMBER 2

    #define STRING_FILED_SIZE 255
   
    UAHCONNECT hConnect = 0;
    UAHARCHIVE hArchive = 0;
    BOOL bIsOk = FALSE;
   
    long nIDFieldValue = 0;
    float fFloatFieldValue = 0;
    char szStringFieldValue[STRING_FILED_SIZE+1] = "";

    // Подключение к среде исполнения пользовательских архивов...

    bIsOk = uaConnect (&hConnect);

    if (!bIsOk) {
        printf
("Ошибка вызова uaConnect: %d\r\n", uaGetLastError());
        return;
        }

    // Подключение к конкретному пользовательскому архиву...

    bIsOk = uaQueryArchiveByName (hConnect, ARCHIVE_NAME, &hArchive);
    if (!bIsOk) {
        printf
("Ошибка вызова uaQueryArchiveByName: %d\r\n", uaGetLastError());
        return;
        }

    // Открытие пользовательского архива...

    bIsOk = uaArchiveOpen (hArchive);
    if (!bIsOk) {
        printf
("Ошибка вызова uaArchiveOpen: %d\r\n", uaGetLastError());
        return;
        }  

    // Перемещаемся к последней записи пользовательского архива...

    bIsOk = uaArchiveMoveLast (hArchive);
    if (!bIsOk) {
        printf
("Ошибка вызова uaArchiveMoveLast: %d\r\n", uaGetLastError());
        return;
        }  

    // Получаем ID текущей записи пользовательского архива...

    bIsOk = uaArchiveGetFieldValueLong (hArchive, ARCHIVE_ID_FIELD_NUMBER, &nIDFieldValue);
    if (!bIsOk) {
        printf
("Ошибка вызова uaArchiveGetFieldValueNumber: %d\r\n", uaGetLastError());
        return;
        }  

    // Начинаем формировать значения полей для новой записи пользовательского архива...

    // Устанавливаем следующий ID новой записи...

    nIDFieldValue++;
    bIsOk = uaArchiveSetFieldValueLong (hArchive, ARCHIVE_ID_FIELD_NUMBER, nIDFieldValue);
    if (!bIsOk) {
        printf
("Ошибка вызова uaArchiveSetFieldValueLong: %d\r\n", uaGetLastError());
        return;
        }      

    // Записываем значение с плавающей точкой в соответствующее поле пользовательского арихва...

    fFloatFieldValue = 1.2345;
    bIsOk = uaArchiveSetFieldValueFloat (hArchive, ARCHIVE_FLOAT_FIELD_NUMBER, fFloatFieldValue);
    if (!bIsOk) {
        printf
("Ошибка вызова uaArchiveSetFieldValueFloat: %d\r\n", uaGetLastError());
        return;
        }

    // Записываем строковое значение в соответствующее поле пользовательского арихва...

    strcpy
(szStringFieldValue, "Добро пожаловать на форум forum.proasutp.com!");
    bIsOk = uaArchiveSetFieldValueString (hArchive, ARCHIVE_STRING_FIELD_NUMBER, szStringFieldValue);
    if (!bIsOk) {
        printf
("Ошибка вызова uaArchiveSetFieldValueString: %d\r\n", uaGetLastError());
        return;
        }

    // Вставляем новую запись в пользовательский арихв...

    bIsOk = uaArchiveInsert (hArchive);
    if (!bIsOk) {
        printf
("Ошибка вызова uaArchiveInsert: %d\r\n", uaGetLastError());
        return;
        }

    // Закрываем пользовательский архив...

    bIsOk = uaArchiveClose (hArchive);

    if (!bIsOk) {
        printf
("Ошибка вызова uaArchiveClose: %d\r\n", uaGetLastError());
        return;
        }  

    // Закрываем подключение к пользовательскому архиву...

    bIsOk = uaReleaseArchive (hArchive);

    if (!bIsOk) {
        printf
("Ошибка вызова uaReleaseArchive: %d\r\n", uaGetLastError());
        return;
        }

    // Отключаемся от среды исполнения пользовательских архивов...

    bIsOk = uaDisconnect (hConnect);

    if (!bIsOk) {
        printf
("Ошибка вызова uaDisconnect: %d\r\n", uaGetLastError());
        return;
        }
}