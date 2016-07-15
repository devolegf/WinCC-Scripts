/*
    (с) 2016 http://www.proasutp.com

    Функция-пример чтения значений полей записей пользовательского архива (User Archive) WinCC.

    Вход:
        Нет.
    Выход:
        Нет.
*/

#include "apdefap.h"

void UserArchiveReadingValues()
{
    #define ARCHIVE_NAME "UserArchiveName"

    #define ARCHIVE_ID_FIELD_NUMBER 0
    #define ARCHIVE_FLOAT_FIELD_NUMBER 1
    #define ARCHIVE_STRING_FIELD_NUMBER 2

    #define STRING_FILED_SIZE 255
   
    UAHCONNECT hConnect = 0;
    UAHARCHIVE hArchive = 0;
    long nError = 0;
    BOOL bIsOk = FALSE;
   

    long nCount = 0, i = 0;

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

    // Получение количества записей в пользовательском архиве...

    bIsOk = uaArchiveGetCount (hArchive, &nCount);
    nError = uaGetLastError();
    if (!bIsOk && (nError != 0)) {
        printf
("Ошибка вызова uaArchiveGetCount: %d\r\n", nError);
        return;
        }

    // Если записей нет, дальнейшее выполнение фукнции бессмысленно

    if (nCount == 0) {
        printf
("Пользовательский архив %s пустой!\r\n", ARCHIVE_NAME);
        }
    else {

        // Перемещаем указатель записей на первую запись архива

        bIsOk = uaArchiveMoveFirst (hArchive);
        if (!bIsOk) {
            printf
("Ошибка вызова uaArchiveMoveFirst: %d\r\n", uaGetLastError());
            return;
            }

        // Выводим значения всех полей в окно диагностики...
   
        for (i = 1; i <= nCount; i++) {

            // Читаем поле ID текущей записи...

            bIsOk = uaArchiveGetFieldValueLong (hArchive, ARCHIVE_ID_FIELD_NUMBER, &nIDFieldValue);

            if (!bIsOk) {
                printf
("Ошибка вызова uaArchiveGetFieldValueLong: %d\r\n", uaGetLastError());
                return;
                }

            // Читаем поле содержащее плавающее значение...

            bIsOk = uaArchiveGetFieldValueFloat (hArchive, ARCHIVE_FLOAT_FIELD_NUMBER, &fFloatFieldValue);

            if (!bIsOk) {
                printf
("Ошибка вызова uaArchiveGetFieldValueLong: %d\r\n", uaGetLastError());
                return;
                }

            // Читаем поле содержащее строку...

            bIsOk = uaArchiveGetFieldValueString (hArchive, ARCHIVE_STRING_FIELD_NUMBER, szStringFieldValue, STRING_FILED_SIZE);

            if (!bIsOk) {
                printf
("Ошибка вызова uaArchiveGetFieldValueString: %d\r\n", uaGetLastError());
                return;
                }
   
            // Выводим в окно диагностики прочитанные значения...

            printf
("%d: %f - %s\r\n", nIDFieldValue, fFloatFieldValue, szStringFieldValue);

            // Переходим к следующей записи...

            if (i < nCount) {

                bIsOk = uaArchiveMoveNext (hArchive);

                if (!bIsOk) {
                    printf
("Ошибка вызова uaArchiveMoveNext: %d\r\n", uaGetLastError());
                    return;
                    }
                }
            }
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