/*
    (с) 2014 www.proasutp.com

    Функция выводит сообщение в окно диагностики и в файл, если имя файла определено во внешней переменной szLogFileName.

    Вход:
        dwErrorCode - код ошибки;
        szErrorText - описание ошибки;
        pszNameFunc - название функции где возникла ошибка.
    Выход:
        Нет.

*/

#include "apdefap.h"

void LogMsg (DWORD dwErrorCode, char* szErrorText, char* pszNameFunc)
{
   
    #define BUFFER_MAX_SIZE 1000        // размер буфера вывода в файл, значение выбрано "с запасом", т.к. размер строки вывода в файл редко будет превышать указанный
   
    extern char szLogFileName[];            // внешняя глобальная переменная с именем файла, которая должна сформироваться при инициализации проекта как <ProjectName>.err
    char szTime [20] = "";          // дата и время записи сообщения
    char szMsgBuff[BUFFER_MAX_SIZE]= "";    // строковый буфер для сообщения

    FILE *file = NULL;

    if (dwErrorCode)
        printf
("Код ошибки - %d(%x);  Сообщение - %s; Функция - %s\r\n", dwErrorCode, dwErrorCode,((szErrorText!=NULL)&&(szErrorText[0]!=0)) ? szErrorText : "", ((pszNameFunc!=NULL)&&(pszNameFunc[0]!=0)) ? pszNameFunc : "");
    else
        printf
("%s; Функция - %s\r\n",((szErrorText!=NULL)&&(szErrorText[0]!=0)) ? szErrorText : "", ((pszNameFunc!=NULL)&&(pszNameFunc[0]!=0)) ? pszNameFunc : "");


    if (szLogFileName==NULL && szLogFileName[0]==0)
    {
        printf
("Сообщение не записано в лог-файл, его имя еще не определено!\r\n");
        return;
    }

    file = fopen
(szLogFileName, "a+");

    if (file)
    {
        GetTime (szTime, sizeof (szTime));

        // устанавливаем заданный размер буфера выводы, т.к. по умолчанию он часто имеет меньший размер,
        // что приводит к его переполнению и ошибке типа "access violation"

        setbuf
(file,szMsgBuff);

        fprintf
(file,"%sКод ошибки: %d(%x); Сообщение: %s; Функция: %s\n", szTime, dwErrorCode, dwErrorCode, ((szErrorText!=NULL)&&(szErrorText[0]!=0)) ? szErrorText : "", ((pszNameFunc!=NULL)&&(pszNameFunc[0]!=0)) ? pszNameFunc : "");
        fclose
(file);
    }
    else printf
("Не могу открыть файл \"%s\" для записи сообщения!\r\n", szLogFileName);
}