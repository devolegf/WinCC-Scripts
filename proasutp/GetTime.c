/*
        (с) 2014 www.proasutp.com

        Функция возвращает в строковый буфер текущее время в формате: dd.mm.yyyy hh:mm:ss

        Вход:
                pszTime - указатель на буфер;
                dwSize - размер буфера.
        Выход:
                Если функция выполнилась успешно, тогда возвращается TRUE, иначе FALSE.

*/

void GetTime (char* pszTime, DWORD dwSize)
{
    time_t t = 0;
    char *pDateTime = NULL;
    char Day [4] = "";
    char Date [3] = "";
    char Month [4] = "";
    char Year [5] = "";
    char Time [9] = "";
    char tm [20] = "";
    int i = 0;

    time
(&t);
    pDateTime = ctime
(&t);

    sscanf
(pDateTime, "%s %s %s %s %s", Day, Month, Date, Time, Year);

    if (strcmp
(Month,"Jan") == 0) strcpy
(Month, "01");
    if (strcmp
(Month,"Feb") == 0) strcpy
(Month, "02");
    if (strcmp
(Month,"Mar") == 0) strcpy
(Month, "03");
    if (strcmp
(Month,"Apr") == 0) strcpy
(Month, "04");
    if (strcmp
(Month,"May") == 0) strcpy
(Month, "05");
    if (strcmp
(Month,"Jun") == 0) strcpy
(Month, "06");
    if (strcmp
(Month,"Jul") == 0) strcpy
(Month, "07");
    if (strcmp
(Month,"Aug") == 0) strcpy
(Month, "08");
    if (strcmp
(Month,"Sep") == 0) strcpy
(Month, "09");
    if (strcmp
(Month,"Oct") == 0) strcpy
(Month, "10");
    if (strcmp
(Month,"Nov") == 0) strcpy
(Month, "11");
    if (strcmp
(Month,"Dec") == 0) strcpy
(Month, "12");

    sprintf
(tm, "%s.%s.%s %s ", Date, Month, Year, Time);

    for (i=0; i<dwSize; i++) pszTime[i] = tm[i];
    pszTime[i] = '\0';
}