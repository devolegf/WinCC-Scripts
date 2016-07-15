/*
        (с) 2014 www.proasutp.com

        Функция возвращает имя текущего проекта.

        Вход:
                pszPrjName - указатель на строковый буфер;
                dwSizeBuf - размер строкового буфера.
        Выход:
                Если функция выполнилась успешно, тогда возвращается TRUE, иначе FALSE.

*/

#include "apdefap.h"

BOOL GetProjectName (char* pszPrjName, DWORD dwSizeBuf)
{
    CMN_ERROR dmErr;
    BOOL ret = FALSE;
    long l1 = 0, l2 = 0, i=0, j=0, c=0;
    char s[_MAX_PATH+1] = "";

    // Получаем проект который запущен в режим RunTime

    ret = DMGetRuntimeProject(s, sizeof(s), &dmErr);

    if (!ret) {
        LogMsg (dmErr.dwError1, "Не могу получить имя проекта","GetProjectName");
        return FALSE;
        }

    // Т.к. мы получаем имя проекта вида  "\\PCNAME\WinCC60_Project_ProjectName\ProjectName.mcp",
    // то необходимо выделить только имя "ProjectName"

    c = strlen
(s);

    for (i = c-1,l1 = 0,l2 = 0; i >= 0; i--)    {
        if (s[i] == '.') l2 = i;
        if (s[i] == '\\') { l1 = i; break;}
        }

    if ( (l1 == 0) || (l2 == 0) ) {
        LogMsg (dmErr.dwError1, "Ошибка выделения границ имени проекта","GetProjectName");
        return FALSE;
        }

    l1++;
 
    if ((l1-l2+1) > dwSizeBuf ) {
        LogMsg (dmErr.dwError1, "Длина имени проекта превышает размер  буфера строки szPrjName","GetProjectName");
        return FALSE;
        }

    for (i = l1, j = 0; i < l2; i++, j++) szPrjName[j] = s[i];
    szPrjName[j] = 0;
   
    return TRUE;
}