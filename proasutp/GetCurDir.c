/*
        (с) 2014 www.proasutp.com

        Функция возвращает в буфер строки полный путь к директории проекта WinCC.

        Вход:
                pszCurDir - ссылка на буфер;
                dwSizeBuf - размер буфера;
        Выход:
                Если функция выполнилась успешно, тогда возвращается TRUE, иначе FALSE.

*/

#include "apdefap.h"

BOOL GetCurDir (char* pszCurDir, DWORD dwSizeBuf)
{
    char szProjectFile [MAX_PATH+1] = "";
    CMN_ERROR dmErr;
    DM_DIRECTORY_INFO dmDirInfo;
    BOOL ret = FALSE;
    int i = 0;

    // Получаем проект который запущен в режим RunTime
    ret = DMGetRuntimeProject(szProjectFile, sizeof (szProjectFile), &dmErr);
    if (!ret)
    {
        LogMsg (dmErr.dwError1, "Не могу получить Runtime проект", "GetCurDir");
        return FALSE;
    }

    // Получаем директорию запущенного проекта
    ret = DMGetProjectDirectory ("MyApp",szProjectFile, &dmDirInfo, &dmErr);
    if (!ret)
    {
        LogMsg (dmErr.dwError1, "Не могу получить проектную директорию", "GetCurDir");
        return FALSE;
    }

    for (i=0; i < dwSizeBuf; i++)  pszCurDir[i] = dmDirInfo.szProjectDir[i];
    pszCurDir[i] = '\0';

    return TRUE;
}