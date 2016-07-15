/*
        (с) 2014 www.proasutp.com

        Функция возвращает в строковый буфер значение переменной окружения.

        Вход:
                pszEnvVarName - указатель на имя переменной окружения;
                pszValue - указатель на строковый буфер куда будет записано значение переменной окружения;
        nSize - размер строкового буфера.
        Выход:
                Если функция выполнилась успешно, тогда возвращается TRUE, иначе FALSE.

*/

#include "apdefap.h"
BOOL GetEnvVarValue (char* pszEnvVarName, char* pszValue, long nSize)
{

    #pragma code ("kernel32.dll")
        DWORD GetEnvironmentVariableA (LPCTSTR lpName,  // address of environment variable name
                                    LPTSTR lpBuffer,    // address of buffer for variable value
                                    DWORD nSize         // size of buffer, in characters
                                   );
    #pragma code ()

    DWORD ret=0;
    char szErrText[256] = "";

    if (pszEnvVarName == NULL) {
        LogMsg ((DWORD)-1, "Передан нулевой указатель на имя переменной окружения (pszEnvVarName=NULL)", "GetEnvVarValue");
        return FALSE;
        }

    if (pszEnvVarName[0] == 0) {
        LogMsg ((DWORD) -1, "Передана пустая строка в переменной имени переменной окружения (pszEnvVarName is Empty)", "GetEnvVarValue");
        return FALSE;
        }

    if (pszValue == NULL) {
        LogMsg ((DWORD)-1, "Передан нулевой указатель на строковый буфер под значение переменной окружения (pszValue =NULL)", "GetEnvVarValue");
        return FALSE;
        }

    if (nSize <= 0) {
        LogMsg ((DWORD)-1, "Передано неправильное значение длины строкового буфера (nSize<=0)", "GetEnvVarValue");
        return FALSE;
        }

    ret = GetEnvironmentVariableA (pszEnvVarName, pszValue, nSize);

    if (ret == 0) {
        sprintf
(szErrText,"Переменная окружения с именем <%s> не найдена!", pszEnvVarName);
        LogMsg ((DWORD)-1, szErrText, "GetEnvVarValue");
        return FALSE;
        }

    if (ret > (nSize - 1)) {
        sprintf
(szErrText,"Размер буфера pszValue недостаточен для значения переменной окружения (текущий = %d), необходим буфер размером не менее %d!",nSize,ret);
        LogMsg ((DWORD)-1, szErrText, "GetEnvVarValue");
        return FALSE;
        }

    return TRUE;
}