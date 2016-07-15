/*
        (с) 2014 www.proasutp.com

        Функция ищет заданный символ с конца строки.

        Вход:
        pszString - указатель на строковый буфер где будет осуществляться поиск;
        chChar - символ, поиск которого будет осуществляться.
   
        Выход:
                Если символ не найден возвращается NULL иначе возвращается указатель на найденный символ в строковом буфере

*/

#include "apdefap.h"

char* strchrback (char* pszString, char chChar)
{
    int i=0;
   
    if (pszString == NULL) {
        LogMsg ((DWORD)-1,"Передан нулевой указатель на строковый буфер в параметре pszString","strchrback");
        return NULL;
        }

    for (i=strlen
(pszString)-1; i>0; i--)
        if (pszString[i]==chChar) break;

    if (i < 0)
        return NULL;
    else
        return &pszString[i];
}