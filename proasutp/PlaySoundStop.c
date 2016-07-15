/*
        (с) 2014 www.proasutp.com

        Функция прекращает проигрывание wav-файл.

        Вход:
                Нет
        Выход:
                Нет

*/

void PlaySoundStop ()
{
    #pragma code ("Winmm.dll")
    long WINAPI PlaySoundA (const char * pszSound , void * hmode, DWORD dwFlag );  
    #pragma code()

    PlaySoundA(NULL,NULL,0);
}