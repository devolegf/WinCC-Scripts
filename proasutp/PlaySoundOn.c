/*
        (с) 2014 www.proasutp.com

        Функция проигрывает wav-файл.

        Вход:
                pszWaveFile - ссылка на строковый буфер содержащий полный путь к звуковому файлу;
        Выход:
                Нет

*/

void PlaySoundOn (char* pszWaveFile)
{
    #pragma code ("Winmm.dll")
    long WINAPI PlaySoundA (char* pszSound , void* hmode, DWORD dwFlag);  
    #pragma code()

    #define SND_ASYNC               0x0001      /* play asynchronously */
    #define SND_NODEFAULT           0x0002      /* silence (!default) if sound not found */
    #define SND_LOOP                0x0008      /* loop the sound until next sndPlaySound */
    #define SND_NOSTOP              0x0010      /* don't stop any currently playing sound */
    #define SND_NOWAIT      0x00002000  /* don't wait if the driver is busy */
    #define SND_FILENAME        0x00020000  /* name is file name */

   
    if ( (pszWaveFile != NULL) && (pszWaveFile[0]!= 0)) {
        int i = SND_ASYNC | SND_NOSTOP | SND_LOOP;
        PlaySoundA (pszWaveFile, NULL, i);
        }
    else {
        printf
("Failed: Wrong argument <pszWaveFile> in function PlaySoundOn!\r\n");
        }
}