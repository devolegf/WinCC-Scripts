/*
        (с) 2014 www.proasutp.com

        Функция возвращает количество миллисекунд прошедших с момента запуска ОС.

        Вход:
        Нет.
   
        Выход:
        Количество миллисекунд прошедших с момента запуска операционной системы.

*/

DWORD Ticks ()
{
    #pragma code ("kernel32.dll")
    DWORD GetTickCount(void);
    #pragma code ()

    return GetTickCount();
}