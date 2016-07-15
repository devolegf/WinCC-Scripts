/*
        (с) 2015 http://www.proasutp.com
   
        Функция возвращает хендл главного окна проекта WinCC режима исполнения (runtime).
        Обычно используется для вывода окон дочерних данному окну.

*/

HWND GetWinCCRTMainWinHandle()
{
    return FindWindow("PDLRTisAliveAndWaitsForYou","WinCC-Runtime - ");
}