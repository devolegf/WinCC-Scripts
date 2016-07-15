/*
        (с) 2014 www.proasutp.com

        Функция выводит стандартное модальный диалоговое окно с сообщением

        Вход:
        pszCaption - указатель на строку заголовка окна;
        pszMessage - указатель на строку сообщения;
        dwFlags - стандартные флаги, см. MSDN, функцию WinAPI MessageBox

        Выход:
                Результат возвращаемый функцией MessageBox.

*/

int MsgBox (char* pszCaption, char* pszMessage, DWORD dwFlags)
{
    HANDLE h = NULL;

    // Получаем хендл главного окна проекта WinCC
    h = FindWindow("PDLRTisAliveAndWaitsForYou","WinCC-Runtime - ");

    // Вызываем модальное диалоговое окно с сообщением
    return MessageBox (h, pszMessage, pszCaption, dwFlags);
}