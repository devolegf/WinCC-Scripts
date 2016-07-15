/*
        (с) 2014 www.proasutp.com

        Функция определяет является ли ПК сервером WinCC.

        Вход:
        Нет

        Выход:
                TRUE - ПК выполняет фунцию сервера WinCC, FALSE - нет, это клиентский ПК.

*/

BOOL IsServer ()
{
    BOOL bIsServer = FALSE;

    bIsServer = (0 == strcmp
(GetTagChar("@local::@LocalMachineName"),GetTagChar("@@local::@ServerName")));
    return bIsServer;
}