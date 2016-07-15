/*
    (с) 2016 http://www.proasutp.com

    Функция-пример использования стандартного диалога Windows для получения полного пути к файлу, например, Excel-файлу.

    Вход:
        Нет.
    Выход:
        Нет.

*/

void GetExcelFullPathUsingStandardDialog()
{
    #pragma code("comdlg32.dll")

        // Определение структуры используемой стандартным диалогом
        // выбора файла (WinAPI)

        typedef struct tagOFN {
            DWORD lStructSize;
            HWND hwndOwner;
            HINSTANCE hInstance;
            LPCTSTR       lpstrFilter;
            LPTSTR        lpstrCustomFilter;
            DWORD         nMaxCustFilter;
            DWORD         nFilterIndex;
            LPTSTR        lpstrFile;
            DWORD         nMaxFile;
            LPTSTR        lpstrFileTitle;
            DWORD         nMaxFileTitle;
            LPCTSTR       lpstrInitialDir;
            LPCTSTR       lpstrTitle;
            DWORD         Flags;
            WORD          nFileOffset;
            WORD          nFileExtension;
            LPCTSTR       lpstrDefExt;
            LPARAM        lCustData;
            void*         lpfnHook;
            LPCTSTR       lpTemplateName;
            #if (_WIN32_WINNT >= 0x0500)
                void          *pvReserved;
                DWORD         dwReserved;
                DWORD         FlagsEx;
            #endif
        } OPENFILENAME, *LPOPENFILENAME;

        // Определение самой функции вызова стандартного диалога выбора файла (WinAPI)

        BOOL WINAPI GetOpenFileNameA(LPOPENFILENAME lpofn);

    #pragma code()

    BOOL bRet = FALSE;
    OPENFILENAME ofn;
    char* psz = NULL;
    char szFilter[] = "Файлы Excel|*.xlsx";
    char szFile[_MAX_PATH+1] = "";
    char szInitialDir[_MAX_PATH+1] = "C:\\";

    // Инициализируем память выделенную под структуру...

    memset
(&ofn, 0, sizeof (OPENFILENAME));

    // Конфигурируем стандартный диалог выбора файла...

    ofn.lStructSize = sizeof (OPENFILENAME);

    // Делаем диалог дочерним по отношению к главному окну WinCC...

    ofn.hwndOwner = FindWindow (NULL, "WinCC-Runtime - ");

    // Устанавливаем фильтр на отображение только файлов Excel в диалоге...

    ofn.lpstrFilter = szFilter;
    for (psz = szFilter; *psz; psz++) {
        if (*psz == '|') {*psz = 0; break;}
        }

    // Устанавливаем буфер сохранения полного путь к выбранному файлу...

    ofn.lpstrFile = szFile;

    // Устанавливаем размер буфера...

    ofn.nMaxFile = sizeof(szFile);

    // Устанавливаем исходную директорию, содержимое которого будет отображаться
    // при первом открытии диалога...

    ofn.lpstrInitialDir = szInitialDir;

    // Открываем диалог...

    bRet = GetOpenFileNameA(&ofn);
    if (bRet == FALSE){
        printf
("Произошла ошибка, либо пользователь нажал кнопку отмены.\r\n");
        return;
        }
       
    printf
("Выбранный полный путь к файлу: \r\n%s\r\n", ofn.lpstrFile);

}