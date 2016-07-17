#include "apdefap.h"
void OnLButtonDown(char* lpszPictureName, char* lpszObjectName, char* lpszPropertyName, UINT nFlags, int x, int y)
{
// WINCC:TAGNAME_SECTION_START
// syntax: #define TagNameInAction "DMTagName"
// next TagID : 1
#define TAG_1 "Main/Total/allPlatesDel.Clr_L1"
// WINCC:TAGNAME_SECTION_END

// WINCC:PICNAME_SECTION_START
// syntax: #define PicNameInAction "PictureName"
// next PicID : 1
// WINCC:PICNAME_SECTION_END

HWND Handle=NULL;

   Handle=FindWindow("PDLRTisAliveAndWaitsForYou","WinCC-Runtime - ");
   if (MessageBox(Handle,"    Очистить все листы?    ","Очистка всех листов из ССМ",MB_YESNO|MB_ICONWARNING|MB_SYSTEMMODAL|MB_SETFOREGROUND)==6){

//  n_GCreateMyOperationMsg( DWORD dwFlags, DWORD dwMsgNum, char* lpszPictureName, char* lpszObjectName, char* lpszTagName, char* lpszObjectName1, char* lpszActionName, float doValueOld, float doValueNew, char* pszComment )
      n_GCreateMyOperationMsg( 0x00000001, 1009, lpszPictureName, lpszObjectName, TAG_1, GetPropChar(lpszPictureName, lpszObjectName, "ToolTipText"), "", 0, 1, "");

//         MessageBox(Handle,"      Готово      ","Лист удален",MB_OK|MB_ICONINFORMATION|MB_SYSTEMMODAL|MB_SETFOREGROUND);
      SetTagBit(TAG_1, 1);
   }
}
