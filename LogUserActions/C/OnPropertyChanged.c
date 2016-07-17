#include "apdefap.h"
void OnPropertyChanged(char* lpszPictureName, char* lpszObjectName, char* lpszPropertyName,  char*  value)
{
// WINCC:TAGNAME_SECTION_START
// syntax: #define TagNameInAction "DMTagName"
#define v_str_length 4
// next TagID : 1
// WINCC:TAGNAME_SECTION_END

// WINCC:PICNAME_SECTION_START
// syntax: #define PicNameInAction "PictureName"
// next PicID : 1
// WINCC:PICNAME_SECTION_END

LINKINFO        pLink;
char v_str[v_str_length+1] = "";

if (PDLRTGetLink (0, lpszPictureName, lpszObjectName, "OutputValue", &pLink, NULL, NULL, NULL)) {
strncpy(v_str, value, v_str_length);

SetTagChar(pLink.szLinkName, v_str);   //Return-Type: BOOL

  n_GCreateMyOperationMsg_str( 0x00000001, 1004, lpszPictureName, lpszObjectName, pLink.szLinkName, GetPropChar(lpszPictureName, lpszObjectName, "ToolTipText"), "",GetPropChar(lpszPictureName, lpszObjectName, "OutputValue"), GetPropChar(lpszPictureName, lpszObjectName, "InputValue"), "");
}
}