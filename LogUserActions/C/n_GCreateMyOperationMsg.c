/* Автор: LexUS
Добавляет сообщение о действии оператора
DWORD dwFlags - FLAG_COMMENT_PARAMETER 0x00000001 - комментарий вводится напрямую при возникновении сообщения о действии оператора
                      FLAG_COMMENT_DIALOG 0x00000003 - вызывается диалог добавления комментария при возникновении сообщения о действии оператора
DWORD dwMsgNum - номер сообщения в Alarm Logging
char* lpszPictureName - имя формы с которой производится ввод информации
char* lpszObjectName - имя объекта в котором производится ввод информации
char* lpszTagName - имя тэга
char* lpszObjectName1 - наименование датчика, объекта
char* lpszActionName - действие
double doValueOld - предыдущее значение lpszTagName
double doValueNew - новое значение lpszTagName
char* pszComment - комментарий
*/

#pragma code ("kernel32.dll")
void GetLocalTime(LPSYSTEMTIME lpSystemTime);
BOOL GetComputerNameA(LPSTR Computername, LPDWORD size);
#pragma code()

#include "apdefap.h"

#define         FLAG_COMMENT_PARAMETER  0x00000001
#define         FLAG_COMMENT_DIALOG              0x00000003

//#define         FLAG_TEXTID_PARAMETER         0x00000100


int n_GCreateMyOperationMsg( DWORD dwFlags, DWORD dwMsgNum, char* lpszPictureName, char* lpszObjectName, char* lpszTagName, char* lpszObjectName1, char* lpszActionName, float doValueOld, float doValueNew, char* pszComment )
{

    MSG_RTDATA_INSTANCECOMMENT_STRUCT   MsgCreateEx;
    MSG_RTDATA_STRUCT                                                   MsgRTData;          // for comment dialog
    CMN_ERROR                                                                          scError;   
    int                                                                                                   iRet= FALSE;
    DWORD                                                                                     dwServiceID = 0;
    BOOL                                                                                           bOK;
    SYSTEMTIME                                                                          time;
    DWORD                                                                                     dwBufSize   = 256;

    char szComputerName[256] = "";
    char* szCurrentUser = NULL;
    char * pszServerPrefix = NULL;
//    char* pszPrefix;        // to define the type of WinCC project
//    char* lpszParent;
    char szTmp[256] = "";   // for diagnosis output

    printf("Start n_GCreateMyOperationMsg \r\n");

    //======================================
    // INIT_MESSAGE_STRUCT
    //======================================
    memset(&MsgCreateEx,0,sizeof(MsgCreateEx));
    memset(&MsgRTData,0,sizeof(MsgRTData));
    memset(&scError,0,sizeof(scError));

    GetLocalTime(&time);
    MsgCreateEx.stMsgTime = time;
    MsgRTData.stMsgTime = time;
   
    MsgCreateEx.dwMsgNr = dwMsgNum;             
    MsgRTData.dwMsgNr = dwMsgNum;
   
    MsgCreateEx.wPValueUsed = (WORD)(0x0000 );  // no real process value used
    MsgRTData.wPValueUsed = (WORD)(0x0000 );
    MsgCreateEx.wTextValueUsed  = 0x03FF;       // text values 1 .. 10 used for textblocks 1 .. 10
    MsgRTData.wTextValueUsed    = 0x03FF;       // text values 1 .. 10 used for textblocks 1 .. 10
    MsgCreateEx.dwFlags =  MSG_FLAG_TEXTVALUES;
    MsgRTData.dwFlags = MSG_FLAG_COMMENT | MSG_FLAG_TEXTVALUES;

    MsgCreateEx.dwMsgState = MSG_STATE_COME;     
    MsgRTData.dwMsgState = MSG_STATE_COME;     

    //======================================
    // INITIALIZATION PROCESS VALUE BLOCKS 1..10
    //======================================
    GetComputerNameA(szComputerName, &dwBufSize);
//    sprintf(szTmp, "Computername = %s  \r\n", szComputerName);
    strncpy ( MsgCreateEx.mtTextValue[0].szText, szComputerName, sizeof (MsgCreateEx.mtTextValue[0].szText) - 1);     // Computer Name

//    printf("Start n_GCreateMyOperationMsg \r\n");

    szCurrentUser = GetTagChar("@local::@CurrentUser");
    strncpy ( MsgCreateEx.mtTextValue[1].szText, szCurrentUser, sizeof (MsgCreateEx.mtTextValue[1].szText) - 1);     // Current User Name

    MsgCreateEx.wPValueUsed  = (WORD)(MsgCreateEx.wPValueUsed | 0x000C);
     MsgCreateEx.dPValue[2] = doValueOld;            // old value
     MsgCreateEx.dPValue[3] = doValueNew;            // new value

    strncpy ( MsgCreateEx.mtTextValue[4].szText, lpszPictureName, sizeof (MsgCreateEx.mtTextValue[4].szText) - 1);     // lpszPictureName
    strncpy ( MsgCreateEx.mtTextValue[5].szText, lpszObjectName, sizeof (MsgCreateEx.mtTextValue[5].szText) - 1);     // lpszObjectName

    strncpy ( MsgCreateEx.mtTextValue[6].szText, lpszTagName, sizeof (MsgCreateEx.mtTextValue[6].szText) - 1);     // lpszTagName
    strncpy ( MsgCreateEx.mtTextValue[7].szText, lpszActionName, sizeof (MsgCreateEx.mtTextValue[7].szText) - 1);     // lpszActionName
    strncpy ( MsgCreateEx.mtTextValue[8].szText, lpszObjectName1, sizeof (MsgCreateEx.mtTextValue[8].szText) - 1);     // lpszObjectName1
    //======================================


    //======================================
    // START_MESSAGE_SERVICE
    //======================================
    memset(&scError,0,sizeof(scError));


    // GetServerPrefix to determine MC or Server
    GetServerTagPrefix(&pszServerPrefix, NULL, NULL);   //Return-Type: void
     if (NULL == pszServerPrefix)
    {
        printf("Serverapplication or Single Client\r\n");
        bOK = MSRTStartMsgService( &dwServiceID, NULL, NULL, 0, NULL, &scError ); // activate service
    }
    else   
    {
        printf("MultiClient with Prefix : %s\r\n",pszServerPrefix);   //Return - Type :char*
        bOK = MSRTStartMsgServiceMC( &dwServiceID, NULL, NULL, 0, NULL,pszServerPrefix, &scError ); // activate service
    }

    if (bOK == FALSE)
    {
       printf("n_GCreateMyOperationMsg() - Unable to start message service! \r\n");
       sprintf(szTmp, " Error1 = 0x%0x, Errortext = %s \r\n", scError.dwError1, scError.szErrorText);
       printf(szTmp);
     
       return (-101);
    }
    //======================================


    //======================================
    // PARSE PARAMETERS
    //======================================
    if (  ( dwFlags & FLAG_COMMENT_PARAMETER )  && ( NULL != pszComment  ) )
    {
      strncpy(MsgCreateEx.szComment, pszComment, sizeof (MsgCreateEx.szComment) - 1);
      MsgCreateEx.dwFlags |= MSG_FLAG_COMMENT;
    }

    if ( dwFlags & FLAG_COMMENT_DIALOG )
      MsgCreateEx.dwFlags |= MSG_FLAG_COMMENT;
    //======================================


    //======================================
    // CREATE MESSAGE
    //======================================
    bOK = MSRTCreateMsgInstanceWithComment(dwServiceID, &MsgCreateEx, &scError) ;
    if ( TRUE == bOK)
    {
        if (FLAG_COMMENT_DIALOG == (dwFlags & FLAG_COMMENT_DIALOG) )
        {
               BOOL   bOkay;
               HWND hWnd = FindWindow(NULL, "WinCC-Runtime - ");

              memset(&scError,0,sizeof(scError));
              bOkay=  MSRTDialogComment (hWnd,  &MsgRTData,  &scError);
              if (TRUE == bOkay)
              {
                      MSG_COMMENT_STRUCT    mComment;
                      mComment.dwMsgNr = dwMsgNum;
                      mComment.stTime = time;
                      sprintf( mComment.szUser, MsgCreateEx.szUser, sizeof(mComment.szUser) - 1 );

                      memset(&scError,0,sizeof(scError));
                      bOkay = MSRTGetComment (dwServiceID, &mComment, &scError);
                      if (TRUE == bOkay)
                      {
                            strncpy(MsgCreateEx.szComment, mComment.szText, sizeof (MsgCreateEx.szComment) - 1);
                     }
              }
              else
              {
                      printf("#E201: n_GCreateMyOperationMsg()  - Error at MSRTGetComment()  szErrorText="%s" error2=%d\r\n", scError.szErrorText, scError.dwError2);
                      iRet = -201;
              }
       }
 
    }


    if(bOK == FALSE)
    {
      printf ("#E301: n_GCreateMyOperationMsg()  - Error at MSRTCreateMsgInstanceWithComment()  szErrorText="%s"\r\n", scError.szErrorText);
      iRet = -301;
    }
    //======================================


    //======================================
    // STOP_MESSAGE_SERVICE
    //======================================
    bOK= MSRTStopMsgService(    dwServiceID, &scError);
    printf("End n_GCreateMyOperationMsg \r\n");
    return (iRet);
}
