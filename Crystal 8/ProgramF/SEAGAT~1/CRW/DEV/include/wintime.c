/*
** File:    WinTime.c
** Author:  Some Microsoft Programmer(s)
** Editor:  Rick Cameron
** Date:    8 Mar 93
**
** Purpose: Time formatting functions.
**
** This code comes from an example in the Windows SDK forum on Compuserve.
** The original header comment follows.
*/

/************************************************************************
 *
 *  Module:  WINTIME.C
 *
 *  System:  Microsoft C 6.00A, Windows 3.00
 *
 * Remarks:
 *
 *    This module contains functions for managing time and date strings
 *    using internationalization.
 *
 *    The functions in this module assume that all WIN.INI international
 *    settings are valid.  If the settings are not there (which is the
 *    case until a change is made through control panel) defaults are
 *    used.
 *
 *    If the WIN.INI strings (such as the sLongDate string) is invalid
 *    in someway (i.e. it was NOT written by Control Panel) this code
 *    will probably break.
 *
 *    The ResetInternational() function reads the day and month strings
 *    from the EXE or DLL file that exports the hInst instance handle.
 *    It only reads these strings the first time it is called.
 *
 * History:
 *
 *   01/22/91  - First version.
 *   02/19/91  - Fixed bug with default settings (i.e. none) in WIN.INI
 *   03/13/91  - Fixed documentation, cleaned up code some.
 *   03/15/91  - Previous version would gp fault if sLongDate did not have
 *               single quotes around the separator strings.  Control Panel
 *               always writes with the quotes, but some other application
 *               appears to be modifying this string and does not add the
 *               quotes.  The WININI2.TXT documents sLongDate as not having
 *               quotes, so this should have been done in the first place
 *               anyway.
 *
 *   11/29/91  - Discovered that by changing the country information in
 *               control panel, it will write the longdate format withou
 *               quotes.  Also discovered that I am not
 *               parsing correctly in this case.  
 *
 ************************************************************************/

#include <windows.h>
#include <time.h>

#include "wintime.h"

#define NUM_DAYSOFWEEK     7

#define LEN_LONGWEEKDAY    13
#define LEN_SHORTWEEKDAY   5

/************************************************************************
 * Macros 
 ************************************************************************/
//
// These macros make taking apart the "tm" time structure easier.
//

#define HOUR(datetime)  (datetime -> tm_hour)
#define MIN(datetime)   (datetime -> tm_min)
#define SEC(datetime)   (datetime -> tm_sec)

extern HINSTANCE UFInstance;

/************************************************************************
 * Local Variables
 *
 * Remember:  If this code goes into a DLL these variables will reside in
 * the DLLs data segment.  Thus 1) The DLL should NOT be a CODE only DLL,
 * and 2) these variables will be shared by all applications using the DLL.
 *
 *    These variables must be initialized by calling the
 *    TimeResetInternational() function.  If you are putting this in
 *    in a DLL then call this function in your LibMain.
 *
 *    Then, in your application, call TimeResetInternational everytime
 *    you detect that the [intl] section of the WIN.INI has changed.
 *    You do this by processing the WM_WININICHANGE message and
 *    comparing (LPSTR)lParam to "intl".
 *
 ************************************************************************/
// 
// These variables are initialized when ResetInternational() is called.
// They are filled with the corresponding values in the [Intl] section
// of the WIN.INI file.
//
char  sDate[3] ;
char  sTime[3] ;
char  sAMPM[3][6] ;

int   iTime ;
int   iTLzero ;

//
// These strings contain the month and day of week strings.
// (i.e. "January", "Jan", "Sunday", "Sun")
// They are loaded from the application's (or DLL's) string table
// when ResetInternational() is called the first time.
//
char  szWday      [NUM_DAYSOFWEEK]     [LEN_LONGWEEKDAY] ;
char  szShortWday [NUM_DAYSOFWEEK]     [LEN_SHORTWEEKDAY] ;

//
// These variables are used by the functions to get the current time
// and date.
//
LONG           lTime ;
struct tm FAR  *lpTM ;

/*************************************************************************
 *  LPSTR TimeGetCurTime( LPSTR lpszTime, WORD wFlags )
 *
 *  Description: 
 *    Function returns the current time, formatted the way windows
 *    wants it!
 *
 *  Arguments:
 *    Type/Name
 *             Description
 *
 *    LPSTR lpszTime
 *             Points to a buffer that is to recieve the formatted time.
 *             Must be large enough to hold the longest date possible
 *             (64 bytes should do it).
 *
 *    WORD  wFlags
 *             Specifies flags that modify the behavior of the function.
 *             May be any of the following.
 *
 *          Flag              Meaning
 *          0                 The time will be formatted exactly the way
 *                            Control Panel does it, using Windows'
 *                            international settings. (i.e. 1:00:00 AM)
 *
 *          TIME_12HOUR       The time will be formatted in 12 hour format
 *                            regardless of the WIN.INI internationalization
 *                            settings.  This flag may not be used with the
 *                            TIME_24HOUR flag. (i.e. 2:32:10 PM)
 *                            
 *          TIME_24HOUR       The time will be formatted in 24 hour format
 *                            regardless of the WIN.INI internationalization
 *                            settings.  This flag may not be used with the
 *                            TIME_12HOUR flag. (i.e. 14:32:10)
 *
 *          TIME_NOSECONDS    The time will be formatted without seconds.
 *                            (i.e. 2:32 PM)
 *
 *  Comments:
 *
 *************************************************************************/
LPSTR TimeGetCurTime( LPSTR lpszTime, WORD wFlags )
{
   // 
   // Get current time
   //
   time( &lTime ) ;
   lpTM = localtime( &lTime ) ;

   return TimeFormatTime( lpszTime, lpTM, wFlags ) ;
}


/*************************************************************************
 *  LPSTR TimeFormatTime( LPSTR lpszTime,
 *                        struct tm FAR *lpTM,
 *                        WORD wFlags )
 *
 *  Description:
 *    Function returns the time, specified in the tm structure,
 *    formatted the way windows
 *    wants it!
 *
 *  Arguments:
 *    Type/Name
 *             Description
 *
 *    LPSTR lpszTime
 *             Points to a buffer that is to recieve the formatted time.
 *             Must be large enough to hold the longest date possible
 *             (64 bytes should do it).
 *
 *    struct tm *lpTM
 *             Pointer to a "tm" structure (from time.h) that contains
 *             the time and date you want formatted.
 *
 *    WORD  wFlags
 *             Specifies flags that modify the behavior of the function.
 *             May be any of the following.
 *
 *          Flag              Meaning
 *          0                 The time will be formatted exactly the way
 *                            Control Panel does it, using Windows'
 *                            international settings. (i.e. 1:00:00 AM)
 *
 *          TIME_12HOUR       The time will be formatted in 12 hour format
 *                            regardless of the WIN.INI internationalization
 *                            settings.  This flag may not be used with the
 *                            TIME_24HOUR flag. (i.e. 2:32:10 PM)
 *                            
 *          TIME_24HOUR       The time will be formatted in 24 hour format
 *                            regardless of the WIN.INI internationalization
 *                            settings.  This flag may not be used with the
 *                            TIME_12HOUR flag. (i.e. 14:32:10)
 *
 *          TIME_NOSECONDS    The time will be formatted without seconds.
 *                            (i.e. 2:32 PM)
 *
 *  Comments:
 *
 *************************************************************************/
LPSTR TimeFormatTime( LPSTR lpszTime, struct tm FAR *lpTM, WORD wFlags )
{
   char szFmt[32] ;
   char szTemp[32] ;

   //
   // Leading zero?
   //
   if (iTLzero)
      lstrcpy( szFmt, "%02d" ) ;
   else
      lstrcpy( szFmt, "%d" ) ;

   //
   // See if we want seconds...
   // 
   if (wFlags & TIME_NOSECONDS)
   {
      wsprintf( szTemp, "%s%s%s",
               (LPSTR)szFmt,
               (LPSTR)sTime,
               (LPSTR)"%02d" ) ;
   }
   else               // HH :MM :SS
      wsprintf( szTemp, "%s%s%s%s%s",
               (LPSTR)szFmt,
               (LPSTR)sTime,
               (LPSTR)"%02d",
               (LPSTR)sTime,
               (LPSTR)"%02d" ) ;

   //
   // see if we want 12 hour or 24 hour format.
   //
   if (wFlags & TIME_24HOUR || iTime)
   {
      if (wFlags & TIME_NOSECONDS)
         wsprintf( lpszTime, szTemp, HOUR( lpTM ), MIN( lpTM ) ) ;
      else
         wsprintf( lpszTime, szTemp, HOUR( lpTM ), MIN( lpTM ), SEC( lpTM ) ) ;
   }
   else
   {
      lstrcat( szTemp, " %s" ) ;

      if (wFlags & TIME_NOSECONDS)
         wsprintf( lpszTime, szTemp,
                   (HOUR( lpTM ) % 12) ? (HOUR( lpTM ) % 12) : 12,
                   MIN( lpTM ),
                   (LPSTR)sAMPM[HOUR( lpTM ) / 12] ) ;
      else
         wsprintf( lpszTime, szTemp,
                   (HOUR( lpTM ) % 12) ? (HOUR( lpTM ) % 12) : 12,
                   MIN( lpTM ), SEC( lpTM ),
                   (LPSTR)sAMPM[HOUR( lpTM ) / 12] ) ;
   }

   return lpszTime ;   

}/* TimeGetCurTime() */


/*************************************************************************
 *  VOID TimeResetInternational( VOID ) ;
 *
 *  Description:
 *    This function sets some local static variables to Windows' current
 *    time and date settings.
 *
 *    It should be called everytime a WM_WININICHANGE is sent for the
 *    intl section.
 *
 *    The variables here are in the DLLs data segment so if one app
 *    changes the stuff, the others will see it.
 *
 *  Comments:
 *
 *************************************************************************/
VOID TimeResetInternational( VOID )
{
   static char szIntl[] = "intl" ;

   iTime = GetProfileInt( szIntl, "iTime", 0 ) ;
   iTLzero = GetProfileInt( szIntl, "iTLzero", 1 ) ;

   GetProfileString( szIntl, "sDate", "/", sDate,  2 ) ;

   GetProfileString( szIntl, "sTime", ":", sTime,  2 ) ;

   GetProfileString( szIntl, "s1159", "AM", sAMPM[0],  5 ) ;

   GetProfileString( szIntl, "s2359", "PM", sAMPM[1],  5 ) ;

   if (*szWday[0] == '\0')
   {
      WORD i ;

      for (i = 0 ; i <= 6 ; i++)
         LoadString( UFInstance, IDS_SUNDAY + i, szWday[i], 12 ) ;

      //
      // In an international setting "Sunday" may not be abbrieviated
      // as "Sun".
      //
      for (i = 0 ; i <= 6 ; i++)
         LoadString( UFInstance, IDS_SUNDAY_SHORT + i, szShortWday[i], 4 ) ;
   }

}/* TimeResetInternational() */

