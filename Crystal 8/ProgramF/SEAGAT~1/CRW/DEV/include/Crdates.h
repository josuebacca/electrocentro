/*
** File:    CRDates.h
**
** Author:  Ron Hayter
** Date:    93-06-22
**
** Purpose: Declarations for functions to convert to and from Crystal
**          Reports' date values.
**
** Copyright (c) 1993  Seagate Software Information Management Group, Inc.
*/

#if !defined (CRDATES_H)
#define CRDATES_H

#include "CRDLL.h"

#if defined (__cplusplus)
extern "C"
{
#endif

extern CRDate CRYearMonthDayToDate (CRInt16s year,
                                    CRInt16u month,
                                    CRInt16u day);
extern void CRDateToYearMonthDay (CRDate date,
                                  CRInt16s *year,
                                  CRInt16u *month,
                                  CRInt16u *day);

extern CRTime CRHourMinuteSecondToTime (CRInt16u hour,
                                        CRInt16u minute,
                                        CRInt16u second);
extern void CRTimeToHourMinuteSecond (CRTime time,
                                      CRInt16u *hour,
                                      CRInt16u *minute,
                                      CRInt16u *second);

#if defined (__cplusplus)
}
#endif

#endif // CRDATES_H
