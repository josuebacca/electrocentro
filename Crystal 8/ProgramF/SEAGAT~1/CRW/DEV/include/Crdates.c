/*
** File:    CRDates.c
**
** Editor:  Ron Hayter
** Date:    93-06-22
**
** Purpose: Functions to convert to and from Crystal Reports' date values.
**
** Copyright (c) 1993  Seagate Software Information Management Group, Inc.
*/

#include "CRDates.h"

static CRInt16u DaysInMonth [13] =
{0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};

static CRInt32s YMDToJulian (CRInt16s year,
                             CRInt16u month,
                             CRInt16u day);

CRDate CRYearMonthDayToDate (CRInt16s year,
                             CRInt16u month,
                             CRInt16u day)
{
    CRBoolean leapYear;

    if (year < -4713)
        return CRNullDate;

    leapYear =
        year >= 1582 ?
            (year % 4 == 0 && year % 100 != 0 || year % 400 == 0) :
            (year % 4 == 0);

    if (month < 1 || month > 12)
        return CRNullDate;

    DaysInMonth [2] = leapYear ? 29 : 28;
    if (day < 1 || day > DaysInMonth [month])
        return CRNullDate;

    return (CRDate) YMDToJulian (year, month, day);
}

static CRInt32s YMDToJulian (CRInt16s year,
                             CRInt16u month,
                             CRInt16u day)
{
    CRInt16s a, b = 0;
    CRInt16u work_month = month, work_day = day;
    CRInt16s work_year = year;

    /*  correct for negative year */
    if (work_year < 0)
        work_year++;

    if (work_month <= 2)
    {
        work_year--;
        work_month += 12;
    }

    /*  deal with Gregorian calendar */
    if (work_year * 10000. + work_month * 100. + work_day >= 15821015.)
    {
        a = work_year / 100.;
        b = 2 - a + a / 4;
    }

    return (CRInt32s) (365.25 * work_year) +
           (CRInt32s) (30.6001 * (work_month + 1)) +
           work_day + 1720994L + b;
}

void CRDateToYearMonthDay (CRDate date,
                           CRInt16s *year,
                           CRInt16u *month,
                           CRInt16u *day)
{
    long a, b, c, d, e, z, alpha;
    z = date + 1;

    /* dealing with Gregorian calendar reform */
    if (z < 2299161L)
        a = z;
    else
    {
        alpha = (long) ((z-1867216.25) / 36524.25);
        a = z + 1 + alpha - alpha / 4;
    }
    b = (a > 1721423L ? a + 1524 : a + 1158);
    c = (long) ((b - 122.1) / 365.25);
    d = (long) (365.25 * c);
    e = (long) ((b - d) / 30.6001);
    *day = (CRInt16u) (b - d - (long) (30.6001 * e));
    *month = (CRInt16u) ((e < 13.5) ? e - 1 : e - 13);
    *year = (CRInt16s) ((*month > 2.5 ) ? (c - 4716) : c - 4715);
}

CRTime CRHourMinuteSecondToTime (CRInt16u hour,
                                 CRInt16u minute,
                                 CRInt16u second)
{
    return (CRTime) hour * 3600 + minute * 60 + second;
}

void CRTimeToHourMinuteSecond (CRTime time,
                               CRInt16u *hour,
                               CRInt16u *minute,
                               CRInt16u *second)
{
    *hour = (CRInt16u) (time / 3600);
    *minute = (CRInt16u) (time / 60 % 60);
    *second = (CRInt16u) (time % 60);
}
