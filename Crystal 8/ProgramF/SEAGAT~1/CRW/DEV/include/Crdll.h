/*
** File:    CRDLL.h
**
** Author:  Ron Hayter
** Date:    93-06-22
**
** Purpose: Declarations common to DLLs created for Crystal Reports.
**
** Copyright (c) 1993-1996  Crystal Services
*/

#if !defined (CRDLL_H)
#define CRDLL_H  

#if defined (WIN32)
  #define CR_CALLBACK  __stdcall
  #define CR_EXPORT    __stdcall
#else
  #define CR_CALLBACK  __far __export __pascal
  #define CR_EXPORT    __far __export __pascal
#endif

#if defined (__cplusplus)
extern "C"
{
#endif

#if defined (__cplusplus)
  #define ExternC extern "C"
#else
  #define ExternC
#endif



typedef signed   char  CRInt8s;
typedef unsigned char  CRInt8u;

typedef signed   short CRInt16s;
typedef unsigned short CRInt16u;

typedef signed   long  CRInt32s;
typedef unsigned long  CRInt32u;

typedef double CRNumber;
typedef CRNumber CRCurrency;
#define CRNumberScalingFactor ((CRNumber) 100.00)

typedef CRInt32s CRDate;
#define CRNullDate ((CRDate) -1)
typedef CRInt32s CRTime;
#define CRNullTime ((CRTime) -1)

typedef struct 
{
    CRDate date;
    CRTime time;
}   
    CRDateTime;

typedef CRInt16s CRBoolean;

typedef CRInt32u CRVersion;

// Macros to encode and decode the Crystal Reports version number:
#define CR_VERSION(major,minor)     ((((CRVersion) (CRInt16u) (major)) << 16) | ((CRInt16u) (minor)))
#define CR_MAJOR_VERSION(crVersion) ((CRInt16u) (((CRVersion) (crVersion)) >> 16))
#define CR_MINOR_VERSION(crVersion) ((CRInt16u) (((CRVersion) (crVersion)) & 0x0000FFFFUL))

#if defined (__cplusplus)
}
#endif

#endif // CRDLL_H
