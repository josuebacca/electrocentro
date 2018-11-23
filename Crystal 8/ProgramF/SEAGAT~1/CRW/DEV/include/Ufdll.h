/*
** File:    UFDll.h
** Author:  Rex Benning, Rick Cameron
** Date:    3 Mar 93
**
** Purpose: Declarations for a user-defined function DLL
**
** Modified : Ivanka Jankovic
** Date: 14.12.1995
** Added time and date-time types
*/

#if !defined (UFDLL_H)
#define UFDLL_H

#include "CRDLL.h"

// Set 1-byte structure alignment
#if defined (__BORLANDC__)               // Borland C/C++
  #pragma option -a-
#elif defined (_MSC_VER)                 // Microsoft Visual C++
  #if _MSC_VER >= 900                    // MSVC 2.x and later
    #pragma pack (push)
  #endif
  #pragma pack (1)
#endif 

/********************************************/


#if defined (__cplusplus)
extern "C"
{
#endif

#define UFVersion1              0x0100
#define UFVersion5              0x0500

#define UFMaxNArguments 50

#define UFMaxStringLength 254
#define UFUnknownStringLength MaxInt16u

typedef CRInt8s UFTInt8s;
typedef CRInt8u UFTInt8u;

typedef CRInt16s UFTInt16s;
typedef CRInt16u UFTInt16u;

typedef CRInt32s UFTInt32s;
typedef CRInt32u UFTInt32u;

typedef CRNumber  UFTNumber;
typedef UFTNumber UFTCurrency;

typedef CRDate UFTDate;
typedef CRTime UFTTime;
typedef CRDateTime UFTDateTime;

typedef CRBoolean UFTBoolean;

#define UFNumberScalingFactor CRNumberScalingFactor

typedef enum UFError
{
    UFUserError = -1,
    UFNoError,
    UFNoMemory,
    UFStringTooBig,
    UFDivideByZero,
    UFArrayIndexTooBig,
    UFStringIndexTooBig,
    UFNotEnoughParameters,
    UFMemCalcFail
}
    UFError;

typedef enum UFReturnTypes
{
    UFUndefinedReturn,
    UFStringReturn,
    UFNumberReturn,
    UFDateReturn,
    UFBooleanReturn,
    UFCurrencyReturn,
    UFDateRangeReturn,
    UFNumberRangeReturn,
    UFCurrencyRangeReturn,
    UFStringRangeReturn,
    UFStringArrayReturn,
    UFNumberArrayReturn,
    UFDateArrayReturn,
    UFBooleanArrayReturn,
    UFCurrencyArrayReturn,
    //UFLastReturnType
    UFTimeReturn,
    UFDateTimeReturn,
    UFTimeRangeReturn,
    UFDateTimeRangeReturn,
    UFTimeArrayReturn,
    UFDateTimeArrayReturn,
    UFLastReturnType
}
    UFReturnTypes;

typedef enum UFTypeDefs
{
    UFUndefinedType,
    UFNumber,
    UFCurrency,
    UFNumberToCurrency,
    UFInteger,
    UFBoolean,
    UFDate,
    UFStringArg,
    UFNumberRange,
    UFCurrencyRange,
    UFNumberToCurrencyRange,
    UFIntegerRange,
    UFBooleanRange, // ... for completeness, not likely used
    UFDateRange,
    UFStringRange,
    UFNumberArray,
    UFCurrencyArray,
    UFNumberToCurrencyArray,
    UFIntegerArray,
    UFBooleanArray,
    UFDateArray,
    UFStringArray,
    UFField,
    UFTime,
    UFDateTime,
    UFTimeRange,
    UFDateTimeRange,
    UFTimeArray,
    UFDateTimeArray
}
    UFTypeDefs;

typedef union UFParamUnion
{
    UFTInt32s   ParamInteger;
    UFTNumber   ParamNumber;
    UFTDate     ParamDate;
    UFTTime     ParamTime;
    UFTDateTime ParamDateTime;
    UFTBoolean  ParamBoolean;
    UFTCurrency ParamCurrency;
    char        FAR *ParamString;
    UFTNumber   FAR *ParamNumberArray;
    UFTInt32s   FAR *ParamIntegerArray;
    UFTDate     FAR *ParamDateArray;
    UFTTime     FAR *ParamTimeArray;
    UFTDateTime FAR *ParamDateTimeArray;
    UFTBoolean  FAR *ParamBooleanArray;
    UFTCurrency FAR *ParamCurrencyArray;
    char        FAR * FAR *ParamStringArray;
    void        FAR *ParamArrayPointer;

    struct
    {
        UFTInt32s ParamIntStart;
        UFTInt32s ParamIntEnd;
    }
        ParamIntRange;

    struct
    {
        UFTDate ParamDateStart;
        UFTDate ParamDateEnd;
    }
        ParamDateRange;

    struct
    {
        UFTTime ParamTimeStart;
        UFTTime ParamTimeEnd;
    }
        ParamTimeRange;

    struct
    {
        UFTDateTime ParamDateTimeStart;
        UFTDateTime ParamDateTimeEnd;
    }
        ParamDateTimeRange;

    struct
    {
        UFTNumber  ParamNumberStart;
        UFTNumber  ParamNumberEnd;
    }
        ParamNumberRange;

    struct
    {
        UFTNumber  ParamCurrencyStart;
        UFTNumber  ParamCurrencyEnd;
    }
        ParamCurrencyRange;

    struct
    {
        char FAR *ParamStringStart;
        char FAR *ParamStringEnd;
    }
        ParamStringRange;

    struct UFMemCalcParam FAR *MemCalcSubParam;
}
    UFParamUnion;

typedef struct UFParamListElement
{
    UFTInt16u StructSize;
    struct UFParamListElement FAR *NextParam;
    UFTInt16s ParamLength;
    UFTypeDefs ParamType;
    UFParamUnion Parameter;
}
    UFParamListElement;
#define UFParamListElementSize  (sizeof (UFParamListElement))

typedef enum UFMemCalcParmTypes
{
    UFLiteral,
    UFArray,
    UFVariable,
    UFExpression
}
    UFMemCalcParmTypes;

// For use when asking the DLL how much memory to allocate for a returned string or array.
typedef struct UFMemCalcParam 
{
    UFTInt16u StructSize;
    struct UFMemCalcParam FAR *NextMemCalcParam;
    UFTInt16s MemCalcParamLength; 
    UFMemCalcParmTypes MemCalcLiteralFlag;
    UFTypeDefs MemCalcParamType;
    UFParamUnion MemCalcParameter;
}
    UFMemCalcParam;
#define UFMemCalcParamSize      (sizeof (UFMemCalcParam))

typedef union UFReturnUnion
{
    UFTNumber   ReturnNumber;
    UFTDate     ReturnDate;
    UFTTime     ReturnTime;
    UFTDateTime ReturnDateTime;
    UFTBoolean  ReturnBoolean;
    UFTCurrency ReturnCurrency;
    char        FAR *ReturnString;

    struct
    {
        UFTInt32s ReturnIntStart;
        UFTInt32s ReturnIntEnd;
    }
        ReturnIntRange;

    struct
    {
        UFTDate ReturnDateStart;
        UFTDate ReturnDateEnd;
    }
        ReturnDateRange;

    struct 
    {
        UFTTime ReturnTimeStart;
        UFTTime ReturnTimeEnd;
    }
        ReturnTimeRange;

    struct
    {
        UFTDateTime ReturnDateTimeStart;
        UFTDateTime ReturnDateTimeEnd;
    }
        ReturnDateTimeRange;

    struct
    {
        UFTNumber ReturnNumberStart;
        UFTNumber ReturnNumberEnd;
    }
        ReturnNumberRange;

    struct
    {
        char FAR *ReturnStringStart;
        char FAR *ReturnStringEnd;
    }
        ReturnStringRange;

    UFTInt32s UFReturnUserError;

    UFTNumber   FAR *ReturnNumberArray;
    UFTDate     FAR *ReturnDateArray;
    UFTTime     FAR *ReturnTimeArray;
    UFTDateTime FAR *ReturnDateTimeArray;
    UFTBoolean  FAR *ReturnBooleanArray;
    UFTCurrency FAR *ReturnCurrencyArray;
    char        FAR * FAR *ReturnStringArray;
}
    UFReturnUnion;

typedef enum UFMemCalcType
{
    UFCalcStringLength,
    UFCalcArraySize
}
    UFMemCalcType;

typedef struct UFParamBlock
{
    UFTInt16u StructSize;
    UFReturnUnion ReturnValue;
    UFTInt16s ReturnValueLength;
    UFTInt16s FunctionNumber;
    void (FAR PASCAL *UFErrorHandler) (UFError);
    UFParamListElement FAR *UFParamList;
    UFTInt32u JobId;
}
    UFParamBlock;
#define UFParamBlockSize        (sizeof (UFParamBlock))

// For use when asking the DLL how much memory to allocate for a returned string or array.
typedef struct UFMemCalcBlock
{
    UFTInt16u StructSize;
    UFMemCalcType MemCalcCommand;
    UFTInt16s UFMemCalcArrayLength;
    UFTInt16s UFMemCalcStringLength;
    UFMemCalcParam FAR *UFMemCalcParamList; // For calculating string sizes
    UFParamListElement FAR *UFParamList;    // For calculating array sizes.
}
    UFMemCalcBlock;
#define UFMemCalcBlockSize      (sizeof (UFMemCalcBlock))

typedef struct UFFunctionDefStrings
{
    char FAR *FuncDefString;
    UFError (FAR PASCAL *UFEntry) (UFParamBlock FAR *);
    UFError (FAR PASCAL *UFMemCalcFunc) (UFMemCalcBlock FAR *);
}
    UFFunctionDefStrings;

typedef struct UF5FunctionDefStrings
{
    char FAR *FuncDefString;
    UFError (FAR PASCAL *UFEntry) (UFParamBlock FAR *);
    UFError (FAR PASCAL *UFMemCalcFunc) (UFMemCalcBlock FAR *);
    UFTBoolean callOnlyWhilePrintingRecords;
    UFTBoolean hasSideEffects;
}
    UF5FunctionDefStrings;

/* An encasing structure for extendability, the first element will indicate, and check, 
** which  version of the interface is in use. It is set by the compiler using the 
** sizeof operator. Any changes to the interface must change the size of this structure.
** typedef struct
*/
typedef struct UFFunctionDefStringList
{
    UFTInt16u StructSize;
    UFFunctionDefStrings FAR *UFFunctionDefStrPtr;
}
    UFFunctionDefStringList;
#define UFFunctionDefStringListSize (sizeof (UFFunctionDefStringList))

typedef struct UF5FunctionDefStringList
{
    UFTInt16u StructSize;
    UF5FunctionDefStrings FAR *UF5FunctionDefStrPtr;
}
    UF5FunctionDefStringList;
#define UF5FunctionDefStringListSize (sizeof (UF5FunctionDefStringList))

typedef struct UFFunctionTemplates
{
    char FAR *FuncTemplate;
}
    UFFunctionTemplates;

/* An encasing structure for extendability, the first element will indicate, and check,
** which version of the interface is in use. It is set by the compiler using the sizeof 
** operator. Any changes to the interface must change the size of this structure.
** typedef struct
*/
typedef struct UFFunctionTemplateList
{
    UFTInt16u StructSize;
    UFFunctionTemplates FAR *UFFuncTemplatePtr;
}
    UFFunctionTemplateList;
#define UFFunctionTemplateListSize (sizeof (UFFunctionTemplateList))

typedef struct UFFunctionExamples
{
    char FAR *FuncExample;
}
    UFFunctionExamples;

/* An encasing structure for extendability, the first element will indicate, and check,
** which version of the interface is in use. It is set by the compiler using the sizeof
** operator. Any changes to the interface must change the size of this structure.
** typedef struct
*/
typedef struct UFFunctionExampleList
{
    UFTInt16u StructSize;
    UFFunctionExamples FAR *FuncExample;
}
    UFFunctionExampleList;
#define UFFunctionExampleListSize (sizeof (UFFunctionExampleList))

#if defined (__cplusplus)
}
#endif

/********************************************/
// Reset structure alignment
#if defined (__BORLANDC__)
  #pragma option -a.
#elif defined (_MSC_VER)
  #if _MSC_VER >= 900
    #pragma pack (pop)
  #else
    #pragma pack ()
  #endif
#endif    

#endif // UFDLL_H
