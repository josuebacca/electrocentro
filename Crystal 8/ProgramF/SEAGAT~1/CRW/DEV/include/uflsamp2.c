// file UFLSamp2.c
// Rex Benning 92 Dec 28
// This file is a sample of some user defined functions,
// to be used for debugging and as a template for real
// functions.

#include <Windows.h>
#include <UFDll.h>
#include "UFMain.h"
#include "UFUser.h"
#include "UFJob.h"

#include <StdLib.h>

// #define DEBUG_TOPN
// #define DEBUG_INSERT

#define MaxN 10


ExternC UFError FAR _export PASCAL TopNStart (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL TopNStore (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL TopNGet (UFParamBlock * ParamBlock);

ExternC UFError FAR _export PASCAL TestErrorReturn (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL TestReplString (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL TestReplStringRetSize (UFMemCalcBlock * ParamBlock);
ExternC UFError FAR _export PASCAL TestReplChars (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL TestTwoStrings (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL TestNumberArray (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL TestCurrencyArray (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL TestDateArray (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFMakeNumberRange (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFMakeDateRange (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFSumNumberArray (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFAvgNumberArray (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFStringLength (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFArrayMax (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL TestStringArray (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL TestStringArrayLength (UFMemCalcBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFSumCurrencyArray (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFSum10Numbers (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFConcat10Str (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFConcat10StrLengths (UFMemCalcBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFTestErrors (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFSumStringLengths (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFArrayElementsInRange (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFStringSubstitutions (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFStringSubMemCalc (UFMemCalcBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFMakeNumberArray (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFNumberArrayLength (UFMemCalcBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFMakeStringArray (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFStringArrayLength (UFMemCalcBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFMakeDateArray (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFDateArrayLength (UFMemCalcBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFMakeCurrencyArray (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFCurrencyArrayLength (UFMemCalcBlock * ParamBlock);
ExternC UFError FAR _export PASCAL UFMakeStringRange (UFParamBlock * ParamBlock);
ExternC UFError FAR _export PASCAL MakeStringRangeFromArray (UFParamBlock * ParamBlock);





UFFunctionDefStrings  FunctionDefStrings [] =
{
        {"Number TopNStart (Number)", TopNStart},
        {"Number TopNStore (Number)", TopNStore},
        {"Number TopNGet   (Number)", TopNGet},

        {"String TestReplString (String, Number)",TestReplString, TestReplStringRetSize},
        {"String TestReplChars (String, Number)", TestReplChars, TestReplStringRetSize},
        {"Number TestNumberArray (Array(Number), Number)", TestNumberArray, NULL},
        {"Currency TestCurrencyArray ( Array  ( Currency ), Number)", TestCurrencyArray, NULL},
        {"Date TestDateArray(Array(Date),Number)", TestDateArray, NULL},
        {"Range (Number) UFMakeNumberRange(Number, Number)", UFMakeNumberRange, NULL},
        {"Range ( Date ) UFMakeDateRange ( Date , Date)", UFMakeDateRange, NULL},
        {"Number UFSumNumberArray (Array (Number))", UFSumNumberArray, NULL},
        {"Number UFAvgNumberArray (Array (Number))", UFAvgNumberArray, NULL},
        {"Number UFStringLength (String)", UFStringLength, NULL},
        {"Number UFArrayMax (  Array (Number))", UFArrayMax, NULL},
        {"  String TestStringArray(Array(String),Number)", TestStringArray, TestStringArrayLength},
        {"  Currency UFSumCurrencyArray (Array(Currency))", UFSumCurrencyArray, NULL},
        {"Number UFSum10Numbers (number,number,number,number,number,number,number,number,number,number)",
                 UFSum10Numbers, NULL},
        {"String UFConcat10Str(string,string,string,string,string,string,string,string,string,string)",
                 UFConcat10Str, UFConcat10StrLengths},
        {"Number UFTestErrors( Number)", UFTestErrors, NULL},
        {"  Number UFSumStringLengths (Array(String))", UFSumStringLengths, NULL},
        {"Boolean UFArrayElementsInRange (Array (Number=>Currency), Range (Number=>Currency), Number)",
                  UFArrayElementsInRange, NULL},
        {"String UFStringSubstitutions (String, Array(String))", UFStringSubstitutions, UFStringSubMemCalc},
        {"Array (Number) UFMakeNumberArray (Number, Number, Number, Number)", UFMakeNumberArray, UFNumberArrayLength},
        {"Array (String) UFMakeStringArray (String, String, String, String)", UFMakeStringArray, UFStringArrayLength},
        {"Array (Date) UFMakeDateArray (Date, Date)", UFMakeDateArray, UFDateArrayLength},
        {"Array (Currency) UFMakeCurrencyArray (Currency, Currency, Currency, Currency, Currency, Currency)", UFMakeCurrencyArray, UFCurrencyArrayLength},
        {"Range (String) UFMakeStringRange(String, String)", UFMakeStringRange, NULL},
        {"Range (String) MakeStringRangeFromArray (Array (String), Number, Number)", MakeStringRangeFromArray, NULL},
        {NULL, NULL, NULL}
};



UFFunctionTemplates FunctionTemplates [] =
{
       {"TopNStart (!)"},
       {"TopNStore (!)"},
       {"TopNGet (!)"},

       {"TestReplString (!, )"},
       {"TestReplChars (!, )"},
       {"TestNumberArray (!, )"},
       {"TestCurrencyArray (!, )"},
       {"TestDateArray (!, )"},
       {"UFMakeNumberRange (!, )"},
       {"UFMakeDateRange (!, )"},
       {"UFSumNumberArray (!)"},
       {"UFAvgNumberArray (!)"},
       {"UFStringLength (!)"},
       {"UFArrayMax (!)"},
       {"TestStringArray (!, )"},
       {"UFSumCurrencyArray (!)"},
       {"UFSum10Numbers (!,,,,,,,,,)"},
       {"UFConcat10Str (!,,,,,,,,,)"},
       {"UFTestErrors (!)"},
       {"UFSumStringLengths (!)"},
       {"UFArrayElementsInRange (!, , )"},
       {"UFStringSubstitutions (!, )"},
       {"UFMakeNumberArray (!,,,)"},
       {"UFMakeStringArray (!,,,)"},
       {"UFMakeDateArray (!,)"},
       {"UFMakeCurrencyArray (!,,,,,)"},
       {"UFMakeStringRange (!,)"},
       {"MakeStringRangeFromArray (!,,)"},
       {NULL}
};



UFFunctionExamples FunctionExamples [] =
{
       {"\tTopNStart (#values)"},
       {"\tTopNStore (number)"},
       {"\tTopNGet (value#)"},

       {"\tTestReplString (String, Number)"},
       {"\tTestReplChars (String, Number)"},
       {"\tTestNumberArray (NumberArray, Number)"},
       {"\tTestCurrencyArray (CurrencyArray, Number)"},
       {"\tTestDateArray (DateArray, Number)"},
       {"\tUFMakeNumberRange (StartNumber, EndNumber)"},
       {"\tUFMakeDateRange (StartDate, EndDate)"},
       {"\tUFSumNumberArray (NumberArray)"},
       {"\tUFAvgNumberArray (NumberArray)"},
       {"\tUFStringLength (String)"},
       {"\tUFArrayMax (NumberArray)"},
       {"\tTestStringArray (StringArray, Number)"},
       {"\tUFSumCurrencyArray (CurrencyArray)"},
       {"\tUFSum10Numbers (Number,Number,Number ... )"},
       {"\tUFConcat10Str (String,String,String ... )"},
       {"\tUFTestErrors (ErrorNumber)"},
       {"\tUFSumStringLengths (StringArray)"},
       {"\tUFArrayElementsInRange (CurrencyOrNumber Array, CurrencyOrNumber Range, Count)"},
       {"\tUFStringSubstitutions (String, StringArray)"},
       {"\tUFMakeNumberArray (Number, Number, Number, Number)"},
       {"\tUFMakeStringArray (String, String, String, String)"},
       {"\tUFMakeDateArray (Date, Date)"},
       {"\tUFMakeCurrencyArray (Currency, Currency, Currency, Currency, Currency, Currency)"},
       {"\tUFMakeStringRange (String, String)"},
       {"\tMakeStringRangeFromArray (StringArray, Number, Number)"},
       {NULL}
};



#define ERR_NOUTOFRANGE     7
#define ERR_INDEXOUTOFRANGE 8

char *ErrorTable [] =
{
    "Error Message number zero",
    "Null String Found",
    "Parameter seems to have disappeared",
    "Error Message number three",
    "Error Message number four",
    "Last Error Message",
    "Invalid error number",

    "N must be between 1 and 10.",
    "The index must be between 1 and the number of values specified"
};

void InitJob (struct JobInfo *jobInfo)
{
}

void TermJob (struct JobInfo *jobInfo)
{
    if (jobInfo->data != 0)
        free (jobInfo->data);
}

/*
** The TopN* functions
*/

// the structure we use to store the top N values
struct TopN
{
    UFTInt16u nMax,    // the N
              nUsed;   // the actual number of data stored

    UFTNumber values [1];
};

ExternC UFError FAR _export PASCAL TopNStart (UFParamBlock * ParamBlock)
{
    // set up to store numbers
    struct JobInfo *jobInfo;

    jobInfo = FindJobInfo (ParamBlock->JobId);

    if (jobInfo != 0 && jobInfo->data == 0)
    {
        UFParamListElement * ParamPtr;
        UFTInt16s N;
        struct TopN *topN;

        ParamPtr = GetParam (ParamBlock, 1);
        if (ParamPtr == NULL)
            return UFNotEnoughParameters;

        N = ParamPtr->Parameter.ParamNumber / UFNumberScalingFactor;

#if defined (DEBUG_TOPN)
        {
            char buffer [80];
            sprintf (buffer, "parameter 1: %g (%d)",
                             ParamPtr->Parameter.ParamNumber,
                             N);
            MessageBox (0, buffer, "CRSampl1", MB_OK);
        }
#endif

        if (N <= 0 || N > MaxN)
        {
            ParamBlock->ReturnValue.UFReturnUserError = ERR_NOUTOFRANGE;
            return UFUserError;
        }

        topN = malloc (sizeof (struct TopN) + (N - 1) * sizeof (UFTNumber));

        if (topN == 0)
            return UFNoMemory;

        topN->nMax = N;
        topN->nUsed = 0;

        jobInfo->data = topN;

        ParamBlock->ReturnValue.ReturnNumber = N * UFNumberScalingFactor;
    }
    else
        ParamBlock->ReturnValue.ReturnNumber = 0;

    return UFNoError;
}

static void insertDatum (struct TopN *topN, UFTNumber datum)
{
    UFTInt16u index;

#if defined (DEBUG_INSERT)
    char buffer [80];
    sprintf (buffer, "starting insert\n"
                     "nUsed: %d, nMax: %d, datum: %g",
                     topN->nUsed,
                     topN->nMax,
                     datum);
    MessageBox (0, buffer, "CRSampl1", MB_OK);
#endif

    for (index = 0; index < topN->nUsed; index++)
    {
        if (topN->values [index] < datum)
            break;
        else if (topN->values [index] == datum)
            // don't store duplicates
            return;
    }

#if defined (DEBUG_INSERT)
    sprintf (buffer, "after search\n"
                     "index: %d",
                     index);
    MessageBox (0, buffer, "CRSampl1", MB_OK);
#endif

    if (index < topN->nMax)
    {
        UFTInt16s copyIndex;

        if (topN->nUsed < topN->nMax)
            ++topN->nUsed;

        for (copyIndex = topN->nUsed - 1; copyIndex > index; copyIndex--)
             topN->values [copyIndex] = topN->values [copyIndex - 1];

        topN->values [index] = datum;
    }
}

ExternC UFError FAR _export PASCAL TopNStore (UFParamBlock * ParamBlock)
{
    // store a number
    struct JobInfo *jobInfo;

    jobInfo = FindJobInfo (ParamBlock->JobId);

    if (jobInfo != 0 && jobInfo->data != 0)
    {
        struct TopN *topN;
        UFParamListElement * ParamPtr;
        UFTNumber datum;

        topN = jobInfo->data;

        ParamPtr = GetParam (ParamBlock, 1);
        if (ParamPtr == NULL)
            return UFNotEnoughParameters;

        datum = ParamPtr->Parameter.ParamNumber;

        insertDatum (topN, datum);

        ParamBlock->ReturnValue.ReturnNumber = datum;
    }
    else
        ParamBlock->ReturnValue.ReturnNumber = 0;

    return UFNoError;
}

ExternC UFError FAR _export PASCAL TopNGet (UFParamBlock * ParamBlock)
{
    // return the nth value
    struct JobInfo *jobInfo;

    jobInfo = FindJobInfo (ParamBlock->JobId);

    if (jobInfo != 0 && jobInfo->data != 0)
    {
        struct TopN *topN;
        UFParamListElement * ParamPtr;
        UFTInt16s index;

        topN = jobInfo->data;

        ParamPtr = GetParam (ParamBlock, 1);
        if (ParamPtr == NULL)
            return UFNotEnoughParameters;

        index = ParamPtr->Parameter.ParamNumber / UFNumberScalingFactor;

        if (index < 1 || index > topN->nMax)
        {
            ParamBlock->ReturnValue.UFReturnUserError = ERR_INDEXOUTOFRANGE;
            return UFUserError;
        }

#if defined (DEBUG_INSERT)
        {
            char buffer [80];
            sprintf (buffer, "fetching datum\n"
                             "index: %d, nUsed: %d",
                             index,
                             topN->nUsed);
            MessageBox (0, buffer, "CRSampl1", MB_OK);
        }
#endif

        if (index <= topN->nUsed)
            ParamBlock->ReturnValue.ReturnNumber = topN->values [index - 1];
        else
            ParamBlock->ReturnValue.ReturnNumber = 0;
    }
    else
        ParamBlock->ReturnValue.ReturnNumber = 0;

    return UFNoError;
}


/*
************************************************************
**
** The following are various test functions.  They don't do anything
** particularly useful.
*/


ExternC UFError FAR _export PASCAL TestErrorReturn (UFParamBlock *ParamBlock)
{
    ParamBlock->ReturnValue.ReturnDate = 3;

    return (UFNoMemory);
}

ExternC UFError FAR _export PASCAL TestReplString (UFParamBlock *ParamBlock)
{
    UFTNumber i;

    ParamBlock->ReturnValue.ReturnString[0] = '\0';
    for ( i = ParamBlock->UFParamList->NextParam->Parameter.ParamNumber/UFNumberScalingFactor; i > 0; --i)
        lstrcat (ParamBlock->ReturnValue.ReturnString,
                 ParamBlock->UFParamList->Parameter.ParamString);

    return (UFNoError);
}

ExternC UFError FAR _export PASCAL TestReplStringRetSize (UFMemCalcBlock * ParamBlock)
{
    UFMemCalcParam * ParamPtr;
    UFTInt16s x, y;

    if ((ParamPtr =  GetMemCalcParam (ParamBlock, 2)) == NULL)
    {
        ParamBlock->UFMemCalcStringLength = 99;
        return (UFNoError);
    }

    if (ParamPtr->MemCalcParamType != UFNumber)
    {
        ParamBlock->UFMemCalcStringLength = 98;
        return (UFNoError);
    }

    if (ParamPtr->MemCalcLiteralFlag == UFLiteral)
        x = (ParamPtr->MemCalcParameter.ParamNumber/UFNumberScalingFactor);
    else
    {
        ParamBlock->UFMemCalcStringLength = UFMaxStringLength;
        return (UFNoError);
    }

    if ((ParamPtr =  GetMemCalcParam (ParamBlock, 1)) == NULL)
    {
        ParamBlock->UFMemCalcStringLength = 97;
        return (UFNoError);
    }

    if (ParamPtr->MemCalcParamType != UFStringArg)
    {
        ParamBlock->UFMemCalcStringLength = 96;
        return (UFNoError);
    }

    ParamBlock->UFMemCalcStringLength = (x * (ParamPtr->MemCalcParamLength-1))+1;

    return (UFNoError);
}

ExternC UFError FAR _export PASCAL TestReplChars (UFParamBlock *ParamBlock)
{
    int i, j, k, l;

    ParamBlock->ReturnValue.ReturnString[0] = '\0';
    k = lstrlen(ParamBlock->UFParamList->Parameter.ParamString);
    l = ParamBlock->UFParamList->NextParam->Parameter.ParamNumber/UFNumberScalingFactor;
    for ( j = 0; j < k; ++j)
        for ( i = 0; i <l ; ++i)
            ParamBlock->ReturnValue.ReturnString[(j*l)+i] =
                ParamBlock->UFParamList->Parameter.ParamString[j];

    ParamBlock->ReturnValue.ReturnString[(k*l)] = '\0';

    return (UFNoError);
}

ExternC UFError FAR _export PASCAL TestTwoStrings (UFParamBlock *ParamBlock)
{
    ParamBlock->ReturnValue.ReturnString[0] = '\0';

    lstrcat (ParamBlock->ReturnValue.ReturnString,
             ParamBlock->UFParamList->Parameter.ParamString);
    lstrcat (ParamBlock->ReturnValue.ReturnString,
             ParamBlock->UFParamList->NextParam->Parameter.ParamString);

    return (UFNoError);
}

ExternC UFError FAR _export PASCAL TestNumberArray (UFParamBlock * ParamBlock)
{
    int index = (int)(ParamBlock->UFParamList->NextParam->Parameter.ParamNumber/UFNumberScalingFactor);

    if (index < 1 || index > ParamBlock->UFParamList->ParamLength)
        return UFArrayIndexTooBig;

    ParamBlock->ReturnValue.ReturnNumber
        = ParamBlock->UFParamList->Parameter.ParamNumberArray[index - 1];

    return UFNoError;
}

ExternC UFError FAR _export PASCAL TestCurrencyArray (UFParamBlock * ParamBlock)
{
    int index = (int)(ParamBlock->UFParamList->NextParam->Parameter.ParamNumber/UFNumberScalingFactor);

    if (index < 1 || index > ParamBlock->UFParamList->ParamLength)
        return UFArrayIndexTooBig;

    ParamBlock->ReturnValue.ReturnCurrency
        = ParamBlock->UFParamList->Parameter.ParamCurrencyArray[index - 1];

    return UFNoError;
}

ExternC UFError FAR _export PASCAL TestDateArray (UFParamBlock * ParamBlock)
{
    int index = (int)(ParamBlock->UFParamList->NextParam->Parameter.ParamNumber/UFNumberScalingFactor);

    if (index < 1 || index > ParamBlock->UFParamList->ParamLength)
        return UFArrayIndexTooBig;

    ParamBlock->ReturnValue.ReturnDate
        = ParamBlock->UFParamList->Parameter.ParamDateArray[index - 1];

    return UFNoError;
}

ExternC UFError FAR _export PASCAL TestStringArray (UFParamBlock * ParamBlock)
{
    int index = (int)(GetParam(ParamBlock,2)->Parameter.ParamNumber/UFNumberScalingFactor);

    if (index < 1 || index > ParamBlock->UFParamList->ParamLength)
        return UFArrayIndexTooBig;

    lstrcpy (ParamBlock->ReturnValue.ReturnString,
             GetParam(ParamBlock, 1)->Parameter.ParamStringArray[index - 1]);

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFMakeNumberRange (UFParamBlock * ParamBlock)
{
    UFParamListElement * ParamPtr;

    if ((ParamPtr = GetParam(ParamBlock, 1)) == NULL)
        return UFNotEnoughParameters;

    ParamBlock->ReturnValue.ReturnNumberRange.ReturnNumberStart = ParamPtr->Parameter.ParamNumber;

    if ((ParamPtr = GetParam(ParamBlock, 2)) == NULL)
        return UFNotEnoughParameters;

    ParamBlock->ReturnValue.ReturnNumberRange.ReturnNumberEnd = ParamPtr->Parameter.ParamNumber;

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFMakeStringRange (UFParamBlock * ParamBlock)
{
    UFParamListElement * ParamPtr;

    if ((ParamPtr = GetParam(ParamBlock, 1)) == NULL)
        return UFNotEnoughParameters;

    lstrcpy (ParamBlock->ReturnValue.ReturnStringRange.ReturnStringStart, ParamPtr->Parameter.ParamString);

    if ((ParamPtr = GetParam(ParamBlock, 2)) == NULL)
        return UFNotEnoughParameters;

    lstrcpy (ParamBlock->ReturnValue.ReturnStringRange.ReturnStringEnd, ParamPtr->Parameter.ParamString);

    return UFNoError;
}

ExternC UFError FAR _export PASCAL MakeStringRangeFromArray (UFParamBlock * ParamBlock)
{
    UFParamListElement * ParamPtr, * IndexPtr;
    int i,j;

    if ((ParamPtr = GetParam(ParamBlock, 1)) == NULL)
        return UFNotEnoughParameters;

    if ((IndexPtr = GetParam(ParamBlock, 2)) == NULL)
        return UFNotEnoughParameters;

    i = (int) IndexPtr->Parameter.ParamNumber / UFNumberScalingFactor;
    if ((IndexPtr = GetParam(ParamBlock, 3)) == NULL)
        return UFNotEnoughParameters;

    j = (int) IndexPtr->Parameter.ParamNumber / UFNumberScalingFactor;
    if ((i > ParamPtr->ParamLength ) || (i < 1) || ( j < 1) ||
        (j > ParamPtr->ParamLength ))
        return UFArrayIndexTooBig;

    lstrcpy (ParamBlock->ReturnValue.ReturnStringRange.ReturnStringStart, ParamPtr->Parameter.ParamStringArray[i-1]);
    lstrcpy (ParamBlock->ReturnValue.ReturnStringRange.ReturnStringEnd, ParamPtr->Parameter.ParamStringArray[j-1]);

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFMakeDateRange (UFParamBlock * ParamBlock)
{
    UFParamListElement * ParamPtr;

    if ((ParamPtr = GetParam(ParamBlock, 1)) == NULL)
        return UFNotEnoughParameters;

    ParamBlock->ReturnValue.ReturnDateRange.ReturnDateStart = ParamPtr->Parameter.ParamDate;

    if ((ParamPtr = GetParam(ParamBlock, 2)) == NULL)
        return UFNotEnoughParameters;

    ParamBlock->ReturnValue.ReturnDateRange.ReturnDateEnd = ParamPtr->Parameter.ParamDate;

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFSumNumberArray (UFParamBlock * ParamBlock)
{
    UFTInt16s i;
    UFParamListElement * ParamPtr = GetParam (ParamBlock, 1);

    if (ParamPtr == NULL)
        return UFNotEnoughParameters;

    ParamBlock->ReturnValue.ReturnNumber = 0;
    for (i = 0; i < ParamPtr->ParamLength; ++i)
        ParamBlock->ReturnValue.ReturnNumber += ParamPtr->Parameter.ParamNumberArray[i];

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFSum10Numbers (UFParamBlock * ParamBlock)
{
    UFTInt16s i;
    UFParamListElement * ParamPtr;

    ParamBlock->ReturnValue.ReturnNumber = 0;
    for (i = 1; i <= 10; ++i)
    {
        ParamPtr = GetParam (ParamBlock, i);
        if (ParamPtr == NULL)
            return UFNotEnoughParameters;

        ParamBlock->ReturnValue.ReturnNumber += ParamPtr->Parameter.ParamNumber;
    }

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFConcat10Str (UFParamBlock * ParamBlock)
{
    UFTInt16s i;
    UFParamListElement * ParamPtr;

    ParamBlock->ReturnValue.ReturnString[0] = '\0';
    for (i = 1; i <= 10; ++i)
    {
        ParamPtr = GetParam (ParamBlock, i);

        if (ParamPtr == NULL)
            return UFNotEnoughParameters;

        if (lstrlen (ParamBlock->ReturnValue.ReturnString) + lstrlen (ParamPtr->Parameter.ParamString)
            > UFMaxStringLength)
           return UFStringTooBig;

        lstrcat (ParamBlock->ReturnValue.ReturnString, ParamPtr->Parameter.ParamString);
    }

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFConcat10StrLengths (UFMemCalcBlock * ParamBlock)
{
    UFTInt16s i, j = 0;
    UFMemCalcParam * ParamPtr;

    for (i = 1; i <= 10; ++i)
    {
        ParamPtr = GetMemCalcParam (ParamBlock, i);

        if (ParamPtr == NULL)
            return UFNotEnoughParameters;

        if (ParamPtr->MemCalcParamType != UFStringArg)
            return UFMemCalcFail;

        if ((j += (ParamPtr->MemCalcParamLength-1))
            > UFMaxStringLength)
        {
           j = UFMaxStringLength;
           break;
        }

    }

    ParamBlock->UFMemCalcStringLength = j+1;;

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFSumCurrencyArray (UFParamBlock * ParamBlock)
{
    UFTInt16s i;
    UFParamListElement * ParamPtr = GetParam (ParamBlock, 1);

    if (ParamPtr == NULL)
        return UFNotEnoughParameters;

    ParamBlock->ReturnValue.ReturnNumber = 0;
    for (i = 0; i < ParamPtr->ParamLength; ++i)
        ParamBlock->ReturnValue.ReturnCurrency += ParamPtr->Parameter.ParamCurrencyArray[i];

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFArrayMax (UFParamBlock * ParamBlock)
{
    UFTInt16s i;
    UFParamListElement * ParamPtr = GetParam (ParamBlock, 1);

    if (ParamPtr == NULL)
        return UFNotEnoughParameters;

    ParamBlock->ReturnValue.ReturnNumber = ParamPtr->Parameter.ParamNumberArray[0];
    for (i = 0; i < ParamPtr->ParamLength; ++i)
        if (ParamBlock->ReturnValue.ReturnNumber < ParamPtr->Parameter.ParamNumberArray[i])
            ParamBlock->ReturnValue.ReturnNumber = ParamPtr->Parameter.ParamNumberArray[i];

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFAvgNumberArray (UFParamBlock * ParamBlock)
{
    UFTInt16s i;
    UFParamListElement * ParamPtr = GetParam (ParamBlock, 1);

    if (ParamPtr == NULL)
        return UFNotEnoughParameters;

    ParamBlock->ReturnValue.ReturnNumber = 0;
    for (i = 0; i < ParamPtr->ParamLength; ++i)
        ParamBlock->ReturnValue.ReturnNumber += ParamPtr->Parameter.ParamNumberArray[i];

    ParamBlock->ReturnValue.ReturnNumber /= ParamPtr->ParamLength;

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFStringLength (UFParamBlock * ParamBlock)
{
    ParamBlock->ReturnValue.ReturnNumber = (UFTNumber) lstrlen (GetParam(ParamBlock, 1)->Parameter.ParamString) * UFNumberScalingFactor;

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFTestErrors (UFParamBlock * ParamBlock)
{
    UFTInt32s ErrorNumber;

    if ((ErrorNumber = (UFTInt32s) (GetParam(ParamBlock, 1)->Parameter.ParamNumber)/UFNumberScalingFactor) > 6
        || (ErrorNumber < 0))
        return UFStringIndexTooBig;

    ParamBlock->ReturnValue.UFReturnUserError = ErrorNumber;

    return UFUserError;
}

ExternC UFError FAR _export PASCAL UFSumStringLengths (UFParamBlock * ParamBlock)
{
    UFTInt16s i;
    UFParamListElement * ParamPtr = GetParam(ParamBlock, 1);

    if (ParamPtr == NULL)
    {
       ParamBlock->ReturnValue.UFReturnUserError = 2;
       return UFUserError;
    }

    for (i = 0; i < ParamPtr->ParamLength; ++i)
    {
        if (lstrlen (ParamPtr->Parameter.ParamStringArray[i]) == 0)
        {
            ParamBlock->ReturnValue.UFReturnUserError = 1;
            return UFUserError;
        }

        ParamBlock->ReturnValue.ReturnNumber += lstrlen (ParamPtr->Parameter.ParamStringArray[i]);
    }

    ParamBlock->ReturnValue.ReturnNumber *= UFNumberScalingFactor;

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFArrayElementsInRange (UFParamBlock * ParamBlock)
{
    UFTInt16s Count, i;
    UFTNumber Max, Min;
    UFTNumber * NumArray;
    UFParamListElement * ParamPtr;

    Count = (UFTInt16s) (GetParam(ParamBlock, 3)->Parameter.ParamNumber/UFNumberScalingFactor);
    ParamPtr = GetParam(ParamBlock, 2);
    Min = ParamPtr->Parameter.ParamNumberRange.ParamNumberStart;
    Max = ParamPtr->Parameter.ParamNumberRange.ParamNumberEnd;
    ParamPtr = GetParam(ParamBlock, 1);
    for (i = 0; i < ParamPtr->ParamLength; ++i)
        if ((Min < ParamPtr->Parameter.ParamNumberArray[i]) &&
            (ParamPtr->Parameter.ParamNumberArray[i] < Max))
            --Count;

    ParamBlock->ReturnValue.ReturnBoolean = (Count <= 0);

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFStringSubstitutions (UFParamBlock * ParamBlock)
{
    UFTInt16s StringArrayLength, StringArrayIndex = 0;
    char * * StringArrayPtr, * ArrayElementPtr;
    char * TargetString = ParamBlock->ReturnValue.ReturnString;

    char * TemplateString = GetParam (ParamBlock, 1)->Parameter.ParamString;
    StringArrayPtr = GetParam (ParamBlock, 2)->Parameter.ParamStringArray;
    StringArrayLength = GetParam (ParamBlock, 2)->ParamLength;
    while (*TemplateString)
    {
        if ((*TemplateString == '*') &&(StringArrayIndex < StringArrayLength))
        {
            ArrayElementPtr = StringArrayPtr[StringArrayIndex++];
            while (*ArrayElementPtr)
                *(TargetString++) = *(ArrayElementPtr++);
            ++TemplateString;
        }
        else
            *(TargetString++) = *(TemplateString++);
    }

    *TargetString = '\0';

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFStringSubMemCalc (UFMemCalcBlock * ParamBlock)
{
    UFMemCalcParam * ParamPtr, * SubParamPtr;
    UFTInt16s x, y;

    if ((ParamPtr =  GetMemCalcParam (ParamBlock, 1)) == NULL)
    {
        ParamBlock->UFMemCalcStringLength = UFMaxStringLength;
        return (UFMemCalcFail);
    }

    if (ParamPtr->MemCalcParamType != UFStringArg)
    {
        ParamBlock->UFMemCalcStringLength = UFMaxStringLength;
        return (UFMemCalcFail);
    }

    if (ParamPtr->MemCalcLiteralFlag == UFLiteral)
        x = (ParamPtr->MemCalcParamLength);
    else
    {
        ParamBlock->UFMemCalcStringLength = UFMaxStringLength;
        return (UFMemCalcFail);
    }

    if ((ParamPtr =  GetMemCalcParam (ParamBlock, 2)) == NULL)
    {
        ParamBlock->UFMemCalcStringLength = UFMaxStringLength;
        return (UFMemCalcFail);
    }

    if (ParamPtr->MemCalcParamType != UFStringArray)
    {
        ParamBlock->UFMemCalcStringLength = UFMaxStringLength;
        return (UFMemCalcFail);
    }

    for (y=1; y <= ParamPtr->MemCalcParamLength; ++y)
    {
        SubParamPtr = GetMemCalcArrayElement (ParamPtr, y);
        if ((SubParamPtr != NULL) && (SubParamPtr->MemCalcParamType == UFStringArg))
            x += SubParamPtr->MemCalcParamLength - 2; // Minus the "*", and the '\0';
        else
        {
            ParamBlock->UFMemCalcStringLength = UFMaxStringLength;
            return (UFMemCalcFail);
        }
    }

    ParamBlock->UFMemCalcStringLength = x + 1; // The last '\0'

    return (UFNoError);
}

ExternC UFError FAR _export PASCAL TestStringArrayLength (UFMemCalcBlock * ParamBlock)
{
    UFMemCalcParam * ParamPtr, * SubParamPtr;
    UFTInt16s x, y;

    if ((ParamPtr =  GetMemCalcParam (ParamBlock, 1)) == NULL)
    {
        ParamBlock->UFMemCalcStringLength = UFMaxStringLength;
        return (UFMemCalcFail);
    }

    if (ParamPtr->MemCalcParamType != UFStringArray)
    {
        ParamBlock->UFMemCalcStringLength = UFMaxStringLength;
        return (UFMemCalcFail);
    }

    x = 0;
    for (y=1; y <= ParamPtr->MemCalcParamLength; ++y)
    {
        SubParamPtr = GetMemCalcArrayElement (ParamPtr, y);
        if ((SubParamPtr != NULL) && (SubParamPtr->MemCalcParamType == UFStringArg))
        {
            if (x < SubParamPtr->MemCalcParamLength)
                x = SubParamPtr->MemCalcParamLength;
        }
        else
        {
            ParamBlock->UFMemCalcStringLength = UFMaxStringLength;
            return (UFMemCalcFail);
        }
    }

    ParamBlock->UFMemCalcStringLength = x;

    return (UFNoError);
}

ExternC UFError FAR _export PASCAL UFMakeNumberArray (UFParamBlock * ParamBlock)
{
    UFTInt16s i;
    UFParamListElement * ParamPtr;

    for (i=1; i<=4; ++i)
    {
        if ((ParamPtr = GetParam (ParamBlock, i)) == NULL)
            ParamBlock->ReturnValue.ReturnNumberArray[i-1] = 0;
        else
            ParamBlock->ReturnValue.ReturnNumberArray[i-1] =
                ParamPtr->Parameter.ParamNumber;
    }

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFNumberArrayLength (UFMemCalcBlock * ParamBlock)
{
    ParamBlock->UFMemCalcArrayLength = 4;

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFMakeStringArray (UFParamBlock * ParamBlock)
{
    UFTInt16s i;
    UFParamListElement * ParamPtr;

    for (i=1; i<=4; ++i)
    {
        if ((ParamPtr = GetParam (ParamBlock, i)) == NULL)
            ParamBlock->ReturnValue.ReturnStringArray[i-1][0] = '\0';
        else
            lstrcpy (ParamBlock->ReturnValue.ReturnStringArray[i-1],
                     ParamPtr->Parameter.ParamString);
    }

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFStringArrayLength (UFMemCalcBlock * ParamBlock)
{
    UFTInt16s i;
    UFParamListElement * ParamPtr;
    UFMemCalcParam * ParseParamPtr;

    ParamBlock->UFMemCalcStringLength = 0;

    if (ParamBlock->MemCalcCommand == UFCalcStringLength)
    {   // This is at parse time - have to estimate the max string length
        // from the parse tree values.
        for (i=1; i<=4; ++i)
            if ((ParseParamPtr = GetMemCalcParam (ParamBlock, i)) == NULL)
                return UFMemCalcFail;
            else // find the longest string, use that length for the array elements sizes.
                if (ParamBlock->UFMemCalcStringLength <
                    ParseParamPtr->MemCalcParamLength + 1)
                        ParamBlock->UFMemCalcStringLength =
                            ParseParamPtr->MemCalcParamLength + 1; // Don't forget the NULL!

        return UFNoError;
    }

    for (i=1; i<=4; ++i)
    {
        if ((ParamPtr = GetMemCalcRunParam (ParamBlock, i)) == NULL)
            return UFMemCalcFail;
        else // find the longest string, use that length for the array elements sizes.
            if (ParamBlock->UFMemCalcStringLength <
                lstrlen(ParamPtr->Parameter.ParamString) + 1)
                    ParamBlock->UFMemCalcStringLength =
                        lstrlen(ParamPtr->Parameter.ParamString) + 1; // Don't forget the NULL!
    }

    ParamBlock->UFMemCalcArrayLength = 4;

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFMakeDateArray (UFParamBlock * ParamBlock)
{
    UFTInt16s i;
    UFParamListElement * ParamPtr;

    for (i=1; i<=2; ++i)
        if ((ParamPtr = GetParam (ParamBlock, i)) == NULL)
            ParamBlock->ReturnValue.ReturnDateArray[i-1] = 0;
        else
            ParamBlock->ReturnValue.ReturnDateArray[i-1] =
                ParamPtr->Parameter.ParamDate;

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFDateArrayLength (UFMemCalcBlock * ParamBlock)
{
    ParamBlock->UFMemCalcArrayLength = 2;
    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFMakeCurrencyArray (UFParamBlock * ParamBlock)
{
    UFTInt16s i;
    UFParamListElement * ParamPtr;

    for (i=1; i<=6; ++i)
    {
        if ((ParamPtr = GetParam (ParamBlock, i)) == NULL)
            ParamBlock->ReturnValue.ReturnCurrencyArray[i-1] = 0;
        else
            ParamBlock->ReturnValue.ReturnCurrencyArray[i-1] =
                ParamPtr->Parameter.ParamCurrency;
    }

    return UFNoError;
}

ExternC UFError FAR _export PASCAL UFCurrencyArrayLength (UFMemCalcBlock * ParamBlock)
{
    ParamBlock->UFMemCalcArrayLength = 6;

    return UFNoError;
}
