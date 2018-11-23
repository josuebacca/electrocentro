/*
** File:    UFLSamp1.c
** Author:  Rick Cameron
** Date:    8 Mar 93
**
** Purpose: Example user-defined functions.
*/

#include <Windows.h>
#include "UFDll.h"
#include "UFMain.h"
#include "UFUser.h"
#include "UFJob.h"

#include <CType.h>

#include "WinTime.h"

#define PicturePlaceholder 'x'

#define WildOneChar  '?'
#define WildAllChars '*'

#define SoundexLength 4

ExternC UFError CR_EXPORT Now (UFParamBlock * ParamBlock);
ExternC UFError CR_EXPORT Picture (UFParamBlock * ParamBlock);
ExternC UFError CR_EXPORT LooksLike (UFParamBlock * ParamBlock);
ExternC UFError CR_EXPORT Soundex (UFParamBlock * ParamBlock);

ExternC UFError CR_EXPORT NowStringRetSize (UFMemCalcBlock * ParamBlock);
ExternC UFError CR_EXPORT SoundexStringRetSize (UFMemCalcBlock * ParamBlock);

UFFunctionDefStrings FunctionDefStrings [] =
{
    {"String Now", Now, NowStringRetSize},

    {"String Picture (String, String)", Picture},

    {"Boolean LooksLike (String, String)", LooksLike},

    {"String Soundex (String)", Soundex, SoundexStringRetSize},

    {NULL, NULL, NULL}
};

UFFunctionTemplates FunctionTemplates [] =
{
    {"Now!"},

    {"Picture (!, )"},

    {"LooksLike (!, )"},

    {"Soundex (!)"},

    {NULL}
};

UFFunctionExamples FunctionExamples [] =
{
    {"\tNow"},

    {"\tPicture (string, picture)"},

    {"\tLooksLike (string, mask)"},

    {"\tSoundex (string)"},

    {NULL}
};

char *ErrorTable [] =
{
    "no error",
};

void InitJob (struct JobInfo *jobInfo)
{
    TimeResetInternational ();
}

void TermJob (struct JobInfo *jobInfo)
{
}

ExternC UFError CR_EXPORT Now (UFParamBlock * ParamBlock)
{
    TimeGetCurTime (ParamBlock->ReturnValue.ReturnString, 0);

    return UFNoError;
}

ExternC UFError CR_EXPORT NowStringRetSize (UFMemCalcBlock * ParamBlock)
{
    // The longest string possible is "01:01:01 abcd"
    ParamBlock->UFMemCalcStringLength = 2 + 1 + 2 + 1 + 2 + 1 + 4 + 1;
    return UFNoError;
}

static void copyUsingPicture (char *dest,
                              const char *source,
                              const char *picture)
{
    while (*picture != '\0')
    {
        if (tolower (*picture) == PicturePlaceholder)
            if (*source != '\0')
                *dest++ = *source++;
            else
                ; // don't insert anything
        else
            *dest++ = *picture;

        picture++;
    }

    // copy the rest of the source
    lstrcpy (dest, source);
}

ExternC UFError CR_EXPORT Picture (UFParamBlock * ParamBlock)
{
    UFParamListElement *FirstParam,
                       *SecondParam;

    FirstParam  = GetParam (ParamBlock, 1);
    SecondParam = GetParam (ParamBlock, 2);

    if (FirstParam == NULL || SecondParam == NULL)
        return UFNotEnoughParameters;

    copyUsingPicture (ParamBlock->ReturnValue.ReturnString,
                      FirstParam->Parameter.ParamString,
                      SecondParam->Parameter.ParamString);

    return UFNoError;
}

static UFTBoolean looksLike (const char *string, const char *mask)
{
    while (*mask != '\0')
    {
        if (*mask == WildAllChars)
            return TRUE;
        else if (*mask == WildOneChar)
        {
            if (*string == '\0')
                return FALSE;
        }
        else if (*mask != *string)
            return FALSE;

        ++mask;
        ++string;
    }

    return (*string == '\0');
}

ExternC UFError CR_EXPORT LooksLike (UFParamBlock * ParamBlock)
{
    UFParamListElement *FirstParam,
                       *SecondParam;

    FirstParam  = GetParam (ParamBlock, 1);
    SecondParam = GetParam (ParamBlock, 2);

    if (FirstParam == NULL || SecondParam == NULL)
        return UFNotEnoughParameters;

    ParamBlock->ReturnValue.ReturnBoolean
        = looksLike (FirstParam->Parameter.ParamString,
                     SecondParam->Parameter.ParamString);

    return UFNoError;
}

// static Int16u soundexCode

typedef enum
{
    nonAlpha = -1,
    ignored  = '0',
    bfpv     = '1',
    cgjkqsxz = '2',
    dt       = '3',
    l        = '4',
    mn       = '5',
    r        = '6'
} SoundexCode;

static char AnsiLowerChar (char c)
{
    return (char)AnsiLower ((LPSTR)MAKELONG (c, 0));
}

static char AnsiUpperChar (char c)
{
    return (char)AnsiUpper ((LPSTR)MAKELONG (c, 0));
}

static SoundexCode calcCode (char c)
{
    if (!IsCharAlpha (c))
        return nonAlpha;

    c = AnsiLowerChar (c);

    switch (c)
    {
        case 'b': case 'f': case 'p': case 'v':
             return bfpv;

        case 'c': case 'g': case 'j': case 'k':
        case 'q': case 's': case 'x': case 'z':
        case 'ß': case 'ç': case 'š': // s with an inverted ^
             return cgjkqsxz;

        case 'd': case 't':
        case 'ð': case 'þ':
             return dt;

        case 'l':
             return l;

        case 'm': case 'n':
        case 'ñ':
             return mn;

        case 'r':
             return r;

        default:
             return ignored;
    }
}

static const char *skipBlanks (const char *string)
{
    while (*string == ' ')
        ++string;

    return string;
}

static void calcSoundex (char *buffer, const char *string)
{
    const char *inCursor  = string,
               *bufferEnd = buffer + SoundexLength;
    char       *outCursor = buffer;
    SoundexCode code;

    inCursor = skipBlanks (inCursor);

    code = calcCode (*inCursor);

    if (code == nonAlpha)
        *outCursor++ = '0';
    else
        *outCursor++ = AnsiUpperChar (*inCursor++);

    while (code != nonAlpha
           && *inCursor != '\0'
           && outCursor < bufferEnd)
    {
        SoundexCode lastCode = code;

        code = calcCode (*inCursor);

        if (code == nonAlpha)
            ;
        else
        {
            if (code == lastCode)
                ;
            else if (code == ignored)
                code = lastCode;
            else
                *outCursor++ = code;

            ++inCursor;
        }
    }

    while (outCursor < bufferEnd)
           *outCursor++ = '0';

    *outCursor = '\0';
}

ExternC UFError CR_EXPORT Soundex (UFParamBlock * ParamBlock)
{
    UFParamListElement * ParamPtr;

    ParamPtr = GetParam (ParamBlock, 1);
    if (ParamPtr == NULL)
        return UFNotEnoughParameters;

    calcSoundex (ParamBlock->ReturnValue.ReturnString,
                 ParamPtr->Parameter.ParamString);

    return UFNoError;
}

ExternC UFError CR_EXPORT SoundexStringRetSize (UFMemCalcBlock *ParamBlock)
{
    ParamBlock->UFMemCalcStringLength = SoundexLength + 1;
    return UFNoError;
}
