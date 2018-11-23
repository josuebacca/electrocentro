/*
** File:    UFMain.c
** Author:  Rick Cameron, Rex Benning
** Date:    7 Mar 93
**
** Purpose: The main file for a user-defined function DLL.
*/

#include <Windows.h>
#include "UFDll.h"
#include "UFFuncs.h"
#include "UFMain.h"
#include "UFUser.h"


HINSTANCE UFInstance = 0;

UFFunctionDefStringList FunctionDefStringList =
{
       {sizeof (UFFunctionDefStringList)},
       {FunctionDefStrings}
};

UFFunctionTemplateList FunctionTemplateList =
{
       {sizeof (UFFunctionTemplateList)},
       {FunctionTemplates}
};

UFFunctionExampleList FunctionExampleList =
{
       {sizeof (UFFunctionExampleList)},
       {FunctionExamples}
};
 
#if defined (WIN32)
BOOL WINAPI DllMain (HINSTANCE hInstance, ULONG reson_for_call, LPVOID lpReserved)
{
    UFInstance = hInstance;
    return TRUE;
}
#else

int FAR PASCAL LibMain (HINSTANCE hInstance,
                        WORD      wDataSeg,
                        WORD      cbHeapSize,
                        LPSTR     lpCmdLine)
{   
    UFInstance = hInstance;

    return TRUE;
}

int FAR PASCAL WEP (int nParam)
{
    return 1;
}
#endif  

ExternC UFError CR_EXPORT UFInitialize (void)
{
    return (UFNoError);
}

ExternC UFTInt16u CR_EXPORT UFGetVersion ()
{                           
    return CurVersionNumber;
}

ExternC UFError CR_EXPORT UFTerminate ()
{
    return (UFNoError);
}

ExternC UFFunctionTemplateList FAR * CR_EXPORT UFGetFunctionTemplates ()
{
    return (&FunctionTemplateList);
}

ExternC UFFunctionExampleList FAR * CR_EXPORT UFGetFunctionExamples ()
{
    return (&FunctionExampleList);
}

ExternC UFFunctionDefStringList FAR * CR_EXPORT UFGetFunctionDefStrings ()
{
    return (&FunctionDefStringList);
}

ExternC char FAR * CR_EXPORT UFErrorRecovery (UFParamBlock FAR *paramPtr,
                                      UFError errN)
{
    return ErrorTable [(UFTInt16u)paramPtr->ReturnValue.UFReturnUserError];
}

ExternC void CR_EXPORT UFStartJob (UFTInt32u jobId)
{
    // initialize any job-related data structures
    InitForJob (jobId);
}

ExternC void CR_EXPORT UFEndJob (UFTInt32u jobId)
{
    // clean up any job-related data structures
    TermForJob (jobId);
}
 

/* GetParam
** This is a routine internal to the DLL. It should be copied into every
** DLL as a standard access to the parameter block.
*/

UFParamListElement FAR *GetParam (UFParamBlock FAR *Param,
                                  UFTInt16s ParamNumber)
{
    UFParamListElement FAR *ParamPtr = Param->UFParamList;

    while (--ParamNumber && (ParamPtr != NULL))
        ParamPtr = ParamPtr->NextParam;

    return ParamPtr;
}
/* GetMemCalcParam
** This routine gets the parameters from the list passed to memory allocation
** (string length or array size) calculation routine.
*/

UFMemCalcParam FAR *GetMemCalcParam (UFMemCalcBlock FAR *Param,
                                     UFTInt16s ParamNumber)
{
    UFMemCalcParam FAR *ParamPtr = Param->UFMemCalcParamList;

    while (--ParamNumber && (ParamPtr != NULL))
        ParamPtr = ParamPtr->NextMemCalcParam;

    return ParamPtr;
}

/* GetMemCalcParam
** This routines gets array-element sub-arguments from the chain
** of arguments passed to the memory calculation routines.
*/

UFMemCalcParam FAR *GetMemCalcArrayElement (UFMemCalcParam FAR *Param,
                                            UFTInt16s ParamNumber)
{
    UFMemCalcParam FAR *ParamPtr = Param->MemCalcParameter.MemCalcSubParam;

    while (--ParamNumber && (ParamPtr != NULL))
        ParamPtr = ParamPtr->NextMemCalcParam;

    return ParamPtr;
}

/* GetMemCalcRunParam
** This is a routine internal to the DLL. It should be copied into every
** DLL as a standard access to the parameter block.
*/

UFParamListElement FAR *GetMemCalcRunParam (UFMemCalcBlock FAR *Param,
                                            UFTInt16s ParamNumber)
{
    UFParamListElement FAR *ParamPtr = Param->UFParamList;

    while (--ParamNumber && (ParamPtr != NULL))
        ParamPtr = ParamPtr->NextParam;

    return ParamPtr;
}
