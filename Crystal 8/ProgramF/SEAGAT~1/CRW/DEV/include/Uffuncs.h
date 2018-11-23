/*
** File:    UFFuncs.h
** Author:  Rex Benning
** Date:
**
** Purpose: Declaration of the entry points a user-defined function DLL
**          must provide.
*/


#if !defined (UFFUNCS_H)
#define UFFUNCS_H

#include "UFDll.h"

#if defined (__cplusplus)
extern "C"
{
#endif

// 'UF5Initialize' is the initialization function for new UFLs.  It lets the
// UFL know which version of CRW or CRPE it is working with.  'UFInitialize'
// will be used if the UFL is missing the newer function.
UFError CR_EXPORT UFInitialize (void);
UFError CR_EXPORT UF5Initialize (CRVersion crVersion);

UFError CR_EXPORT UFTerminate (void);

UFFunctionDefStringList FAR *CR_EXPORT UFGetFunctionPrototypes (void);

UFFunctionTemplateList FAR *CR_EXPORT UFGetFunctionTemplates (void);

UFFunctionExampleList FAR *CR_EXPORT UFGetFunctionExamples (void);

UFError CR_EXPORT UFCallFunction (UFTInt16s i,
                                  UFParamBlock FAR *paramPtr);

char FAR *CR_EXPORT UFErrorRecovery (UFParamBlock FAR *paramPtr,
                                     UFError errN);

UFTInt16u CR_EXPORT UFGetVersion (void);

// 'UF5GetFunctionDefStrings' is the preferred function for new UFLs.  It
// tells CRW or CRPE more about the functions defined by the UFL.
// 'UFGetFunctionDefStrings' will be used if the UFL is missing the newer
// function.
UFFunctionDefStringList FAR *CR_EXPORT UFGetFunctionDefStrings (void);
UF5FunctionDefStringList FAR *CR_EXPORT UF5GetFunctionDefStrings (void);

void CR_EXPORT UFStartJob (UFTInt32u jobId);
void CR_EXPORT UFEndJob (UFTInt32u jobId);     

HGLOBAL CR_EXPORT UF5SaveState (UFTInt32u jobId);
UFError CR_EXPORT UF5RestoreState (UFTInt32u jobId,
                                   HGLOBAL savedState);

#if defined (__cplusplus)
}
#endif

#endif // !defined (UFFUNCS_H)
