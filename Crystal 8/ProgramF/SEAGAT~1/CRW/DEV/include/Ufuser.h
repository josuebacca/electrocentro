/*
** File:    UFUser.h
** Author:  Rick Cameron, Rex Benning
** Date:    7 Mar 93
**
** Purpose: Declarations of tables and functions which must be provided
**          for a user-defined function DLL.
*/

#if !defined (UFUSER_H)
#define UFUSER_H

#if defined (__cplusplus)
extern "C"
{
#endif

#define CurVersionNumber 0x0100

extern UFFunctionDefStrings      FunctionDefStrings [];
extern UFFunctionTemplates        FunctionTemplates  [];
extern UFFunctionExamples         FunctionExamples   [];

extern UFFunctionDefStringList   FunctionDefStringList;
extern UFFunctionTemplateList     FunctionTemplateList;
extern UFFunctionExampleList      FunctionExampleList;
extern char                       FAR *ErrorTable    [];

extern void InitForJob (UFTInt32u jobId);
extern void TermForJob (UFTInt32u jobId);

#if defined (__cplusplus)
}
#endif

#endif // !defined (UFUSER_H)
