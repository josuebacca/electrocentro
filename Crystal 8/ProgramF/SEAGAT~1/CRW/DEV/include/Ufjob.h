/*
** File:    UFJob.h
** Author:  Rick Cameron
** Date:    13 Mar 93
**
** Purpose: Declarations of types and functions to support job management.
*/

#if !defined (UFJOB_H)
#define UFJOB_H

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

struct JobInfo
{
    UFTInt32u jobId;
    struct JobInfo FAR *prev;
    struct JobInfo FAR *next;

    // use this pointer to point to your data block
    void FAR *data;
};

extern void InitForJob (UFTInt32u jobId);
extern void TermForJob (UFTInt32u jobId);

extern struct JobInfo FAR *FindJobInfo (UFTInt32u jobId);

// you must implement these
extern void InitJob (struct JobInfo FAR *jobInfo);
extern void TermJob (struct JobInfo FAR *jobInfo);

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

#endif // !defined (UFJOB_H)
