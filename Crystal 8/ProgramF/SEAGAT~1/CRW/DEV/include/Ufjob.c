/*
** File:    UFJob.c
** Author:  Rick Cameron
** Date:    13 Mar 93
**
** Purpose: Implementation of functions to support job management.
*/

#include <Windows.h>
#include "UFDll.h"
#include "UFMain.h"
#include "UFUser.h"
#include "UFJob.h"

#include <StdLib.h>

static struct JobInfo FAR *Root = 0;

void InitForJob (UFTInt32u jobId)
{
    struct JobInfo FAR *jobInfo;

    jobInfo = FindJobInfo (jobId);

    if (jobInfo == 0)
    {
        jobInfo = malloc (sizeof (struct JobInfo));

        if (jobInfo != 0)
        {
            jobInfo->jobId = jobId;
            jobInfo->data  = 0;

            jobInfo->prev = 0;
            jobInfo->next = Root;

            if (Root != 0)
                Root->prev = jobInfo;

            Root = jobInfo;

            InitJob (jobInfo);
        }
    }
}

void TermForJob (UFTInt32u jobId)
{
    struct JobInfo FAR *jobInfo;

    jobInfo = FindJobInfo (jobId);

    if (jobInfo != 0)
    {
        TermJob (jobInfo);

        if (jobInfo->prev != 0)
            jobInfo->prev->next = jobInfo->next;
        else
            Root = jobInfo->next;

        if (jobInfo->next != 0)
            jobInfo->next->prev = jobInfo->prev;

        free (jobInfo);
    }
}

struct JobInfo FAR *FindJobInfo (UFTInt32u jobId)
{
    struct JobInfo FAR *cursor = Root;

    while (cursor != 0)
    {
        if (cursor->jobId == jobId)
            break;

        cursor = cursor->next;
    }

    return cursor;
}

