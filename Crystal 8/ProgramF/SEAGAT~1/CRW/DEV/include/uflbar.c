/*
** File:    UFLBAR.c
** Authors: Kai Chan, Craig Chaplin, Randy Spiers
** Date:    28 Oct 94
**
** Purpose: Conversion functions to support Azalea bar code fonts
*/

#include <Windows.h>
#include <math.h>
#include <ctype.h>
#include <stdlib.h>
#include <stdio.h>
#include <UFDll.h>
#include "UFMain.h"
#include "UFUser.h"
#include "UFJob.h"

ExternC UFError CR_EXPORT NumberToPostnet (UFParamBlock * ParamBlock);
ExternC UFError CR_EXPORT StringToPostnet (UFParamBlock * ParamBlock);
ExternC UFError CR_EXPORT StringToCode39 (UFParamBlock * ParamBlock);
ExternC UFError CR_EXPORT NumberToCode39 (UFParamBlock * ParamBlock);

UFFunctionDefStrings  FunctionDefStrings [] =
{
	{"String NumberToPostnet (Number)", NumberToPostnet},
	{"String StringToPostnet (String)", StringToPostnet},
   {"String StringToCode39 (String)", StringToCode39},
   {"String NumberToCode39 (Number, Number)", NumberToCode39},
	{NULL, NULL, NULL}
};


UFFunctionTemplates FunctionTemplates [] =
{
       {"NumberToPostnet (!)"},
       {"StringToPostnet (!)"},
       {"StringToCode39 (!)"},
       {"NumberToCode39 (!, )"},
       {NULL}
};


UFFunctionExamples FunctionExamples [] =
{
       {"\tNumberToPostnet (number)"},
       {"\tStringToPostnet (string)"},
       {"\tStringToCode39 (string)"},
       {"\tNumberToCode39 (number, #places)"},
       {NULL}
};

char *ErrorTable [] =
{
    "no error"
};

void InitJob (struct JobInfo *jobInfo)
{
}

void TermJob (struct JobInfo *jobInfo)
{
}

/******************************************************************************/
/* Function:  static void GetCheckDigit (double Value, char *check)           */
/*                                                                            */
/* Author:    Kai Chan                                                        */
/*                                                                            */
/* Date:      Oct. 28, 1994                                                   */
/*                                                                            */
/* Purpose:   Takes a number, sums up all the digit in this number and        */
/*            returns a number which when added to the sum is a multiple      */
/*            of ten.                                                         */
/******************************************************************************/
static void GetCheckDigit (const double numval, char *check)
{
    double total, value, checkbit;
    int dig, sign;

    value = numval;
    total = 0.0;

    do {
       value /= 10.0;
       total += modf(value, &value);    /* add up all the remainders */
    } while (value);
    checkbit = (ceil(total) * 10.0)  - (total * 10.0);   /* get the checkbit  */

    if (checkbit > 9.001 || checkbit <  .999)    /* converting the checkbit to string */
       wsprintf(check,"%s","0\0");
    else
       wsprintf(check,"%s",ecvt(checkbit,1,&dig, &sign));

}


/******************************************************************************/
/* Function:  ExternC UFError NumberToPostnet (UFParamBlock * ParamBlock)     */
/*                                                                            */
/* Author:    Kai Chan                                                        */
/*                                                                            */
/* Date:      Oct. 28, 1994                                                   */
/*                                                                            */
/* Purpose:   Takes a number and converts it to an Azalea PostNet string.     */
/******************************************************************************/
ExternC UFError CR_EXPORT NumberToPostnet (UFParamBlock * ParamBlock)
{
   UFParamListElement * Param1;
   int dec, sign;
   double ndig, value;
   long rslt;
   char err[20], check[1];

   Param1 = GetParam (ParamBlock, 1);

   if (Param1 == NULL)
      return UFNotEnoughParameters;

   value = (UFTNumber)(Param1->Parameter.ParamNumber/100);

   /* get the number of digit */
   if (value == 0)
      ndig = 1;
   else
      ndig = 1 + log10(value);

      
   GetCheckDigit(value, check);
   wsprintf(ParamBlock->ReturnValue.ReturnString,"<%s%s>",ecvt(value, (int) ndig, &dec, &sign),
	    check);

   return UFNoError;
}


/******************************************************************************/
/* Function:  ExternC UFError StringToPostnet (UFParamBlock * ParamBlock)     */
/*                                                                            */
/* Author:    Kai Chan                                                        */
/*                                                                            */
/* Date:      Oct. 28, 1994                                                   */
/*                                                                            */
/* Purpose:   Takes a string and converts it to an Azalea PostNet string.     */
/******************************************************************************/
ExternC UFError CR_EXPORT StringToPostnet (UFParamBlock * ParamBlock)
{
   UFParamListElement * Param1;
   char *endptr, check[1];

   Param1 = GetParam (ParamBlock, 1);

   if (Param1 == NULL)
      return UFNotEnoughParameters;

   GetCheckDigit(strtod(Param1->Parameter.ParamString, &endptr), check);
   wsprintf(ParamBlock->ReturnValue.ReturnString,"<%s%s>",Param1->Parameter.ParamString,
	    check);

   return UFNoError;
}

/******************************************************************************/
/* Function:  static void Build39String (const char *source, char *dest)      */
/*                                                                            */
/* Author:    Craig Chaplin                                                   */
/*                                                                            */
/* Date:      Oct. 28, 1994                                                   */
/*                                                                            */
/* Purpose:   Converts source string into a Code 39 string which can be       */
/*            translated by Azalea's Code 39 bar code fonts.  Converted       */
/*            string is stored in dest.                                       */
/******************************************************************************/
static void Build39String (const char *source, char *dest)
{
   *(dest++) = '*'; // Azalea's Code 39 string prefix

   // loop until end of source string reached
   while (*source) {
      if (islower(*source)) {     // if source character is lower case
         *(dest++) = '+';            // convert into '+X' sequence used to
         *dest = toupper(*source);   // represent lowercase letters in bar code
      }
      else
         *dest = *source;         // else just copy the character
      ++source;
      ++dest;
   }

   *(dest++) = '*'; // Azalea's Code 39 string suffix
   *dest = '\0';    // null terminated dest string
}


/******************************************************************************/
/* Function:  ExternC UFError StringToCode39 (UFParamBlock * ParamBlock)      */
/*                                                                            */
/* Author:    Craig Chaplin                                                   */
/*                                                                            */
/* Date:      Oct. 28, 1994                                                   */
/*                                                                            */
/* Purpose:   User function that takes a string and converts it to a format   */
/*            compatible with Azalea's code 39 bar code fonts.                */
/******************************************************************************/
ExternC UFError CR_EXPORT StringToCode39 (UFParamBlock * ParamBlock)
{
   UFParamListElement *ParamPtr; // pointer to list of parameters passed

   // check to ensure parameter was passed
   if ((ParamPtr = GetParam (ParamBlock, 1)) == NULL)
      return UFNotEnoughParameters;

   *(ParamBlock->ReturnValue.ReturnString) = '\0'; // null terminate return string

   // if parameter is not a null string, convert it to a Code 39 string
   if (*(ParamPtr->Parameter.ParamString))
      Build39String (ParamPtr->Parameter.ParamString,
                   ParamBlock->ReturnValue.ReturnString);

   return UFNoError;
}

/******************************************************************************/
/* Function:  static void NumberToCode39String (char *dest, double cvtnum     */
/*                                                        , double decnum)    */
/* Author:    Randy Spiers                                                    */
/*                                                                            */
/* Date:      Oct. 28, 1994                                                   */
/*                                                                            */
/* Purpose:   Converts source number into a Code 39 string which can be       */
/*            translated by Azalea's Code 39 bar code fonts.  Converted       */
/*            string is stored in dest.                                       */
/******************************************************************************/
static void NumberToCode39String (char *dest, double cvtnum, double decnum)
{
  char string1[30], string2[30], string3[30];
  int i, j;
  int founddecimal = 0;

  if (decnum > 1100.0)                   //Limit the number of decimal places to 10
     decnum = 1100.0;

  itoa((int)decnum/100, string1, 10);    //Convert decimal number input to a string

  lstrcpy(string3, "%.");                //Create the format specifier string ("%.")
  lstrcat(string3, string1);             //                                   ("%.10")
  lstrcat(string3, "f");                 //                                   ("%.10f")
  sprintf(string2, string3, cvtnum/100);//Convert the input value to a string with the
                                         // specified number of decimal places

  for (i = 0; string2[i] != '\0' && i != 17; i++) //Set a flag if there is a 
    if (string2[i] == '.')                        // decimal point in the string
      founddecimal = 1;

//  MessageBox(NULL, string2, "String 2 before null added", MB_OK);

  if (lstrlen(string2) > 17 && !founddecimal)  //If the numeric string is longer than 17
  {
    string2[17] = '\0';                        // characters AND there is no decimal, 
                                               // set the 18th character to end-of-line.

    MessageBox(NULL, "The value you have input is too large.\n Output will be truncated."
                   , "Input Value Exceeds Limit", MB_ICONEXCLAMATION | MB_OK);
  }
  else                                         
    if (lstrlen(string2) > 17 && founddecimal) //If the numeric string is longer than 17
      string2[18] = '\0';                      // characters AND there is a decimal, 
                                               // set the 19th character to end-of-line.

//  MessageBox(NULL, string2, "String 2 after null added", MB_OK);

  for (i = 0; string2[i] != '.' && string2[i] != '\0'; i++) //Check if there is a decimal
    ;                                                       // point in the string.

  if (string2[i] == '.')                               //if the string contains a decimal
  {                                                    // point, advance until a maximum of
    for (j = 0; string2[i] != '\0' && j < 10; i++, j++)// ten more characters are read or the
      ;                                                // end-of-line character is reached.
    string2[++i] = '\0';                               // Put in an end-of-line character
  }                                                     
    
  lstrcpy(string1, "*");                 //Place asterisk in front of string
  lstrcat(string1, string2);             //Append string
  lstrcat(string1, "*");                 //Place asterisk at end of string

	lstrcpy(dest, string1);                //Copy the string to the destination
}

/******************************************************************************/
/* Function:  ExternC UFError NumberToCode39 (UFParamBlock * ParamBlock)      */
/*                                                                            */
/* Author:    Randy Spiers                                                    */
/*                                                                            */
/* Date:      Oct. 28, 1994                                                   */
/*                                                                            */
/* Purpose:   User function that takes a number and converts it to a format   */
/*            compatible with Azalea's code 39 bar code fonts.                */
/******************************************************************************/
ExternC UFError CR_EXPORT NumberToCode39 (UFParamBlock * ParamBlock)
{
	UFParamListElement *FirstParam, *SecondParam;
	FirstParam = GetParam (ParamBlock, 1);
	SecondParam = GetParam (ParamBlock, 2);

	if (FirstParam == NULL || SecondParam == NULL)
	   return UFNotEnoughParameters;

	NumberToCode39String (ParamBlock->ReturnValue.ReturnString,
                         FirstParam->Parameter.ParamNumber,
                         SecondParam->Parameter.ParamNumber);

	return UFNoError;
}

