# Microsoft Visual C++ Generated NMAKE File, Format Version 2.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

!IF "$(CFG)" == ""
CFG=Win32 Debug
!MESSAGE No configuration specified.  Defaulting to Win32 Debug.
!ENDIF 

!IF "$(CFG)" != "Win32 Release" && "$(CFG)" != "Win32 Debug"
!MESSAGE Invalid configuration "$(CFG)" specified.
!MESSAGE You can specify a configuration when running NMAKE on this makefile
!MESSAGE by defining the macro CFG on the command line.  For example:
!MESSAGE 
!MESSAGE NMAKE /f "u2lsamp1.mak" CFG="Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 
!ERROR An invalid configuration is specified.
!ENDIF 

################################################################################
# Begin Project
# PROP Target_Last_Scanned "Win32 Debug"
MTL=MkTypLib.exe
CPP=cl.exe
RSC=rc.exe

!IF  "$(CFG)" == "Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "WinRel"
# PROP BASE Intermediate_Dir "WinRel"
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "WinRel"
# PROP Intermediate_Dir "WinRel"
OUTDIR=.\WinRel
INTDIR=.\WinRel

ALL : $(OUTDIR)/u2lsamp1.dll $(OUTDIR)/u2lsamp1.bsc

$(OUTDIR) : 
    if not exist $(OUTDIR)/nul mkdir $(OUTDIR)

# ADD BASE MTL /nologo /D "NDEBUG" /win32
# ADD MTL /nologo /D "NDEBUG" /win32
MTL_PROJ=/nologo /D "NDEBUG" /win32 
# ADD BASE CPP /nologo /MT /W3 /GX /YX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /FR /c
# ADD CPP /nologo /Zp8 /MT /W3 /GX /YX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /FR /c
CPP_PROJ=/nologo /Zp8 /MT /W3 /GX /YX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS"\
 /FR$(INTDIR)/ /Fp$(OUTDIR)/"u2lsamp1.pch" /Fo$(INTDIR)/ /c 
CPP_OBJS=.\WinRel/
# ADD BASE RSC /l 0x1009 /d "NDEBUG"
# ADD RSC /l 0x1009 /d "NDEBUG"
RSC_PROJ=/l 0x1009 /fo$(INTDIR)/"UFLSAMP1.res" /d "NDEBUG" 
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o$(OUTDIR)/"u2lsamp1.bsc" 
BSC32_SBRS= \
	$(INTDIR)/UFJOB.SBR \
	$(INTDIR)/UFMAIN.SBR \
	$(INTDIR)/UFLSAMP1.SBR \
	$(INTDIR)/WINTIME.SBR

$(OUTDIR)/u2lsamp1.bsc : $(OUTDIR)  $(BSC32_SBRS)
    $(BSC32) @<<
  $(BSC32_FLAGS) $(BSC32_SBRS)
<<

LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /NOLOGO /SUBSYSTEM:windows /DLL /MACHINE:I386
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /NOLOGO /SUBSYSTEM:windows /DLL /MACHINE:I386
LINK32_FLAGS=kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib\
 advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib\
 odbccp32.lib /NOLOGO /SUBSYSTEM:windows /DLL /INCREMENTAL:no\
 /PDB:$(OUTDIR)/"u2lsamp1.pdb" /MACHINE:I386 /DEF:".\U2LSAMP1.DEF"\
 /OUT:$(OUTDIR)/"u2lsamp1.dll" /IMPLIB:$(OUTDIR)/"u2lsamp1.lib" 
DEF_FILE=.\U2LSAMP1.DEF
LINK32_OBJS= \
	$(INTDIR)/UFJOB.OBJ \
	$(INTDIR)/UFMAIN.OBJ \
	$(INTDIR)/UFLSAMP1.OBJ \
	$(INTDIR)/WINTIME.OBJ \
	$(INTDIR)/UFLSAMP1.res

$(OUTDIR)/u2lsamp1.dll : $(OUTDIR)  $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

!ELSEIF  "$(CFG)" == "Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "WinDebug"
# PROP BASE Intermediate_Dir "WinDebug"
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "WinDebug"
# PROP Intermediate_Dir "WinDebug"
OUTDIR=.\WinDebug
INTDIR=.\WinDebug

ALL : $(OUTDIR)/u2lsamp1.dll $(OUTDIR)/u2lsamp1.bsc

$(OUTDIR) : 
    if not exist $(OUTDIR)/nul mkdir $(OUTDIR)

# ADD BASE MTL /nologo /D "_DEBUG" /win32
# ADD MTL /nologo /D "_DEBUG" /win32
MTL_PROJ=/nologo /D "_DEBUG" /win32 
# ADD BASE CPP /nologo /MT /W3 /GX /Zi /YX /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /FR /c
# ADD CPP /nologo /Zp1 /MT /W3 /GX /Zi /YX /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /FR /c
CPP_PROJ=/nologo /Zp1 /MT /W3 /GX /Zi /YX /Od /D "WIN32" /D "_DEBUG" /D\
 "_WINDOWS" /FR$(INTDIR)/ /Fp$(OUTDIR)/"u2lsamp1.pch" /Fo$(INTDIR)/\
 /Fd$(OUTDIR)/"u2lsamp1.pdb" /c 
CPP_OBJS=.\WinDebug/
# ADD BASE RSC /l 0x1009 /d "_DEBUG"
# ADD RSC /l 0x1009 /d "_DEBUG"
RSC_PROJ=/l 0x1009 /fo$(INTDIR)/"UFLSAMP1.res" /d "_DEBUG" 
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
BSC32_FLAGS=/nologo /o$(OUTDIR)/"u2lsamp1.bsc" 
BSC32_SBRS= \
	$(INTDIR)/UFJOB.SBR \
	$(INTDIR)/UFMAIN.SBR \
	$(INTDIR)/UFLSAMP1.SBR \
	$(INTDIR)/WINTIME.SBR

$(OUTDIR)/u2lsamp1.bsc : $(OUTDIR)  $(BSC32_SBRS)
    $(BSC32) @<<
  $(BSC32_FLAGS) $(BSC32_SBRS)
<<

LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /NOLOGO /SUBSYSTEM:windows /DLL /DEBUG /MACHINE:I386
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /NOLOGO /SUBSYSTEM:windows /DLL /DEBUG /MACHINE:I386
LINK32_FLAGS=kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib\
 advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib\
 odbccp32.lib /NOLOGO /SUBSYSTEM:windows /DLL /INCREMENTAL:yes\
 /PDB:$(OUTDIR)/"u2lsamp1.pdb" /DEBUG /MACHINE:I386 /DEF:".\U2LSAMP1.DEF"\
 /OUT:$(OUTDIR)/"u2lsamp1.dll" /IMPLIB:$(OUTDIR)/"u2lsamp1.lib" 
DEF_FILE=.\U2LSAMP1.DEF
LINK32_OBJS= \
	$(INTDIR)/UFJOB.OBJ \
	$(INTDIR)/UFMAIN.OBJ \
	$(INTDIR)/UFLSAMP1.OBJ \
	$(INTDIR)/WINTIME.OBJ \
	$(INTDIR)/UFLSAMP1.res

$(OUTDIR)/u2lsamp1.dll : $(OUTDIR)  $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

!ENDIF 

.c{$(CPP_OBJS)}.obj:
   $(CPP) $(CPP_PROJ) $<  

.cpp{$(CPP_OBJS)}.obj:
   $(CPP) $(CPP_PROJ) $<  

.cxx{$(CPP_OBJS)}.obj:
   $(CPP) $(CPP_PROJ) $<  

################################################################################
# Begin Group "Source Files"

################################################################################
# Begin Source File

SOURCE=.\UFJOB.C
DEP_UFJOB=\
	.\UFDLL.H\
	.\UFMAIN.H\
	.\UFUSER.H\
	.\UFJOB.H

$(INTDIR)/UFJOB.OBJ :  $(SOURCE)  $(DEP_UFJOB) $(INTDIR)

# End Source File
################################################################################
# Begin Source File

SOURCE=.\UFMAIN.C
DEP_UFMAI=\
	.\UFDLL.H\
	.\UFFUNCS.H\
	.\UFMAIN.H\
	.\UFUSER.H

$(INTDIR)/UFMAIN.OBJ :  $(SOURCE)  $(DEP_UFMAI) $(INTDIR)

# End Source File
################################################################################
# Begin Source File

SOURCE=.\UFLSAMP1.C
DEP_UFLSA=\
	.\UFDLL.H\
	.\UFMAIN.H\
	.\UFUSER.H\
	.\UFJOB.H\
	.\WINTIME.H

$(INTDIR)/UFLSAMP1.OBJ :  $(SOURCE)  $(DEP_UFLSA) $(INTDIR)

# End Source File
################################################################################
# Begin Source File

SOURCE=.\WINTIME.C
DEP_WINTI=\
	.\WINTIME.H

$(INTDIR)/WINTIME.OBJ :  $(SOURCE)  $(DEP_WINTI) $(INTDIR)

# End Source File
################################################################################
# Begin Source File

SOURCE=.\UFLSAMP1.RC
DEP_UFLSAM=\
	.\WINTIME.H

$(INTDIR)/UFLSAMP1.res :  $(SOURCE)  $(DEP_UFLSAM) $(INTDIR)
   $(RSC) $(RSC_PROJ)  $(SOURCE) 

# End Source File
################################################################################
# Begin Source File

SOURCE=.\U2LSAMP1.DEF
# End Source File
# End Group
# End Project
################################################################################
