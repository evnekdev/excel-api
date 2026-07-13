@echo off
setlocal

set TYPE=DEBUG
set OP=
set PLATFORM_OPTION=

if "%1" == "" goto usage

if not "%VCINSTALLDIR%"=="" goto nextarg
@echo ERROR: VCINSTALLDIR variable is not set. 
@goto end

:nextarg
	shift

if "%0" == ""   goto build
if /i "%0" == "ship" (set TYPE=RELEASE)
if /i "%0" == "release" (set TYPE=RELEASE)
if /i "%0" == "clean" (set OP=clean)
if /i "%0" == "minclean" (set OP=minclean)

if /i "%0" == "x86" (set PLATFORM_OPTION=x86 & set PLATFORM=x86)
if /i "%0" == "amd64" (set PLATFORM_OPTION=amd64 & set PLATFORM=x64)
if /i "%0" == "x64" (set PLATFORM_OPTION=x64 & set PLATFORM=x64)
if /i "%0" == "x86_amd64"  (set PLATFORM_OPTION=x86_amd64 & set PLATFORM=x64)
if /i "%0" == "x86_x64"  (set PLATFORM_OPTION=x86_amd64 & set PLATFORM=x64)

goto nextarg

:usage
	echo.
	echo usage:
	echo.
	echo make TYPE OP PLATFORM
	echo.
	echo TYPE = DEBUG^|RELEASE^|SHIP
	echo.
	echo if no type is specified, DEBUG is assumed.
	echo SHIP and RELEASE refer to same type
	echo.
	echo OP = MINCLEAN^|CLEAN^|BUILD
	echo.
	echo if no OP is specified BUILD is assumed.
	echo.
	echo MINCLEAN deletes all build files expect for the 
	echo buildlog and generated LIB or XLL file(s)
	echo.
	echo CLEAN deletes all build files and the build folder
	echo.
	echo either OP or TYPE must be specified
	echo.
	echo PLATFORM = X86^|X64^|AMD64^|X86_X64^|X86_AMD64
	echo.
	echo if no platform is specified, X86 is assumed.
	echo X64 and AMD64 refer to same platform
	echo X86_X64 and X86_AMD64 refer to same cross-compile platform
	echo.
	goto eof

:build
	@call "%VCINSTALLDIR%\vcvarsall.bat" %PLATFORM_OPTION%
	@for /F %%d in (targets) do @call build.bat %%d %OP%
:eof
