// stdafx.h : ��׼ϵͳ�����ļ��İ����ļ���
// ���Ǿ���ʹ�õ��������ĵ�
// �ض�����Ŀ�İ����ļ�

#pragma once

#ifndef STRICT
#define STRICT
#endif

#include "targetver.h"

#define _ATL_APARTMENT_THREADED

#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// ĳЩ CString ���캯��������ʽ��


#define ATL_NO_ASSERT_ON_DESTROY_NONEXISTENT_WINDOW

#include "resource.h"
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>
# include <Windows.h>
#import "C:\Program Files (x86)\Common Files\DESIGNER\MSADDNDR.DLL" raw_interfaces_only, raw_native_types, no_namespace, named_guids, auto_search

#import "C:\Program Files (x86)\Microsoft Office\Office12\MSWORD.OLB" \
	raw_interfaces_only, raw_native_types, named_guids\
	rename("FindText", "FindTextEx") \
	rename("FontNames", "FontNamesEx") \
	rename("ExitWindows", "ExitWindowsEx")

#import "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE12\MSO.DLL" \
	rename("RGB", "RGBEx") \
	rename("DocumentProperties", "DocumentPropertiesEx")

//#import "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
