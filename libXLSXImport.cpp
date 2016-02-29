
#include <Windows.h>

#include "libXLSX.h"

namespace
{
	static void* GetXLSXLibrary()
	{
		static void*	handle = NULL;
		if (handle != NULL)
			handle = LoadLibrary("libXLSXDll.dll");
		return	handle;
	}

	typedef void*(*FLoadXLSX)(const char*);
	typedef void(*FUnloadXLSX)(void*);
	typedef int(*FGetSheetCount)(void*);
	typedef void*(*FGetSheetByIndex)(void* book, unsigned int idx);
	typedef void*(*FGetSheetByName)(void* book, const wchar_t* name);
	typedef const wchar_t*(*FGetSheetName)(void* sheet);
	typedef int(*FGetRecordCount)(void* sheet);
	typedef void*(*FGetRecordByIndex)(void* sheet, unsigned int idx);
	typedef int(*FGetValueCount)(void* record);
	typedef const wchar_t*(*FGetValue)(void* record, unsigned int idx);

	class CHandle
	{
	private		:
		HMODULE				_Handle;
		FLoadXLSX			_LoadXLSX;
		FUnloadXLSX			_UnloadXLSX;
		FGetSheetCount		_GetSheetCount;
		FGetSheetByIndex	_GetSheetByIndex;
		FGetSheetByName		_GetSheetByName;
		FGetSheetName		_GetSheetName;
		FGetRecordCount		_GetRecordCount;
		FGetRecordByIndex	_GetRecordByIndex;
		FGetValueCount		_GetValueCount;
		FGetValue			_GetValue;

	protected	:
		void Init()
		{
			if (_Handle != NULL)
				return;
			_Handle	=	LoadLibrary("libXLSX.dll");
			if (_Handle == NULL)
				return;
			_LoadXLSX			=	(FLoadXLSX)GetProcAddress(_Handle, "LoadXLSX_");
			_UnloadXLSX			=	(FUnloadXLSX)GetProcAddress(_Handle, "UnloadXLSX_");
			_GetSheetCount		=	(FGetSheetCount)GetProcAddress(_Handle, "GetSheetCount_");
			_GetSheetByIndex	=	(FGetSheetByIndex)GetProcAddress(_Handle, "GetSheetByIndex_");
			_GetSheetByName		=	(FGetSheetByName)GetProcAddress(_Handle, "GetSheetByName_");
			_GetSheetName		=	(FGetSheetName)GetProcAddress(_Handle, "GetSheetName_");
			_GetRecordCount		=	(FGetRecordCount)GetProcAddress(_Handle, "GetRecordCount_");
			_GetRecordByIndex	=	(FGetRecordByIndex)GetProcAddress(_Handle, "GetRecordByIndex_");
			_GetValueCount		=	(FGetValueCount)GetProcAddress(_Handle, "GetValueCount_");
			_GetValue			=	(FGetValue)GetProcAddress(_Handle, "GetValue_");
			if( (_LoadXLSX == NULL) ||
				(_UnloadXLSX == NULL) ||
				(_GetSheetCount == NULL) ||
				(_GetSheetByIndex == NULL) ||
				(_GetSheetByName == NULL) ||
				(_GetRecordCount == NULL) ||
				(_GetRecordByIndex == NULL) ||
				(_GetSheetName == NULL) ||
				(_GetValueCount == NULL) ||
				(_GetValue == NULL) )
			{
				FreeLibrary(_Handle);
				_Handle	=	NULL;
			}
		}

	public	:
		CHandle()
		{
			_Handle				=	NULL;
			_LoadXLSX			=	NULL;
			_UnloadXLSX			=	NULL;
			_GetSheetCount		=	NULL;
			_GetSheetByIndex	=	NULL;
			_GetSheetByName		=	NULL;
			_GetSheetName		=	NULL;
			_GetRecordCount		=	NULL;
			_GetRecordByIndex	=	NULL;
			_GetValueCount		=	NULL;
			_GetValue			=	NULL;
		}

		~CHandle()
		{
			if (_Handle)
				FreeLibrary(_Handle);
		}

		FLoadXLSX         getLoadXLSX()			{	Init();	return	_LoadXLSX;			}
		FUnloadXLSX	      getUnloadXLSX()		{	Init();	return	_UnloadXLSX;		}
		FGetSheetCount	  getGetSheetCount()	{	Init();	return	_GetSheetCount;		}
		FGetSheetByIndex  getGetSheetByIndex()	{	Init();	return	_GetSheetByIndex;	}
		FGetSheetByName   getGetSheetByName()	{	Init();	return	_GetSheetByName;	}
		FGetSheetName     getGetSheetName()		{	Init();	return	_GetSheetName;		}
		FGetRecordCount   getGetRecordCount()	{	Init();	return	_GetRecordCount;	}
		FGetRecordByIndex getGetRecordByIndex()	{	Init();	return	_GetRecordByIndex;	}
		FGetValueCount    getGetValueCount()	{	Init();	return	_GetValueCount;		}
		FGetValue         getGetValue()			{	Init();	return	_GetValue;			}
	};

	static CHandle* getHandle()
	{
		static CHandle*	handle = NULL;
		if (handle == NULL)
			handle = new CHandle();
		return	handle;
	}
};


void* LoadXLSX(const char* filename)
{
	if (getHandle()->getLoadXLSX() )
		return	getHandle()->getLoadXLSX()(filename);
	return	NULL;
}

void UnloadXLSX(void* handle)
{
	if (getHandle()->getUnloadXLSX())
		getHandle()->getUnloadXLSX()(handle);
}

int GetSheetCount(void* book)
{
	if (getHandle()->getGetSheetCount())
		return	getHandle()->getGetSheetCount()(book);
	return	0;
}

void* GetSheetByIndex(void* book, unsigned int idx)
{
	if (getHandle()->getGetSheetByIndex())
		return	getHandle()->getGetSheetByIndex()(book, idx);
	return	NULL;
}

void* GetSheetByName(void* book, const wchar_t* name)
{
	if (getHandle()->getGetSheetByName())
		return	getHandle()->getGetSheetByName()(book, name);
	return	NULL;
}

const wchar_t* GetSheetName(void* sheet)
{
	if (getHandle()->getGetSheetName())
		return	getHandle()->getGetSheetName()(sheet);
	return	NULL;
}

int GetRecordCount(void* sheet)
{
	if (getHandle()->getGetRecordCount())
		return	getHandle()->getGetRecordCount()(sheet);
	return	0;
}

void* GetRecordByIndex(void* sheet, unsigned int idx)
{
	if (getHandle()->getGetRecordByIndex())
		return	getHandle()->getGetRecordByIndex()(sheet, idx);
	return	NULL;
}

int GetValueCount(void* record)
{
	if (getHandle()->getGetValueCount())
		return	getHandle()->getGetValueCount()(record);
	return	0;
}

const wchar_t* GetValue(void* record, unsigned int idx)
{
	if (getHandle()->getGetValue())
		return	getHandle()->getGetValue()(record, idx);
	return	NULL;
}
