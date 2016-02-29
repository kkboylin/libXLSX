
/******************************************************************************
 *
 *	\file	libXLSX.h
 *	\brief	簡易 Excel xlsx 檔案讀取器
 *	\author	林祖池
 *
 ******************************************************************************/
#ifndef __XLSX_LIB_DLL_H__
#define	__XLSX_LIB_DLL_H__

void*          LoadXLSX        (const char* filename);
void           UnloadXLSX      (void* handle);

int            GetSheetCount   (void* book);
void*          GetSheetByIndex (void* book, unsigned int idx);
void*          GetSheetByName  (void* book, const wchar_t* name);

const wchar_t* GetSheetName    (void* sheet);
int            GetRecordCount  (void* sheet);
void*          GetRecordByIndex(void* sheet, unsigned int idx);

int            GetValueCount   (void* record);
const wchar_t* GetValue        (void* record, unsigned int idx);

#if defined(__cplusplus)

namespace xlsxLib
{
	class CRecord
	{
	private	:
		void*	_Handle;

	public	:
		CRecord(void* handle)
		{
			_Handle	=	handle;
		}

		CRecord(const CRecord& other)
		{
			_Handle	=	other._Handle;
		}

		const CRecord& operator=(const CRecord& other)
		{
			_Handle	=	other._Handle;
			return	*this;
		}

		unsigned int getCount() const
		{
			return	GetValueCount(_Handle);
		}

		const wchar_t* operator[](size_t idx) const
		{
			return	getValue( static_cast<unsigned int>(idx) );
		}

		const wchar_t* getValue(unsigned int idx) const
		{
			return	GetValue(_Handle, idx);
		}
	};

	class CSheet
	{
	private	:
		void*	_Handle;

	public	:
		CSheet(const CSheet& other)
		{
			_Handle	=	other._Handle;
		}

		CSheet(void* handle)
		{
			_Handle	=	handle;
		}

		const CSheet& operator=(const CSheet& other)
		{
			_Handle	=	other._Handle;
			return	*this;
		}

		unsigned int getCount() const
		{
			return	GetRecordCount(_Handle);
		}

		CRecord operator[](size_t idx) const
		{
			return	GetRecordByIndex(_Handle, static_cast<unsigned int>(idx) );
		}

		CRecord getRecord(unsigned int idx) const
		{
			return	GetRecordByIndex(_Handle, idx);
		}

		const wchar_t* getName() const
		{
			return	GetSheetName(_Handle);
		}
	};

	class CWorkBook
	{
	private:
		void*	_Handle;

		CWorkBook(const CWorkBook&)						{					}
		const CWorkBook& operator=(const CWorkBook&)	{	return	*this;	}
		CWorkBook(const char* filename)
		{
			_Handle	=	LoadXLSX(filename);
		}

	public	:
		static CWorkBook* Load(const char* filename)
		{
			CWorkBook*	obj	=	new CWorkBook(filename);
			if (obj)
			{
				if (obj->_Handle != NULL)
					return	obj;
				delete	obj;
			}
			return	NULL;
		}

		~CWorkBook()
		{
			if (_Handle)
				UnloadXLSX(_Handle);
		}

		unsigned int getCount() const
		{
			return	GetSheetCount(_Handle);
		}

		CSheet operator[](size_t idx) const
		{
			return	GetSheetByIndex(_Handle, static_cast<unsigned int>(idx));
		}

		CSheet operator[](const wchar_t* name) const
		{
			return	GetSheetByName(_Handle, name);
		}

		CSheet getSheet(unsigned int idx) const
		{
			return	GetSheetByIndex(_Handle, idx);
		}

		CSheet getSheet(const wchar_t* name) const
		{
			return	GetSheetByName(_Handle, name);
		}
	};
};

#endif	//	__cplusplus

#endif	//	__XLSX_LIB_DLL_H__
