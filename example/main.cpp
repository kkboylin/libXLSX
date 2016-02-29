
#include <Windows.h>
#include <stdlib.h>
#include <stdio.h>

#include <string>

#include "libXLSX.h"

static bool isNullRecord(const xlsxLib::CRecord& record)
{
	for (int k = 0; k < record.getCount(); ++k)
	{
		if ((record.getValue(k) != NULL) &&
			(record.getValue(k)[0] != 0))
			return	false;
	}
	return	true;
}

int main(int argc, const char** argv)
{
	xlsxLib::CWorkBook*	book	=	xlsxLib::CWorkBook::Load("test.xlsx");
	wchar_t				wbuffer[1024];
	swprintf(wbuffer, L"Sheet Count : %d\n", book->getCount());
	OutputDebugStringW(wbuffer);
	for (int i = 0; i < book->getCount(); ++i)
	{
		xlsxLib::CSheet	sheet	=	book->getSheet(i);
		swprintf(wbuffer, L"\tSheet %d     : %s\n", i, sheet.getName());
		OutputDebugStringW(wbuffer);

		swprintf(wbuffer, L"\tRecord Count : %d\n", sheet.getCount() );
		OutputDebugStringW(wbuffer);

		for (int x = 0; x < sheet.getCount(); ++x)
		{
			xlsxLib::CRecord	record	=	sheet.getRecord(x);
			if (isNullRecord(record) != true)
			{
				std::wstring	value(L"\t\t");
				for (int k = 0; k < record.getCount(); ++k)
				{
					if (k > 0)
						value += L",";
					value	+=	record.getValue(k);
				}
				value	+=	L"\n";
				OutputDebugStringW(value.c_str());
			}
		}
	}
	delete	book;
	return	0;
}
