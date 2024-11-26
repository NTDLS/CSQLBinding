///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Copyright © NetworkDLS 2002, All rights reserved
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF 
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO 
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A 
// PARTICULAR PURPOSE.
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#ifndef _CRECORDSET_H
#define _CRECORDSET_H
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include <SQL.H>
#include <SQLExt.H>

typedef int(*ErrorHandler)(void* pCaller, const char* sSource, const char* sErrorMsg, const int iErrorNumber);

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

class CRecordSet {

private:
	ErrorHandler pErrorHandler;

	bool bcTrimCharData;
	bool bcReplaceSingleQuotes;
	bool bcThrowErrors;

public:
	void* pPublicData;

	void SetErrorHandler(ErrorHandler pHandler);

	bool Close();
	bool Fetch();
	bool Fetch(int* iErrorCode);
	bool Reset();

	bool GetData(const unsigned short iColumnIndex, short iType, void* pvBuf, SQLLEN iBufSz, SQLLEN* piIndicator);
	bool GetData(const unsigned short iColumnIndex, short iType, void* pvBuf, SQLLEN iBufSz, SQLLEN* piIndicator, short* piResult);
	bool GetColumnInfo(const int iColumnIndex, char* sOutName, int iSzOfOutName,
		int* iOutColNameLen, int* ioutDataType, SQLULEN* iOutColSize, int* iNumOfDeciPlaces, int* iColNullable);
	bool GetColumnBinaryValue(const int iColumnIndex, char* sBuf, const SQLLEN iBufSz, SQLLEN* iOutLen);
	bool GetColumnStringValue(const int iColumnIndex, char* sBuf, const SQLLEN iBufSz, SQLLEN* iOutLen);
	bool GetColumnLongValue(const int iColumnIndex, long* plOutVal);
	long GetColumnLongValue(const int iColumnIndex);
	bool GetColumnDoubleValue(const int iColumnIndex, double* pdOutVal);
	double GetColumnDoubleValue(const int iColumnIndex);
	bool GetColumnFloatValue(const int iColumnIndex, float* pfOutVal);
	float GetColumnFloatValue(const int iColumnIndex);
	bool GetErrorMessage(int* iOutErr, char* sOutError, const int iErrBufSz);
	bool GetSQLState(char* sSQLState);
	bool ThrowErrorIfSet();
	bool ThrowError();

	SQLLEN RTrim(char* sData, SQLLEN iDataSz);
	void ReplaceSingleQuotes(char* sData, SQLLEN iDataSz);

	bool ThrowErrors();
	bool ThrowErrors(bool bThrowErrors);

	bool TrimCharData(bool bTrimCharData);
	bool TrimCharData();

	bool ReplaceSingleQuotes(bool bReplaceSingleQuotes);
	bool ReplaceSingleQuotes();

	HSTMT hSTMT;
	SQLLEN RowCount;
	SQLLEN ColumnCount;

	CRecordSet();
	~CRecordSet();
};

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#endif
