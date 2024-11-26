///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Copyright © NetworkDLS 2002, All rights reserved
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF 
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO 
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A 
// PARTICULAR PURPOSE.
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#ifndef _CBoundRecordSet_H
#define _CBoundRecordSet_H
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include <SQL.H>
#include <SQLExt.H>
#include "CSQLValue.H"

typedef int(*ErrorHandler)(void* pCaller, const char* sSource, const char* sErrorMsg, const int iErrorNumber);

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

class CBoundRecordSet {

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

	bool GetColumnInfo(const int iColumnIndex, char* sOutName, int iSzOfOutName, int* iOutColNameLen,
		int* ioutDataType, SQLULEN* iOutColSize, int* iNumOfDeciPlaces, int* iColNullable);

	CSQLValue* Values(const char* sColumnName);
	CSQLValue* Values(int iColumnIndex);

	bool GetErrorMessage(int* iOutErr, char* sOutError, const int iErrBufSz);
	bool GetSQLState(char* sSQLState);
	bool ThrowErrorIfSet();
	bool ThrowError();

	int GetColumnIndex(const char* sColumnName);
	int GetColumnOrdinal(const char* sColumnName);
	CTypes::CType GetDefaultConversion(SQLTypes::SQLType DataType);

	LPVOID Bind(int iColumnIndex, CTypes::CType conversion);
	LPVOID Bind(const char* sColumnName, CTypes::CType conversion);

	char* BindString(const char* sColumnName);
	char* BindBinary(const char* sColumnName);
	signed int& BindSignedInteger(const char* sColumnName);
	unsigned int& BindUnsignedInteger(const char* sColumnName);
	double& BindDouble(const char* sColumnName);
	float& BindFloat(const char* sColumnName);
	unsigned __int64& BindUnsignedI64(const char* sColumnName);
	signed __int64& BindSignedI64(const char* sColumnName);

	bool ThrowErrors();
	bool ThrowErrors(bool bThrowErrors);

	bool TrimCharData(bool bTrimCharData);
	bool TrimCharData();

	bool ReplaceSingleQuotes(bool bReplaceSingleQuotes);
	bool ReplaceSingleQuotes();

	RECORDSET_COLUMN Columns;

	HSTMT hSTMT;
	SQLLEN RowCount;

	CBoundRecordSet();
	~CBoundRecordSet();
};

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#endif
