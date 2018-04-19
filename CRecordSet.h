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

typedef int(*ErrorHandler)(void *pCaller, const char *sSource, const char *sErrorMsg, const int iErrorNumber);

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

class CRecordSet {

private:
	ErrorHandler pErrorHandler;

	bool bcTrimCharData;
	bool bcReplaceSingleQuotes;
	bool bcThrowErrors;

public:
	void *pPublicData;

	void SetErrorHandler(ErrorHandler pHandler);

	bool Close(void);
	bool Fetch(void);
	bool Fetch(int *iErrorCode);
	bool Reset(void);

	bool GetData(const unsigned short iCol, short iType, void *pvBuf, SQLLEN iBufSz, SQLLEN *piIndicator);
	bool GetData(const unsigned short iCol, short iType, void * pvBuf, SQLLEN iBufSz, SQLLEN *piIndicator, short *piResult);
	bool GetColumnInfo(const int iCol, char *sOutName, int iSzOfOutName,
		int *iOutColNameLen, int *ioutDataType, SQLULEN *iOutColSize, int *iNumOfDeciPlaces, int *iColNullable);
	bool BinColumnEx(const int iCol, char *sBuf, const SQLLEN iBufSz, SQLLEN *iOutLen);
	bool sColumnEx(const int iCol, char *sBuf, const SQLLEN iBufSz, SQLLEN *iOutLen);
	bool lColumnEx(const int iCol, long *plOutVal);
	long lColumn(const int iCol);
	bool dColumnEx(const int iCol, double *pdOutVal);
	double dColumn(const int iCol);
	bool fColumnEx(const int iCol, float *pfOutVal);
	float fColumn(const int iCol);
	bool GetErrorMessage(int *iOutErr, char *sOutError, const int iErrBufSz);
	bool GetSQLState(char *sSQLState);
	bool ThrowErrorIfSet(void);
	bool ThrowError(void);

	SQLLEN RTrim(char *sData, SQLLEN iDataSz);
	void ReplaceSingleQuotes(char *sData, SQLLEN iDataSz);

	bool ThrowErrors(void);
	bool ThrowErrors(bool bThrowErrors);

	bool TrimCharData(bool bTrimCharData);
	bool TrimCharData(void);

	bool ReplaceSingleQuotes(bool bReplaceSingleQuotes);
	bool ReplaceSingleQuotes(void);

	HSTMT hSTMT;
	SQLLEN RowCount;
	SQLLEN ColumnCount;

	CRecordSet();
	~CRecordSet();
};

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#endif
