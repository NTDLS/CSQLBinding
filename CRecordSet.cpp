///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Copyright © NetworkDLS 2002, All rights reserved
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF 
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO 
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A 
// PARTICULAR PURPOSE.
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#ifndef _CRECORDSET_CPP
#define _CRECORDSET_CPP
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include <Windows.H>
#include <Stdio.H>
#include <SQL.H>
#include <SqlExt.H>

#include "CSQL.H"
#include "CRecordSet.H"

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void CRecordSet::ReplaceSingleQuotes(char *sData, SQLLEN iDataSz)
{
	int iRPos = 0;

	while (iRPos < iDataSz)
	{
		if (sData[iRPos] == '\'')
		{
			sData[iRPos] = '`';
		}
		iRPos++;
	}
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

SQLLEN CRecordSet::RTrim(char *sData, SQLLEN iDataSz)
{
	SQLLEN iPos = iDataSz;

	if (iDataSz == 0)
		return iDataSz;//Return the new data length.

	if (sData[iDataSz - 1] != ' ')
		return iDataSz;//Return the new data length.

	iPos--;

	while (iPos != 0 && sData[iPos] == ' ')
		iPos--;

	if (sData[iPos] != ' ')
		iPos++;

	sData[iPos] = '\0';

	return iPos; //Return the new data length.
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::Reset(void)
{
	bool bResult = false;

	bResult = (SQLFreeStmt(this->hSTMT, SQL_UNBIND) == SQL_SUCCESS);
	bResult = (SQLFreeStmt(this->hSTMT, SQL_RESET_PARAMS) == SQL_SUCCESS);

	return bResult;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::Close(void)
{
	bool bResult = (SQLFreeHandle(SQL_HANDLE_STMT, this->hSTMT) == SQL_SUCCESS);

	this->RowCount = 0;
	this->ColumnCount = 0;
	this->hSTMT = NULL;

	return bResult;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::Fetch(int *iErrorCode)
{
	if (SQL_SUCCEEDED((*iErrorCode = SQLFetch(this->hSTMT))))
	{
		return true;
	}
	else {
		this->ThrowErrorIfSet();
		return false;
	}
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::Fetch(void)
{
	int iErrorCode = 0;
	return this->Fetch(&iErrorCode);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::GetData(const unsigned short iCol, short iType,
	void * pvBuf, SQLLEN iBufSz, SQLLEN *piIndicator, short *piResult)
{
	//SQL_SUCCESS, SQL_SUCCESS_WITH_INFO,
	//	SQL_NO_DATA, SQL_STILL_EXECUTING,
	//	SQL_ERROR, or SQL_INVALID_HANDLE

	SQLRETURN iResult = SQLGetData(this->hSTMT,
		(SQLUSMALLINT)iCol,
		(SQLSMALLINT)iType,
		(SQLPOINTER)pvBuf,
		iBufSz,
		piIndicator);
	if (piResult)
	{
		*piResult = iResult;
	}
	if (SQL_SUCCEEDED(iResult))
	{
		return true;
	}
	else {
		ThrowErrorIfSet();
		return false;
	}
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::GetData(const unsigned short iCol, short iType,
	void *pvBuf, SQLLEN iBufSz, SQLLEN *piIndicator)
{
	return this->GetData(iCol, iType, pvBuf, iBufSz, piIndicator, NULL);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::GetColumnInfo(const int iCol, char *sOutName, int iSzOfOutName, int *iOutColNameLen,
	int *ioutDataType, SQLULEN *iOutColSize, int *iNumOfDeciPlaces, int *iColNullable)
{
	if (SQL_SUCCEEDED(SQLDescribeCol(this->hSTMT, iCol, (SQLCHAR *)sOutName, (SQLSMALLINT)iSzOfOutName,
		(SQLSMALLINT *)iOutColNameLen, (SQLSMALLINT *)ioutDataType, iOutColSize,
		(SQLSMALLINT *)iNumOfDeciPlaces, (SQLSMALLINT *)iColNullable)))
	{
		return true;
	}
	else return false;
}


//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::sColumnEx(const int iCol, char *sBuf, const SQLLEN iBufSz, SQLLEN *iOutLen)
{
	SQLLEN iLen = 0;

	if (SQL_SUCCEEDED(SQLGetData(this->hSTMT, iCol, SQL_C_CHAR, sBuf, iBufSz, &iLen)))
	{
		if (iLen > iBufSz)
		{
			iLen = iBufSz;
		}

		if (this->ReplaceSingleQuotes())
		{
			this->ReplaceSingleQuotes(sBuf, iLen);
		}

		if (this->TrimCharData())
		{
			iLen = RTrim(sBuf, iLen);
		}

		if (iOutLen)
		{
			*iOutLen = iLen;
		}

		return true;
	}
	else {
		this->ThrowErrorIfSet();
		return false;
	}
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::lColumnEx(const int iCol, long *plOutVal)
{
	SQLLEN Indicator = 0;

	int i = sizeof(Indicator);

	if (SQL_SUCCEEDED(SQLGetData(this->hSTMT, iCol, SQL_C_LONG, plOutVal, sizeof(long), &Indicator)))
	{
		return true;
	}
	else {
		this->ThrowErrorIfSet();
		return false;
	}
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

long CRecordSet::lColumn(const int iCol)
{
	long lTemp = 0;

	if (this->lColumnEx(iCol, &lTemp))
	{
		return lTemp;
	}
	else return -1;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::dColumnEx(const int iCol, double *pdOutVal)
{
	SQLLEN Indicator = 0;

	if (SQL_SUCCEEDED(SQLGetData(this->hSTMT, iCol, SQL_DOUBLE, pdOutVal, sizeof(double), &Indicator)))
	{
		return true;
	}
	else {
		this->ThrowErrorIfSet();
		return false;
	}
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

double CRecordSet::dColumn(const int iCol)
{
	double dTemp = 0;

	if (this->dColumnEx(iCol, &dTemp))
	{
		return dTemp;
	}
	else return -1;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::fColumnEx(const int iCol, float *pfOutVal)
{
	SQLLEN Indicator = 0;

	if (SQL_SUCCEEDED(SQLGetData(this->hSTMT, iCol, SQL_FLOAT, pfOutVal, sizeof(float), &Indicator)))
	{
		return true;
	}
	else {
		this->ThrowErrorIfSet();
		return false;
	}
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

float CRecordSet::fColumn(const int iCol)
{
	float fTemp = 0;

	if (this->fColumnEx(iCol, &fTemp))
	{
		return fTemp;
	}
	else return -1;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::BinColumnEx(const int iCol, char *sBuf, const SQLLEN iBufSz, SQLLEN *iOutLen)
{
	SQLLEN iLen = 0;

	if (SQL_SUCCEEDED(SQLGetData(this->hSTMT, iCol, SQL_C_BINARY, sBuf, iBufSz, &iLen)))
	{
		if (iLen > iBufSz)
		{
			iLen = iBufSz;
		}

		if (this->ReplaceSingleQuotes())
		{
			this->ReplaceSingleQuotes(sBuf, iLen);
		}

		if (this->TrimCharData())
		{
			iLen = RTrim(sBuf, iLen);
		}

		if (iOutLen)
		{
			*iOutLen = iLen;
		}

		return true;
	}
	else {
		this->ThrowErrorIfSet();
		return false;
	}
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

CRecordSet::CRecordSet()
{
	this->hSTMT = NULL;
	this->RowCount = 0;
	this->ColumnCount = 0;

	this->TrimCharData(true);
	this->ReplaceSingleQuotes(true);
	this->ThrowErrors(true);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

CRecordSet::~CRecordSet()
{
	this->hSTMT = NULL;
	this->RowCount = 0;
	this->ColumnCount = 0;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::GetSQLState(char *sSQLState)
{
	return SQL_SUCCEEDED(SQLError(NULL, NULL,
		this->hSTMT, (SQLCHAR *)sSQLState, NULL, NULL, NULL, NULL));
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::GetErrorMessage(int *iOutErr, char *sOutError, const int iErrBufSz)
{
	SQLCHAR     sSQLState[20];
	SQLSMALLINT iOutErrorMsgSz;

	return SQL_SUCCEEDED(SQLError(NULL, NULL, this->hSTMT, sSQLState,
		(SQLINTEGER *)iOutErr, (SQLCHAR *)sOutError, iErrBufSz, &iOutErrorMsgSz));
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::ThrowErrorIfSet(void)
{
	if (this->ThrowErrors())
	{
		return this->ThrowError();
	}

	return false;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::ThrowError(void)
{
	char sErrorMsg[2048];
	int iNativeError = 0;

	if (this->GetErrorMessage(&iNativeError, sErrorMsg, sizeof(sErrorMsg)))
	{
		if (this->pErrorHandler)
		{
			this->pErrorHandler(this, "CRecordSet", sErrorMsg, iNativeError);
		}
		return true;
	}

	return false;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void CRecordSet::SetErrorHandler(ErrorHandler pHandler)
{
	this->pErrorHandler = pHandler;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::ThrowErrors(void)
{
	return this->bcThrowErrors;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::ThrowErrors(bool bThrowErrors)
{
	bool bOldValue = this->bcThrowErrors;
	this->bcThrowErrors = bThrowErrors;
	return bOldValue;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::TrimCharData(void)
{
	return this->bcTrimCharData;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::TrimCharData(bool bTrimCharData)
{
	bool bOldValue = this->bcTrimCharData;
	this->bcTrimCharData = bTrimCharData;
	return bOldValue;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::ReplaceSingleQuotes(void)
{
	return this->bcReplaceSingleQuotes;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CRecordSet::ReplaceSingleQuotes(bool bReplaceSingleQuotes)
{
	bool bOldValue = this->bcReplaceSingleQuotes;
	this->bcReplaceSingleQuotes = bReplaceSingleQuotes;
	return bOldValue;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#endif
