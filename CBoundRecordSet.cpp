///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Copyright © NetworkDLS 2002, All rights reserved
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF 
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO 
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A 
// PARTICULAR PURPOSE.
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#ifndef _CBoundRecordSet_CPP
#define _CBoundRecordSet_CPP
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include <Windows.H>
#include <Stdio.H>
#include <SQL.H>
#include <SqlExt.H>

#include "CSQL.H"
#include "CBoundRecordSet.H"
#include "CSQLValue.H"

#ifdef _USE_GLOBAL_MEMPOOL
#include "../CMemPool/CMemPool.H"
extern CMemPool *pMem; //gMem must be defined and initalized elsewhere.
#endif

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/* Returns one of the following:
	CTypes::Bit
	CTypes::Char
	CTypes::Double
	CTypes::Float
	CTypes::SBigInt
	CTypes::SLong
	CTypes::SShort
*/
CTypes::CType CBoundRecordSet::GetDefaultConversion(SQLTypes::SQLType DataType)
{
	CTypes::CType Value = CTypes::Char;

	if (DataType == SQLTypes::BigInt) {
		return CTypes::SBigInt;
	}
	else if (DataType == SQLTypes::Binary) {
		return CTypes::Char;
	}
	else if (DataType == SQLTypes::Bit) {
		return CTypes::Bit;
	}
	else if (DataType == SQLTypes::Char) {
		return CTypes::Char;
	}
	else if (DataType == SQLTypes::DateTime) {
		return CTypes::Char;
	}
	else if (DataType == SQLTypes::Decimal) {
		return CTypes::Double;
	}
	else if (DataType == SQLTypes::Float) {
		return CTypes::Float;
	}
	else if (DataType == SQLTypes::Image) {
		return CTypes::Char;
	}
	else if (DataType == SQLTypes::Int) {
		return CTypes::SLong;
	}
	else if (DataType == SQLTypes::Money) {
		return CTypes::Double;
	}
	else if (DataType == SQLTypes::NChar) {
		return CTypes::Char;
	}
	else if (DataType == SQLTypes::NText) {
		return CTypes::Char;
	}
	else if (DataType == SQLTypes::Numeric) {
		return CTypes::SLong;
	}
	else if (DataType == SQLTypes::NVarChar) {
		return CTypes::Char;
	}
	else if (DataType == SQLTypes::Real) {
		return CTypes::Double;
	}
	else if (DataType == SQLTypes::SmallDateTime) {
		return CTypes::Char;
	}
	else if (DataType == SQLTypes::SmallInt) {
		return CTypes::SShort;
	}
	else if (DataType == SQLTypes::SmallMoney) {
		return CTypes::Double;
	}
	else if (DataType == SQLTypes::Text) {
		return CTypes::Char;
	}
	else if (DataType == SQLTypes::TimeStamp) {
		return CTypes::Char;
	}
	else if (DataType == SQLTypes::TinyInt) {
		return CTypes::SShort;
	}
	else if (DataType == SQLTypes::UniqueIdentifier) {
		return CTypes::Char;
	}
	else if (DataType == SQLTypes::VarBinary) {
		return CTypes::Char;
	}
	else if (DataType == SQLTypes::VarChar) {
		return CTypes::Char;
	}
	else if (DataType == SQLTypes::Variant) {
		return CTypes::Char;
	}

	return Value;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::Reset(void)
{
	bool bResult = true;

	bResult = bResult && (SQLFreeStmt(this->hSTMT, SQL_UNBIND) == SQL_SUCCESS);
	bResult = bResult && (SQLFreeStmt(this->hSTMT, SQL_RESET_PARAMS) == SQL_SUCCESS);

	return bResult;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::Close(void)
{
	bool bResult = true;

	if (this->Columns.Count > 0)
	{
		for (int iPos = 0; iPos < this->Columns.Count; iPos++)
		{
			if (this->Columns.Column[iPos].IsBound)
			{
#ifdef _USE_GLOBAL_MEMPOOL
				pMem->Free(this->Columns.Column[iPos].Data.Buffer);
#else
				free(this->Columns.Column[iPos].Data.Buffer);
#endif
				this->Columns.Column[iPos].Data.Buffer = NULL;
			}
		}

#ifdef _USE_GLOBAL_MEMPOOL
		pMem->Free(this->Columns.Column);
#else
		free(this->Columns.Column);
#endif
	}

	memset(&this->Columns, 0, sizeof(RECORDSET_COLUMN));

	RowCount = 0;

	bResult = bResult && (SQLFreeStmt(this->hSTMT, SQL_UNBIND) == SQL_SUCCESS);
	bResult = bResult && (SQLFreeStmt(this->hSTMT, SQL_CLOSE) == SQL_SUCCESS);
	bResult = bResult && (SQLFreeHandle(SQL_HANDLE_STMT, this->hSTMT) == SQL_SUCCESS);

	this->hSTMT = NULL;

	return bResult;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/*
	This function returns [true] if a row is "fetched", otherwise returns [false].
*/
bool CBoundRecordSet::Fetch(int *iErrorCode)
{
	if (SQL_SUCCEEDED((((int)*iErrorCode) = SQLFetch(this->hSTMT))))
	{
		for (int iCol = 0; iCol < this->Columns.Count; iCol++)
		{
			if (this->Columns.Column[iCol].IsBound)
			{
				this->Columns.Column[iCol].Data.IsNull = (this->Columns.Column[iCol].Data.Size == -1);
			}
		}
		return true;
	}
	else {
		if (*iErrorCode == SQL_NO_DATA)
		{
			return false; //No rows returned, no need to throw an error here.
		}

		this->ThrowErrorIfSet();
		return false;
	}
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::Fetch(void)
{
	int iErrorCode = 0;
	return this->Fetch(&iErrorCode);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/*
	[CSQLValue *] must be saved and [delete]ed after you are finished with it.

	CSQLValue *pValue = Values("Name");
	...
	delete pValue;
*/
CSQLValue *CBoundRecordSet::Values(const char *sColumnName)
{
	int iIndex = this->GetColumnIndex(sColumnName);

	if (iIndex >= 0 && iIndex < this->Columns.Count)
	{
		CSQLValue *pSQLValue = new CSQLValue;
		memset(pSQLValue, 0, sizeof(CSQLValue));

		pSQLValue->Initialize(&this->Columns.Column[iIndex],
			this->TrimCharData(), this->ReplaceSingleQuotes(), this->ThrowErrors());

		return pSQLValue;
	}
	return NULL;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/*
	[CSQLValue *] must be saved and [delete]ed after you are finished with it.

	CSQLValue *pValue = Values("Name");
	...
	delete pValue;
*/
CSQLValue *CBoundRecordSet::Values(int iColumnIndex)
{
	if (iColumnIndex >= 0 && iColumnIndex < this->Columns.Count)
	{
		CSQLValue *pSQLValue = new CSQLValue;
		memset(pSQLValue, 0, sizeof(CSQLValue));

		pSQLValue->Initialize(&this->Columns.Column[iColumnIndex],
			this->TrimCharData(), this->ReplaceSingleQuotes(), this->ThrowErrors());
		return pSQLValue;
	}
	return NULL;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::GetColumnInfo(const int iCol, char *sOutName, int iSzOfOutName, int *iOutColNameLen,
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

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

int CBoundRecordSet::GetColumnIndex(const char *sColumnName)
{
	for (int iPos = 0; iPos < this->Columns.Count; iPos++)
	{
		if (_stricmp(this->Columns.Column[iPos].Name, sColumnName) == 0)
		{
			return iPos;
		}
	}

	return -1;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

int CBoundRecordSet::GetColumnOrdinal(const char *sColumnName)
{
	for (int iPos = 0; iPos < this->Columns.Count; iPos++)
	{
		if (_stricmp(this->Columns.Column[iPos].Name, sColumnName) == 0)
		{
			return this->Columns.Column[iPos].Ordinal;
		}
	}

	return -1;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Bind [any] type.
LPVOID CBoundRecordSet::Bind(int iIndex, CTypes::CType conversion)
{
	if (iIndex < 0 || iIndex >= Columns.Count)
	{
		return NULL;
	}

	LPRECORDSET_COLUMNINFO Pointer = &this->Columns.Column[iIndex];

	if (Pointer->IsBound)
	{
		return Pointer->Data.Buffer;
	}

#ifdef _USE_GLOBAL_MEMPOOL
	Pointer->Data.Buffer = pMem->Allocate(Pointer->MaxSize + 1, 1);
#else
	Pointer->Data.Buffer = calloc(Pointer->MaxSize + 1, 1);
#endif

	SQLBindCol(this->hSTMT, Pointer->Ordinal,
		(SQLSMALLINT)conversion, Pointer->Data.Buffer,
		Pointer->MaxSize + 1, &Pointer->Data.Size);

	Pointer->IsBound = true;

	return Pointer->Data.Buffer;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

LPVOID CBoundRecordSet::Bind(const char *sColumnName, CTypes::CType conversion)
{
	return this->Bind(this->GetColumnIndex(sColumnName), conversion);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Bind [binary] type. (returns char *)
char *CBoundRecordSet::BindBinary(const char *sColumnName)
{
	return (char *) this->Bind(this->GetColumnIndex(sColumnName), CTypes::Binary);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Bind [char] type.
char *CBoundRecordSet::BindString(const char *sColumnName)
{
	return (char *) this->Bind(this->GetColumnIndex(sColumnName), CTypes::Char);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Bind [signed int] type
signed int &CBoundRecordSet::BindSignedInteger(const char *sColumnName)
{
	return *(signed int *)this->Bind(this->GetColumnIndex(sColumnName), CTypes::SLong);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Bind [unsigned int] type
unsigned int &CBoundRecordSet::BindUnsignedInteger(const char *sColumnName)
{
	return *(unsigned int *)this->Bind(this->GetColumnIndex(sColumnName), CTypes::ULong);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Bind [double] type
double &CBoundRecordSet::BindDouble(const char *sColumnName)
{
	return *(double *)this->Bind(this->GetColumnIndex(sColumnName), CTypes::Double);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Bind [float] type
float &CBoundRecordSet::BindFloat(const char *sColumnName)
{
	return *(float *)this->Bind(this->GetColumnIndex(sColumnName), CTypes::Float);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Bind [unsigned __int64] type
unsigned __int64 &CBoundRecordSet::BindUnsignedI64(const char *sColumnName)
{
	return *(unsigned __int64 *)this->Bind(this->GetColumnIndex(sColumnName), CTypes::UBigInt);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Bind [__int64] type
signed __int64 &CBoundRecordSet::BindSignedI64(const char *sColumnName)
{
	return *(signed __int64 *)this->Bind(this->GetColumnIndex(sColumnName), CTypes::SBigInt);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

CBoundRecordSet::CBoundRecordSet()
{
	this->pPublicData = NULL;
	this->pErrorHandler = NULL;

	this->hSTMT = NULL;
	this->RowCount = 0;
	memset(&this->Columns, 0, sizeof(RECORDSET_COLUMN));

	this->TrimCharData(true);
	this->ReplaceSingleQuotes(true);
	this->ThrowErrors(true);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

CBoundRecordSet::~CBoundRecordSet()
{
	this->hSTMT = NULL;
	this->RowCount = 0;
	this->Columns.Count = 0;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::GetSQLState(char *sSQLState)
{
	return SQL_SUCCEEDED(SQLError(NULL, NULL,
		this->hSTMT, (SQLCHAR *)sSQLState, NULL, NULL, NULL, NULL));
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::GetErrorMessage(int *iOutErr, char *sOutError, const int iErrBufSz)
{
	SQLCHAR     sSQLState[20];
	SQLSMALLINT iOutErrorMsgSz;

	return SQL_SUCCEEDED(SQLError(NULL, NULL, this->hSTMT, sSQLState,
		(SQLINTEGER *)iOutErr, (SQLCHAR *)sOutError, iErrBufSz, &iOutErrorMsgSz));
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::ThrowErrorIfSet(void)
{
	if (this->ThrowErrors() == true)
	{
		return this->ThrowError();
	}

	return false;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::ThrowError(void)
{
	char sErrorMsg[2048];
	int iNativeError = 0;

	if (this->GetErrorMessage(&iNativeError, sErrorMsg, sizeof(sErrorMsg)))
	{
		if (this->pErrorHandler)
		{
			this->pErrorHandler(this, "CBoundRecordSet", sErrorMsg, iNativeError);
		}
#ifdef _DEBUG
		if (IsDebuggerPresent())
		{
			__debugbreak();
		}
#endif
		return true;
	}

	return false;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void CBoundRecordSet::SetErrorHandler(ErrorHandler pHandler)
{
	this->pErrorHandler = pHandler;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::ThrowErrors(void)
{
	return this->bcThrowErrors;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::ThrowErrors(bool bThrowErrors)
{
	bool bOldValue = this->bcThrowErrors;
	this->bcThrowErrors = bThrowErrors;
	return bOldValue;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::TrimCharData(void)
{
	return this->bcTrimCharData;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::TrimCharData(bool bTrimCharData)
{
	bool bOldValue = this->bcTrimCharData;
	this->bcTrimCharData = bTrimCharData;
	return bOldValue;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::ReplaceSingleQuotes(void)
{
	return this->bcReplaceSingleQuotes;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::ReplaceSingleQuotes(bool bReplaceSingleQuotes)
{
	bool bOldValue = this->bcReplaceSingleQuotes;
	this->bcReplaceSingleQuotes = bReplaceSingleQuotes;
	return bOldValue;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#endif
