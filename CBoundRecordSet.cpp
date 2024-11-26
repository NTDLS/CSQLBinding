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
	switch (DataType)
	{
	case SQLTypes::BigInt:
		return CTypes::SBigInt;

	case SQLTypes::UniqueIdentifier:
	case SQLTypes::VarBinary:
	case SQLTypes::VarChar:
	case SQLTypes::Variant:
	case SQLTypes::NChar:
	case SQLTypes::NText:
	case SQLTypes::NVarChar:
	case SQLTypes::Image:
	case SQLTypes::Char:
	case SQLTypes::DateTime:
	case SQLTypes::Binary:
		//case SQLTypes::SmallDateTime: //Same as DateTime
	case SQLTypes::Text:
		//case SQLTypes::TimeStamp: //Same as DateTime
		return CTypes::Char;

	case SQLTypes::Bit:
		return CTypes::Bit;

	case SQLTypes::Real:
		//case SQLTypes::SmallMoney:  //Same as Money
		//case SQLTypes::Decimal: //Same as Money
	case SQLTypes::Money:
		return CTypes::Double;

	case SQLTypes::Float:
		return CTypes::Float;

	case SQLTypes::Int:
	case SQLTypes::Numeric:
		return CTypes::SLong;

	case SQLTypes::SmallInt:
	case SQLTypes::TinyInt:
		return CTypes::SShort;
		return CTypes::SShort;

	default:
		return CTypes::Char;
	}
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::Reset()
{
	bool bResult = true;

	bResult = bResult && (SQLFreeStmt(this->hSTMT, SQL_UNBIND) == SQL_SUCCESS);
	bResult = bResult && (SQLFreeStmt(this->hSTMT, SQL_RESET_PARAMS) == SQL_SUCCESS);

	return bResult;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::Close()
{
	bool bResult = true;

	if (this->Columns.Count > 0)
	{
		for (int iColumnIndex = 0; iColumnIndex < this->Columns.Count; iColumnIndex++)
		{
			if (this->Columns.Column[iColumnIndex].IsBound)
			{
				free(this->Columns.Column[iColumnIndex].Data.Buffer);
				this->Columns.Column[iColumnIndex].Data.Buffer = NULL;
			}
		}

		free(this->Columns.Column);
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
bool CBoundRecordSet::Fetch(int* iErrorCode)
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

bool CBoundRecordSet::Fetch()
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
CSQLValue* CBoundRecordSet::Values(const char* sColumnName)
{
	int iColumnIndex = this->GetColumnIndex(sColumnName);

	if (iColumnIndex >= 0 && iColumnIndex < this->Columns.Count)
	{
		CSQLValue* pSQLValue = new CSQLValue;
		memset(pSQLValue, 0, sizeof(CSQLValue));

		pSQLValue->Initialize(&this->Columns.Column[iColumnIndex],
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
CSQLValue* CBoundRecordSet::Values(int iColumnIndex)
{
	if (iColumnIndex >= 0 && iColumnIndex < this->Columns.Count)
	{
		CSQLValue* pSQLValue = new CSQLValue;
		memset(pSQLValue, 0, sizeof(CSQLValue));

		pSQLValue->Initialize(&this->Columns.Column[iColumnIndex],
			this->TrimCharData(), this->ReplaceSingleQuotes(), this->ThrowErrors());

		return pSQLValue;
	}
	return NULL;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::GetColumnInfo(const int iColumnIndex, char* sOutName, int iSzOfOutName, int* iOutColNameLen,
	int* iOutDataType, SQLULEN* iOutColSize, int* iNumOfDeciPlaces, int* iColNullable)
{
	if (SQL_SUCCEEDED(SQLDescribeCol(this->hSTMT, iColumnIndex, (SQLCHAR*)sOutName, (SQLSMALLINT)iSzOfOutName,
		(SQLSMALLINT*)iOutColNameLen, (SQLSMALLINT*)iOutDataType, iOutColSize,
		(SQLSMALLINT*)iNumOfDeciPlaces, (SQLSMALLINT*)iColNullable)))
	{
		return true;
	}
	else return false;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

int CBoundRecordSet::GetColumnIndex(const char* sColumnName)
{
	for (int iColumnIndex = 0; iColumnIndex < this->Columns.Count; iColumnIndex++)
	{
		if (_stricmp(this->Columns.Column[iColumnIndex].Name, sColumnName) == 0)
		{
			return iColumnIndex;
		}
	}

	return -1;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

int CBoundRecordSet::GetColumnOrdinal(const char* sColumnName)
{
	for (int iColumnIndex = 0; iColumnIndex < this->Columns.Count; iColumnIndex++)
	{
		if (_stricmp(this->Columns.Column[iColumnIndex].Name, sColumnName) == 0)
		{
			return this->Columns.Column[iColumnIndex].Ordinal;
		}
	}

	return -1;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Bind [any] type.
LPVOID CBoundRecordSet::Bind(int iColumnIndex, CTypes::CType conversion)
{
	if (iColumnIndex < 0 || iColumnIndex >= Columns.Count)
	{
		return NULL;
	}

	LPRECORDSET_COLUMNINFO Pointer = &this->Columns.Column[iColumnIndex];

	if (Pointer->IsBound)
	{
		return Pointer->Data.Buffer;
	}

	Pointer->Data.Buffer = calloc(Pointer->MaxSize + 1, 1);

	SQLBindCol(this->hSTMT, Pointer->Ordinal,
		(SQLSMALLINT)conversion, Pointer->Data.Buffer,
		Pointer->MaxSize + 1, &Pointer->Data.Size);

	Pointer->IsBound = true;

	return Pointer->Data.Buffer;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

LPVOID CBoundRecordSet::Bind(const char* sColumnName, CTypes::CType conversion)
{
	return this->Bind(this->GetColumnIndex(sColumnName), conversion);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Bind [binary] type. (returns char *)
char* CBoundRecordSet::BindBinary(const char* sColumnName)
{
	return (char*)this->Bind(this->GetColumnIndex(sColumnName), CTypes::Binary);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Bind [char] type.
char* CBoundRecordSet::BindString(const char* sColumnName)
{
	return (char*)this->Bind(this->GetColumnIndex(sColumnName), CTypes::Char);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Bind [signed int] type
signed int& CBoundRecordSet::BindSignedInteger(const char* sColumnName)
{
	return *(signed int*)this->Bind(this->GetColumnIndex(sColumnName), CTypes::SLong);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Bind [unsigned int] type
unsigned int& CBoundRecordSet::BindUnsignedInteger(const char* sColumnName)
{
	return *(unsigned int*)this->Bind(this->GetColumnIndex(sColumnName), CTypes::ULong);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Bind [double] type
double& CBoundRecordSet::BindDouble(const char* sColumnName)
{
	return *(double*)this->Bind(this->GetColumnIndex(sColumnName), CTypes::Double);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Bind [float] type
float& CBoundRecordSet::BindFloat(const char* sColumnName)
{
	return *(float*)this->Bind(this->GetColumnIndex(sColumnName), CTypes::Float);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Bind [unsigned __int64] type
unsigned __int64& CBoundRecordSet::BindUnsignedI64(const char* sColumnName)
{
	return *(unsigned __int64*)this->Bind(this->GetColumnIndex(sColumnName), CTypes::UBigInt);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Bind [__int64] type
signed __int64& CBoundRecordSet::BindSignedI64(const char* sColumnName)
{
	return *(signed __int64*)this->Bind(this->GetColumnIndex(sColumnName), CTypes::SBigInt);
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

bool CBoundRecordSet::GetSQLState(char* sSQLState)
{
	return SQL_SUCCEEDED(SQLError(NULL, NULL,
		this->hSTMT, (SQLCHAR*)sSQLState, NULL, NULL, NULL, NULL));
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::GetErrorMessage(int* iOutErr, char* sOutError, const int iErrBufSz)
{
	SQLCHAR sSQLState[20];
	SQLSMALLINT iOutErrorMsgSz;

	return SQL_SUCCEEDED(SQLError(NULL, NULL, this->hSTMT, sSQLState,
		(SQLINTEGER*)iOutErr, (SQLCHAR*)sOutError, iErrBufSz, &iOutErrorMsgSz));
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::ThrowErrorIfSet()
{
	if (this->ThrowErrors() == true)
	{
		return this->ThrowError();
	}

	return false;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CBoundRecordSet::ThrowError()
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

/// <summary>
/// Sets the error handler callback.
/// </summary>
void CBoundRecordSet::SetErrorHandler(ErrorHandler pHandler)
{
	this->pErrorHandler = pHandler;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/// <summary>
/// Gets the Throw Errors mode.
/// </summary>
bool CBoundRecordSet::ThrowErrors()
{
	return this->bcThrowErrors;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/// <summary>
/// Sets the Throw Errors mode.
/// </summary>
bool CBoundRecordSet::ThrowErrors(bool bThrowErrors)
{
	bool bOldValue = this->bcThrowErrors;
	this->bcThrowErrors = bThrowErrors;
	return bOldValue;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/// <summary>
/// Gets the Trim Char Data mode.
/// </summary>
bool CBoundRecordSet::TrimCharData()
{
	return this->bcTrimCharData;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/// <summary>
/// Sets the Trim Char Data mode.
/// </summary>
bool CBoundRecordSet::TrimCharData(bool bTrimCharData)
{
	bool bOldValue = this->bcTrimCharData;
	this->bcTrimCharData = bTrimCharData;
	return bOldValue;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/// <summary>
/// Gets the Replace Single Quotes mode.
/// </summary>
bool CBoundRecordSet::ReplaceSingleQuotes()
{
	return this->bcReplaceSingleQuotes;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/// <summary>
/// Sets the Replace Single Quotes mode.
/// </summary>
bool CBoundRecordSet::ReplaceSingleQuotes(bool bReplaceSingleQuotes)
{
	bool bOldValue = this->bcReplaceSingleQuotes;
	this->bcReplaceSingleQuotes = bReplaceSingleQuotes;
	return bOldValue;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#endif
