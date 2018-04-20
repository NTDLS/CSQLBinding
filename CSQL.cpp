///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Copyright © NetworkDLS 2002, All rights reserved
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF 
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO 
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A 
// PARTICULAR PURPOSE.
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#ifndef _SQLCLASS_CPP
#define _SQLCLASS_CPP
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include <Windows.H>
#include <Stdio.H>

#include "CSQL.H"
#include "CBoundRecordSet.H"
#include "CRecordSet.H"

#ifdef _USE_GLOBAL_MEMPOOL
#include "../NSWFL/NSWFL.h"
extern NSWFL::Memory::MemoryPool *pMem; //gMem must be defined and initalized elsewhere.
#endif

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

int DefaultSQLErrorHandler(void *pCaller, const char *sSource, const char *sErrorMsg, const int iErrorNumber)
{
	MessageBox(GetActiveWindow(), sErrorMsg, sSource, MB_ICONERROR);
	return 0;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

CSQL::CSQL(void)
{
	this->pErrorHandler = NULL;
	this->pPublicData = NULL;
	this->LastConnectionString = NULL;

	this->bConnected = false;
	this->Disconnect();
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

CSQL::~CSQL(void)
{
	this->Disconnect();
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::IsConnected(void)
{
	return this->bConnected;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

const char *CSQL::ConnectionString(void)
{
	return this->LastConnectionString;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Connect(const char *sServer, const char *sDatabase, ErrorHandler pHandler)
{
	SQLCONNECTIONSTRING CS;
	memset(&CS, 0, sizeof(CS));

	CS.bUseTrustedConnection = true;
	strcpy_s(CS.sDriver, sizeof(CS.sDriver), "{SQL Server}");
	strcpy_s(CS.sDatabase, sizeof(CS.sDatabase), sDatabase);
	strcpy_s(CS.sServer, sizeof(CS.sServer), sServer);

	ErrorHandler plHandler = pHandler;
	if (!plHandler)
	{
		plHandler = DefaultSQLErrorHandler;
	}

	return this->Connect(&CS, plHandler);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Connect(const char *sServer, const char *sDatabase)
{
	SQLCONNECTIONSTRING CS;
	memset(&CS, 0, sizeof(CS));

	CS.bUseTrustedConnection = true;
	strcpy_s(CS.sDriver, sizeof(CS.sDriver), "{SQL Server}");
	strcpy_s(CS.sDatabase, sizeof(CS.sDatabase), sDatabase);
	strcpy_s(CS.sServer, sizeof(CS.sServer), sServer);

	return this->Connect(&CS, &DefaultSQLErrorHandler);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Connect(const char *sServer, const char *sDatabase,
	const char *sUser, const char *sPassword)
{
	return this->Connect(sServer, sDatabase, sUser, sPassword, DefaultSQLErrorHandler);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Connect(const char *sServer, const char *sDatabase,
	const char *sUser, const char *sPassword, ErrorHandler pHandler)
{
	SQLCONNECTIONSTRING CS;
	memset(&CS, 0, sizeof(CS));

	CS.bUseTrustedConnection = false;
	strcpy_s(CS.sDriver, sizeof(CS.sDriver), "{SQL Server}");
	strcpy_s(CS.sDatabase, sizeof(CS.sDatabase), sDatabase);
	strcpy_s(CS.sServer, sizeof(CS.sServer), sServer);
	strcpy_s(CS.sUID, sizeof(CS.sUID), sUser);
	strcpy_s(CS.sPwd, sizeof(CS.sPwd), sPassword);

	ErrorHandler plHandler = pHandler;
	if (!plHandler)
	{
		plHandler = DefaultSQLErrorHandler;
	}

	return this->Connect(&CS, plHandler);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Connect(const char *sConnectionString)
{
	return this->Connect(sConnectionString, &DefaultSQLErrorHandler);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Connect(LPSQLCONNECTIONSTRING SQLCon, ErrorHandler pHandler)
{
	char sConnectionString[1024];

	if (!this->BuildConnectionString(SQLCon, sConnectionString, sizeof(sConnectionString)))
	{
		return false;
	}

	return this->Connect(sConnectionString, pHandler);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Connect(LPSQLCONNECTIONSTRING SQLCon)
{
	return this->Connect(SQLCon, &DefaultSQLErrorHandler);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Connect(const char *sConnectionString, ErrorHandler pHandler)
{
	SQLRETURN Result;
	SQLCHAR sConStrOut[1024];
	SQLSMALLINT iConStrOutSz = 0;

	this->SetErrorHandler(pHandler);

	this->bConnected = false;
	this->Disconnect();

	this->LastConnectionString = _strdup(sConnectionString);

	// Allocate the environment handle
	Result = SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, &this->hcSQLEnvironment);
	if (!SQL_SUCCEEDED(Result))
	{
		ThrowErrorIfSet(NULL);
		return false;
	}

	// Set the environment attributes
	Result = SQLSetEnvAttr(this->hcSQLEnvironment, SQL_ATTR_ODBC_VERSION, (void *)SQL_OV_ODBC3, 0);
	if (!SQL_SUCCEEDED(Result))
	{
		ThrowErrorIfSet(NULL);
		SQLFreeHandle(SQL_HANDLE_ENV, this->hcSQLEnvironment);
		return false;
	}

	// Allocate the connection handle
	Result = SQLAllocHandle(SQL_HANDLE_DBC, this->hcSQLEnvironment, &this->hcSQLConnection);
	if (!SQL_SUCCEEDED(Result))
	{
		ThrowErrorIfSet(NULL);
		SQLFreeHandle(SQL_HANDLE_DBC, this->hcSQLConnection);
		SQLFreeHandle(SQL_HANDLE_ENV, this->hcSQLEnvironment);
		return false;
	}
	
	// Set login timeout to 5 seconds.
	if (!SQL_SUCCEEDED(SQLSetConnectAttr(this->hcSQLConnection, SQL_ATTR_LOGIN_TIMEOUT, (SQLPOINTER)5, NULL)))
	{
		ThrowErrorIfSet(NULL);
		SQLFreeHandle(SQL_HANDLE_DBC, this->hcSQLConnection);
		SQLFreeHandle(SQL_HANDLE_ENV, this->hcSQLEnvironment);
		return false;
	}
	
	if (!SQL_SUCCEEDED(SQLSetConnectAttr(this->hcSQLConnection, SQL_ATTR_CURSOR_TYPE, (SQLPOINTER)this->icCursorType, NULL)))
	{
		ThrowErrorIfSet(NULL);
		SQLFreeHandle(SQL_HANDLE_DBC, this->hcSQLConnection);
		SQLFreeHandle(SQL_HANDLE_ENV, this->hcSQLEnvironment);
		return false;
	}

	if (!SQL_SUCCEEDED(SQLSetConnectAttr(this->hcSQLConnection, SQL_ATTR_CONCURRENCY, (SQLPOINTER)SQL_CONCUR_READ_ONLY, NULL)))
	{
		ThrowErrorIfSet(NULL);
		SQLFreeHandle(SQL_HANDLE_DBC, this->hcSQLConnection);
		SQLFreeHandle(SQL_HANDLE_ENV, this->hcSQLEnvironment);
		return false;
	}

	if (bcUseBulkOperations)
	{
		if (!SQL_SUCCEEDED(SQLSetConnectAttr(this->hcSQLConnection,
			SQL_COPT_SS_BCP, (void *)SQL_BCP_ON, SQL_IS_INTEGER)))
		{
			ThrowErrorIfSet(NULL);
			SQLFreeHandle(SQL_HANDLE_DBC, this->hcSQLConnection);
			SQLFreeHandle(SQL_HANDLE_ENV, this->hcSQLEnvironment);
			return false;
		}
	}

	// Connect
	Result = SQLDriverConnect(
		this->hcSQLConnection,        // Connection handle.
		NULL,                         // Window handle.
		(SQLCHAR *)sConnectionString, // Input connect string.
		SQL_NTS,                      // Null-terminated string.
		sConStrOut,                   // Address of output buffer.
		sizeof(sConStrOut),           // Size of output buffer.
		&iConStrOutSz,                // Address of output length.
		SQL_DRIVER_NOPROMPT
		);

	if (!SQL_SUCCEEDED(Result))
	{
		ThrowErrorIfSet(NULL);
		SQLFreeHandle(SQL_HANDLE_DBC, this->hcSQLConnection);
		SQLFreeHandle(SQL_HANDLE_ENV, this->hcSQLEnvironment);
		return false;
	}

	this->bConnected = true;

	return true;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::BuildConnectionString(LPSQLCONNECTIONSTRING SQLCon, char *sOut, int iOutSz)
{
	char sTemp[1024];

	//-----------------------------------------------------------------------------------------------
	// Important note!
	//-----------------------------------------------------------------------------------------------
	//		When connecting through the SQLOLEDB provider use the syntax Network Library=dbmssocn
	//			and when connecting through MSDASQL provider use the syntax Network=dbmssocn
	//-----------------------------------------------------------------------------------------------

	//Driver: {SQL Server}

	sprintf_s(sOut, iOutSz, "DRIVER=%s;SERVER=%s;", SQLCon->sDriver, SQLCon->sServer);

	if (strlen(SQLCon->sDatabase) > 0)
	{
		sprintf_s(sTemp, sizeof(sTemp), "DATABASE=%s;", SQLCon->sDatabase);
		strcat_s(sOut, iOutSz, sTemp);
	}

	if (strlen(SQLCon->sApplicationName) > 0)
	{
		sprintf_s(sTemp, sizeof(sTemp), "APP=%s;", SQLCon->sApplicationName);
		strcat_s(sOut, iOutSz, sTemp);
	}

	if (SQLCon->bUseTCPIPConnection)
	{
		sprintf_s(sTemp, sizeof(sTemp), "NETWORK=DBMSSOCN;ADDRESS=%s,%d;", SQLCon->sServer, SQLCon->iPort);
		strcat_s(sOut, iOutSz, sTemp);
	}

	if (SQLCon->bUseMARS)
	{
		strcat_s(sOut, iOutSz, "MARS_Connection=yes;");
	}

	if (SQLCon->bUseTrustedConnection)
	{
		strcat_s(sOut, iOutSz, "TRUSTED_CONNECTION=yes;");
	}
	else {
		sprintf_s(sTemp, sizeof(sTemp), "UID=%s;PWD=%s;TRUSTED_CONNECTION=no;", SQLCon->sUID, SQLCon->sPwd);
		strcat_s(sOut, iOutSz, sTemp);
	}

	return true;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::GetSQLState(char *sSQLState, HSTMT hStmt)
{
	return SQL_SUCCEEDED(SQLError(this->hcSQLEnvironment, this->hcSQLConnection,
		hStmt, (SQLCHAR *)sSQLState, NULL, NULL, NULL, NULL));
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::GetErrorMessage(int *iOutErr, char *sOutError, const int iErrBufSz, HSTMT hStmt)
{
	SQLCHAR     sSQLState[20];
	SQLSMALLINT iOutErrorMsgSz;

	return SQL_SUCCEEDED(SQLError(this->hcSQLEnvironment, this->hcSQLConnection,
		hStmt, sSQLState, (SQLINTEGER *)iOutErr, (SQLCHAR *)sOutError, iErrBufSz, &iOutErrorMsgSz));
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::ThrowErrorIfSet(HSTMT hStmt)
{
	if (this->ThrowErrors())
	{
		return this->ThrowError(hStmt);
	}

	return false;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::ThrowError(HSTMT hStmt)
{
	char sErrorMsg[2048];
	int iNativeError = 0;

	if (GetErrorMessage(&iNativeError, sErrorMsg, sizeof(sErrorMsg), hStmt))
	{
		if (this->pErrorHandler)
		{
			this->pErrorHandler(this, "CSQL", sErrorMsg, iNativeError);
		}
#ifdef _DEBUG
		if (IsDebuggerPresent())
		{
			//__debugbreak();
		}
#endif
		return true;
	}

	return false;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void CSQL::Disconnect(void)
{
	if (this->bConnected)
	{
		SQLDisconnect(this->hcSQLConnection);
		SQLFreeHandle(SQL_HANDLE_DBC, this->hcSQLConnection);
		SQLFreeHandle(SQL_HANDLE_ENV, this->hcSQLEnvironment);
	}

	if (this->LastConnectionString)
	{
		free(this->LastConnectionString);
	}

	this->hcSQLConnection = NULL;
	this->hcSQLEnvironment = NULL;

	this->bConnected = false;
	this->ThrowErrors(true);
	this->bcUseBulkOperations = false;
	this->icTimeout = 30;
	this->icCursorType = SQL_CURSOR_STATIC; //Slow but returns row count.
	//this->icCursorType = SQL_CURSOR_FORWARD_ONLY; //Fast but doesnt return row count.
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

short CSQL::Execute(HSTMT hSTMT)
{
	SQLRETURN Result = SQLExecute(hSTMT);
	if (!SQL_SUCCEEDED(Result))
	{
		this->ThrowErrorIfSet(hSTMT);
	}
	return Result;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

short CSQL::BindParameter(HSTMT hSTMT, short iParamNumber, void *pUserData,
	short iType, SQLLEN *iItemCount)
{
	/*
		iUserData will be passed back out of pToken when calling ParamData.
	*/


	//SQLBindParameter(hStmt, 1, SQL_PARAM_INPUT, SQL_C_TCHAR, SQL_VARCHAR, 256, 0, pData1, _tcslen(pData1), &i1);

	SQLRETURN Result = SQLBindParameter(hSTMT, iParamNumber, SQL_PARAM_INPUT,
		SQL_C_BINARY, iType,
		*iItemCount, 0, pUserData, 0, iItemCount);

	*iItemCount = SQL_LEN_DATA_AT_EXEC(0);

	if (!SQL_SUCCEEDED(Result))
	{
		this->ThrowErrorIfSet(hSTMT);
	}
	return Result;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

SQLLEN CSQL::RowCount(HSTMT hSTMT)
{
	SQLLEN iRowCount = 0;
	if (SQLRowCount(hSTMT, &iRowCount) != SQL_SUCCESS)
	{
		iRowCount = -1;
	}
	return iRowCount;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

HSTMT CSQL::Prepare(const char *sSQL)
{
	HSTMT hSTMT = NULL;

	if (!this->bConnected)
	{
		return NULL;
	}

	SQLRETURN Result = 0;

	Result = SQLAllocHandle(SQL_HANDLE_STMT, this->hcSQLConnection, &hSTMT);
	if (SQL_SUCCEEDED(Result))
	{
		SQLSetStmtAttr(hSTMT, SQL_ATTR_RETRIEVE_DATA, (SQLPOINTER)SQL_RD_OFF, NULL);
		SQLSetStmtAttr(hSTMT, SQL_ATTR_QUERY_TIMEOUT, (SQLPOINTER)this->icTimeout, NULL);

		Result = SQLPrepare(hSTMT, (unsigned char *)sSQL, (int)strlen(sSQL));
		if (SQL_SUCCEEDED(Result))
		{
			return hSTMT;
		}
		else this->ThrowErrorIfSet(hSTMT);
	}
	else this->ThrowErrorIfSet(hSTMT);

	return NULL;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

short CSQL::PutData(HSTMT hSTMT, void *pData, SQLLEN iLength)
{
	SQLRETURN Result = SQLPutData(hSTMT, (SQLPOINTER)pData, iLength);
	if (!SQL_SUCCEEDED(Result))
	{
		this->ThrowErrorIfSet(hSTMT);
	}
	return Result;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

short CSQL::FreeHandle(HSTMT hSTMT)
{
	SQLRETURN Result = SQLFreeHandle(SQL_HANDLE_STMT, hSTMT);
	if (!SQL_SUCCEEDED(Result))
	{
		this->ThrowErrorIfSet(hSTMT);
	}
	return Result;
}


//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

short CSQL::ParamData(HSTMT hSTMT, void **pToken)
{
	SQLRETURN Result = SQLParamData(hSTMT, (SQLPOINTER *)pToken);
	if (!SQL_SUCCEEDED(Result))
	{
		this->ThrowErrorIfSet(hSTMT);
	}
	return Result;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Execute(const char *sDatabase, const char *sSQL)
{
	char sOldDatabase[1024];
	this->GetFocus(sOldDatabase, sizeof(sOldDatabase));

	if (_strcmpi(sOldDatabase, sDatabase) != 0)
	{
		this->Focus(sDatabase);
	}

	bool bResult = this->Execute(sSQL);

	if (_strcmpi(sOldDatabase, sDatabase) != 0)
	{
		this->Focus(sOldDatabase);
	}

	return bResult;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Execute(const char *sSQL)
{
	if (!bConnected)
	{
		return false;
	}

	SQLRETURN Result = 0;
	bool bResult = false;

	HSTMT hSTMT = NULL;

	Result = SQLAllocHandle(SQL_HANDLE_STMT, this->hcSQLConnection, &hSTMT);
	if (SQL_SUCCEEDED(Result))
	{
		SQLSetStmtAttr(hSTMT, SQL_ATTR_RETRIEVE_DATA, (SQLPOINTER)SQL_RD_OFF, NULL);
		SQLSetStmtAttr(hSTMT, SQL_ATTR_QUERY_TIMEOUT, (SQLPOINTER)this->icTimeout, NULL);

		Result = SQLExecDirect(hSTMT, (unsigned char *)sSQL, (int)strlen(sSQL));
		if (SQL_SUCCEEDED(Result) || Result == SQL_NO_DATA)
		{
			if (Result == SQL_SUCCESS || Result == SQL_NO_DATA || Result == SQL_SUCCESS_WITH_INFO)
			{
				bResult = true;
			}
			else this->ThrowErrorIfSet(hSTMT);
		}
		else this->ThrowErrorIfSet(hSTMT);
	}
	else this->ThrowErrorIfSet(hSTMT);

	SQLFreeStmt(hSTMT, SQL_CLOSE);
	SQLFreeHandle(SQL_HANDLE_STMT, hSTMT);

	hSTMT = NULL;

	return bResult;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Execute(const char *sSQL, CRecordSet *lpMyRS)
{
	return this->Execute(sSQL, lpMyRS, this->icCursorType);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Execute(const char *sSQL, CRecordSet *lpMyRS, int iSessionCursorType)
{
	if (!this->bConnected)
	{
		return false;
	}

	lpMyRS->SetErrorHandler(this->pErrorHandler);

	SQLRETURN Result = 0;

	Result = SQLAllocHandle(SQL_HANDLE_STMT, this->hcSQLConnection, &lpMyRS->hSTMT);
	if (SQL_SUCCEEDED(Result))
	{
		SQLSetStmtAttr(lpMyRS->hSTMT, SQL_ATTR_QUERY_TIMEOUT, (SQLPOINTER)this->icTimeout, NULL);
		SQLSetStmtAttr(lpMyRS->hSTMT, SQL_ATTR_CONCURRENCY, (SQLPOINTER)SQL_CONCUR_READ_ONLY, NULL);
		SQLSetStmtAttr(lpMyRS->hSTMT, SQL_ATTR_CURSOR_TYPE, (SQLPOINTER)iSessionCursorType, NULL);

		Result = SQLPrepare(lpMyRS->hSTMT, (unsigned char *)sSQL, (int)strlen(sSQL));
		if (SQL_SUCCEEDED(Result))
		{
			Result = SQLExecute(lpMyRS->hSTMT);
			if (Result == SQL_SUCCESS || Result == SQL_SUCCESS_WITH_INFO || Result == SQL_NO_DATA)
			{
				if (SQLRowCount(lpMyRS->hSTMT, &lpMyRS->RowCount) != SQL_SUCCESS)
				{
					lpMyRS->RowCount = 0;
				}

				if (SQLNumResultCols(lpMyRS->hSTMT, (SQLSMALLINT *)&lpMyRS->ColumnCount) != SQL_SUCCESS)
				{
					lpMyRS->ColumnCount = 0;
				}

				return true;

			}
			else this->ThrowErrorIfSet(lpMyRS->hSTMT);
		}
		else this->ThrowErrorIfSet(lpMyRS->hSTMT);
	}
	else this->ThrowErrorIfSet(lpMyRS->hSTMT);

	return false;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Execute(const char *sSQL, CBoundRecordSet *lpMyRS)
{
	return this->Execute(sSQL, lpMyRS, true, this->icCursorType);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Execute(const char *sSQL, CBoundRecordSet *lpMyRS, int iSessionCursorType)
{
	return this->Execute(sSQL, lpMyRS, true, iSessionCursorType);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Execute(const char *sSQL, CBoundRecordSet *lpMyRS, bool AutoBindAll)
{
	return this->Execute(sSQL, lpMyRS, AutoBindAll, this->icCursorType);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Execute(const char *sSQL, CBoundRecordSet *lpMyRS, bool AutoBindAll, int iSessionCursorType)
{
	if (!this->bConnected)
	{
		return false;
	}

	lpMyRS->SetErrorHandler(this->pErrorHandler);

	SQLRETURN Result = 0;

	Result = SQLAllocHandle(SQL_HANDLE_STMT, this->hcSQLConnection, &lpMyRS->hSTMT);
	if (SQL_SUCCEEDED(Result))
	{
		SQLSetStmtAttr(lpMyRS->hSTMT, SQL_ATTR_QUERY_TIMEOUT, (SQLPOINTER)this->icTimeout, NULL);
		SQLSetStmtAttr(lpMyRS->hSTMT, SQL_ATTR_CONCURRENCY, (SQLPOINTER)SQL_CONCUR_READ_ONLY, NULL);
		SQLSetStmtAttr(lpMyRS->hSTMT, SQL_ATTR_CURSOR_TYPE, (SQLPOINTER)iSessionCursorType, NULL);

		Result = SQLPrepare(lpMyRS->hSTMT, (unsigned char *)sSQL, (int)strlen(sSQL));
		if (SQL_SUCCEEDED(Result))
		{
			Result = SQLExecute(lpMyRS->hSTMT);
			if (Result == SQL_SUCCESS || Result == SQL_SUCCESS_WITH_INFO || Result == SQL_NO_DATA)
			{
				if (SQLRowCount(lpMyRS->hSTMT, &lpMyRS->RowCount) != SQL_SUCCESS)
				{
					lpMyRS->RowCount = 0;
				}

				if (SQLNumResultCols(lpMyRS->hSTMT, (SQLSMALLINT *)&lpMyRS->Columns.Count) != SQL_SUCCESS)
				{
					lpMyRS->Columns.Count = 0;
				}

				if (lpMyRS->Columns.Count > 0) //Fill in the column info for the RecordSet.
				{
#ifdef _USE_GLOBAL_MEMPOOL
					lpMyRS->Columns.Column = (LPRECORDSET_COLUMNINFO)
						pMem->Allocate(lpMyRS->Columns.Count, sizeof(RECORDSET_COLUMNINFO));
#else
					lpMyRS->Columns.Column = (LPRECORDSET_COLUMNINFO)
						calloc(lpMyRS->Columns.Count, sizeof(RECORDSET_COLUMNINFO));
#endif

					for (int iCol = 1; iCol < (lpMyRS->Columns.Count + 1); iCol++)
					{
						LPRECORDSET_COLUMNINFO Pointer = &lpMyRS->Columns.Column[iCol - 1];
						memset(&Pointer, sizeof(RECORDSET_COLUMNINFO), 0);

						int iColumnNameLength = 0;
						int iIsNullable = 0;
						int iDataType = 0;

						lpMyRS->GetColumnInfo(iCol, Pointer->Name, sizeof(Pointer->Name),
							&iColumnNameLength, &iDataType, &Pointer->MaxSize,
							&Pointer->DecimalPlaces, &iIsNullable);

						Pointer->DataType = (SQLTypes::SQLType)iDataType;
						Pointer->Name[iColumnNameLength] = '\0';
						Pointer->IsNullable = (iIsNullable > 0);
						Pointer->Ordinal = iCol;
						Pointer->IsBound = false;
						if (AutoBindAll)
						{
							lpMyRS->Bind(iCol - 1, lpMyRS->GetDefaultConversion(Pointer->DataType));
						}
					}
				}

				return true;

			}
			else this->ThrowErrorIfSet(lpMyRS->hSTMT);
		}
		else this->ThrowErrorIfSet(lpMyRS->hSTMT);
	}
	else this->ThrowErrorIfSet(lpMyRS->hSTMT);

	return false;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Execute(const char *sDatabase, const char *sSQL, CBoundRecordSet *lpMyRS)
{
	return this->Execute(sDatabase, sSQL, lpMyRS, this->icCursorType);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Execute(const char *sDatabase, const char *sSQL, CBoundRecordSet *lpMyRS, int iSessionCursorType)
{
	return this->Execute(sDatabase, sSQL, lpMyRS, true, iSessionCursorType);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Execute(const char *sDatabase, const char *sSQL, CBoundRecordSet *lpMyRS, bool AutoBindAll)
{
	return this->Execute(sDatabase, sSQL, lpMyRS, AutoBindAll, this->icCursorType);

}
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Execute(const char *sDatabase, const char *sSQL, CBoundRecordSet *lpMyRS, bool AutoBindAll, int iSessionCursorType)
{
	char sOldDatabase[1024];
	this->GetFocus(sOldDatabase, sizeof(sOldDatabase));

	if (_strcmpi(sOldDatabase, sDatabase) != 0)
	{
		this->Focus(sDatabase);
	}

	bool bResult = Execute(sSQL, lpMyRS, AutoBindAll, iSessionCursorType);

	if (_strcmpi(sOldDatabase, sDatabase) != 0)
	{
		this->Focus(sOldDatabase);
	}

	return bResult;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::GetFocus(char *sOutDatabaseName, int iOutDatabaseNameSz)
{
	if (bConnected)
	{
		SQLLEN iResult = 0;

		CRecordSet rsTemp;

		if (this->Execute("SELECT db_name()", &rsTemp))
		{
			if (rsTemp.Fetch())
			{
				SQLLEN iLength = 0;
				if (rsTemp.sColumnEx(1, sOutDatabaseName, iOutDatabaseNameSz, &iResult))
				{
					return true;
				}
			}

			rsTemp.Close();
		}
	}
	return false;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::Focus(const char *sDatabaseName)
{
	if (!this->bConnected)
		return false;

	char sSQL[1024];
	sprintf_s(sSQL, sizeof(sSQL), "USE [%s]", sDatabaseName);

	return this->Execute(sSQL);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

int CSQL::TransactionDepth(void)
{
	if (!this->bConnected)
		return -1;

	int iResult = 0;

	CRecordSet rsTemp;

	if (this->Execute("SELECT @@TRANCOUNT", &rsTemp))
	{
		if (rsTemp.Fetch())
		{
			if (!rsTemp.lColumnEx(1, (long *)&iResult))
			{
				iResult = -4;
			}
		}
		else iResult = -3;

		rsTemp.Close();
	}
	else iResult = -2;

	return iResult;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::BeginTransaction(void)
{
	return this->BeginTransaction(NULL);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::BeginTransaction(const char *sName)
{
	if (!this->bConnected)
		return false;

	const char *sTrans = "BEGIN TRANSACTION";
	char sSQL[1024];
	if (sName)
	{
		sprintf_s(sSQL, sizeof(sSQL), "%s %s", sTrans, sName);
	}
	else {
		strcpy_s(sSQL, sizeof(sSQL), sTrans);
	}

	return this->Execute(sSQL);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::CommitTransaction(void)
{
	return this->CommitTransaction(NULL);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::CommitTransaction(const char *sName)
{
	if (!this->bConnected)
		return false;

	const char *sTrans = "COMMIT TRANSACTION";
	char sSQL[1024];
	if (sName)
	{
		sprintf_s(sSQL, sizeof(sSQL), "%s %s", sTrans, sName);
	}
	else {
		strcpy_s(sSQL, sizeof(sSQL), sTrans);
	}

	return this->Execute(sSQL);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::RollbackTransaction(void)
{
	return this->RollbackTransaction(NULL);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::RollbackTransaction(const char *sName)
{
	if (!this->bConnected)
		return false;

	const char *sTrans = "ROLLBACK TRANSACTION";
	char sSQL[1024];
	if (sName)
	{
		sprintf_s(sSQL, sizeof(sSQL), "%s %s", sTrans, sName);
	}
	else {
		strcpy_s(sSQL, sizeof(sSQL), sTrans);
	}

	return this->Execute(sSQL);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void CSQL::SetErrorHandler(ErrorHandler pHandler)
{
	this->pErrorHandler = pHandler;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

int CSQL::CursorType(void)
{
	return this->icCursorType;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

int CSQL::CursorType(int iCursorType)
{
	int iOldCursorType = this->icCursorType;
	this->icCursorType = iCursorType;
	return iOldCursorType;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::ThrowErrors(void)
{
	return this->bcThrowErrors;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

bool CSQL::ThrowErrors(bool bThrowErrors)
{
	bool bOldValue = this->bcThrowErrors;
	this->bcThrowErrors = bThrowErrors;
	return bOldValue;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

HDBC CSQL::ConnectionHandle(void)
{
	return this->hcSQLConnection;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

HDBC CSQL::EnvironmentHandle(void)
{
	return this->hcSQLEnvironment;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

unsigned int CSQL::TimeOut(void)
{
	return this->icTimeout;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

unsigned int CSQL::TimeOut(unsigned int iCommandTimeout)
{
	unsigned int iOldTimeOut = this->icTimeout;
	this->icTimeout = iCommandTimeout;
	return iOldTimeOut;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#endif
