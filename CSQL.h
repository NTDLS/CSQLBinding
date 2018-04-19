///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Copyright © NetworkDLS 2002, All rights reserved
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF 
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO 
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A 
// PARTICULAR PURPOSE.
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#ifndef _SQLCLASS_H
#define _SQLCLASS_H
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include <SQL.H>
#include <SqlExt.H>
#include <ODBCSS.h>

#include "CRecordSet.H"
#include "CBoundRecordSet.H"

typedef int(*ErrorHandler)(void *pCaller, const char *sSource, const char *sErrorMsg, const int iErrorNumber);

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

typedef struct _TAG_SQLCONNECTION {

	char *sApplicationName;
	char *sDriver;
	char *sServer;
	char *sDatabase;
	char *sUID;
	char *sPwd;

	int iPort;

	bool bUseTrustedConnection;
	bool bUseTCPIPConnection;
	bool bUseMARS;

} SQLCONNECTIONSTRING, *LPSQLCONNECTIONSTRING;

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

class CSQL {

private:
	ErrorHandler pErrorHandler;

public:
	void *pPublicData;

	void SetErrorHandler(ErrorHandler pHandler);

	bool Connect(const char *sConnectionString, ErrorHandler pHandler);
	bool Connect(LPSQLCONNECTIONSTRING SQLCon);
	bool Connect(LPSQLCONNECTIONSTRING SQLCon, ErrorHandler pHandler);
	bool Connect(const char *sConnectionString);
	bool Connect(const char *sServer, const char *sDatabase, const char *sUser, const char *sPassword);
	bool Connect(const char *sServer, const char *sDatabase);
	bool Connect(const char *sServer, const char *sDatabase, const char *sUser, const char *sPassword, ErrorHandler pHandler);
	bool Connect(const char *sServer, const char *sDatabase, ErrorHandler pHandler);

	bool GetErrorMessage(int *iOutErr, char *sOutError, const int iErrBufSz, HSTMT hStmt);
	bool GetSQLState(char *sSQLState, HSTMT hStmt);
	bool ThrowErrorIfSet(HSTMT hStmt);
	bool ThrowError(HSTMT hStmt);
	void Disconnect(void);

	bool Execute(const char *sSQL);
	bool Execute(const char *sDatabase, const char *sSQL);
	bool Execute(const char *sSQL, CBoundRecordSet *lpMyRS);
	bool Execute(const char *sSQL, CBoundRecordSet *lpMyRS, bool AutoBindAll);
	bool Execute(const char *sSQL, CRecordSet *lpMyRS);
	bool Execute(const char *sDatabase, const char *sSQL, CBoundRecordSet *lpMyRS);
	bool Execute(const char *sDatabase, const char *sSQL, CBoundRecordSet *lpMyRS, bool AutoBindAll);
	bool Execute(const char *sSQL, CBoundRecordSet *lpMyRS, int iSessionCursorType);
	bool Execute(const char *sSQL, CBoundRecordSet *lpMyRS, bool AutoBindAll, int iSessionCursorType);
	bool Execute(const char *sSQL, CRecordSet *lpMyRS, int iSessionCursorType);
	bool Execute(const char *sDatabase, const char *sSQL, CBoundRecordSet *lpMyRS, int iSessionCursorType);
	bool Execute(const char *sDatabase, const char *sSQL, CBoundRecordSet *lpMyRS, bool AutoBindAll, int iSessionCursorType);

	//Deffered Buffering:
	HSTMT Prepare(const char *sSQL);
	short BindParameter(HSTMT hSTMT, short iParamNumber, void *pUserData, short iType, SQLLEN *iItemCount);
	short Execute(HSTMT hSTMT);
	short ParamData(HSTMT hSTMT, void **pToken);
	short PutData(HSTMT hSTMT, void *pData, SQLLEN iLength);
	short FreeHandle(HSTMT hSTMT);
	SQLLEN RowCount(HSTMT hSTMT);

	bool BuildConnectionString(LPSQLCONNECTIONSTRING SQLCon, char *sOut, int iOutSz);
	bool Focus(const char *sDatabaseName);
	bool GetFocus(char *sOutDatabaseName, int iOutDatabaseNameSz);
	bool IsConnected(void);

	CSQL(void);
	~CSQL(void);

	HDBC ConnectionHandle(void);
	HDBC EnvironmentHandle(void);

	unsigned int TimeOut(unsigned int iCommandTimeout);
	unsigned int TimeOut(void);

	bool ThrowErrors(void);
	bool ThrowErrors(bool bThrowErrors);

	int CursorType(void);
	int CursorType(int iCursorType);

	int TransactionDepth(void);
	bool BeginTransaction(void);
	bool BeginTransaction(const char *sName);
	bool CommitTransaction(void);
	bool CommitTransaction(const char *sName);
	bool RollbackTransaction(void);
	bool RollbackTransaction(const char *sName);

	const char *ConnectionString(void);

private:
	bool bcThrowErrors;
	bool bcUseBulkOperations; //Not yet implemented!
	int icCursorType;	// SQL_CURSOR_FORWARD_ONLY, SQL_CURSOR_DYNAMIC,
						//  SQL_CURSOR_KEYSET_DRIVEN or SQL_CURSOR_STATIC

	char *LastConnectionString;

	HENV hcSQLEnvironment;
	HDBC hcSQLConnection;
	bool bConnected;
	unsigned int icTimeout;
};

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#endif
