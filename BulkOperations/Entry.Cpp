#ifndef _ENTRY_CPP
#define _ENTRY_CPP
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include <Windows.H>
#include <Stdio.H>
#include <Stdlib.H>

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include "../CSQLEncapsulation.H"
#include <sql.h>
#include <sqlext.h>
#include <odbcss.h>

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// Column variables.
// cbCola and cbColb must be defined right before Cola and szColb because
//	they are used as  bulk copy indicator variables.
struct ColaData{
	SQLINTEGER cbCola;
	SQLINTEGER Cola;
} ColaInst;

struct ColbData{
	SQLINTEGER cbColb;
	SQLCHAR   szColb[11];
} ColbInst;

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

int main(void)
{
	SQLCONNECTIONSTRING MySqlCon;
	CSQL MySQL;

	memset(&MySqlCon, 0, sizeof(MySqlCon));

	////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////
	//This is unfinished.
	////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////

	MySqlCon.bUseTrustedConnection = true;
	strcpy_s(MySqlCon.sDriver, sizeof(MySqlCon.sDriver), "{SQL Server}");
	strcpy_s(MySqlCon.sDatabase, sizeof(MySqlCon.sDatabase), "BI");
	strcpy_s(MySqlCon.sServer, sizeof(MySqlCon.sServer), "CLT51Test");
	strcpy_s(MySqlCon.sApplicationName, sizeof(MySqlCon.sServer), "SQL Test App");

	MySQL.bThrowErrors = true;
	MySQL.bUseBulkOperations = true;

	if(!MySQL.Connect(&MySqlCon))
	{
		return 0;
	}

	// Initialize the bulk copy.
	retcode = bcp_init(MySQL.ConnectionHandle(), "BI..WHFinancialClassChanges", NULL, NULL, DB_OUT);

	// Bind the program variables for the bulk copy.
	retcode = bcp_bind(hdbc1, (BYTE *)&ColaInst.cbCola, 9,
               SQL_VARLEN_DATA, NULL, (INT)NULL, SQLINT4, 1);


	MySQL.Disconnect();

	system("pause");

	return 0;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#endif

//http://doc.ddart.net/mssql/sql2000/html/howtosql/ht_odbcht10_9xkj.htm