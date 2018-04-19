#ifndef _ENTRY_CPP
#define _ENTRY_CPP
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include <Windows.H>
#include <Stdio.H>
#include <Stdlib.H>

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include "../CSQLEncapsulation.H"

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void TestStandardFetch(SQLCONNECTIONSTRING *conString);
void TestAutoBinding(SQLCONNECTIONSTRING *conString);

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

int main(void)
{
	SQLCONNECTIONSTRING conString;
	memset(&conString, 0, sizeof(conString));
	conString.bUseTrustedConnection = true;
	conString.sDriver = "{SQL Server}";
	conString.sDatabase = "msdb";
	conString.sServer = ".";
	conString.sApplicationName = "SQL Test App";

	TestStandardFetch(&conString);

	TestAutoBinding(&conString);

	
	system("pause");

	return 0;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void TestAutoBinding(SQLCONNECTIONSTRING *conString)
{
	CSQL sql;

	sql.ThrowErrors(true);

	if (!sql.Connect(conString))
	{
		return;
	}

	CBoundRecordSet rsObjects;

	sql.Execute("SELECT TOP 100 * FROM sys.objects where type = 'U'", &rsObjects, true);

	char sSQL[1024];

	while (rsObjects.Fetch())
	{
		CBoundRecordSet rsColumns;

		sprintf_s(sSQL, sizeof(sSQL), "SELECT TOP 100 * FROM sys.columns WHERE object_id = %d",
			rsObjects.Values("object_id")->ToIntegerS());

		sql.Execute(sSQL, &rsColumns, true);

		int objectIdIndex = rsObjects.GetColumnIndex("object_id");
		int objectNameIndex = rsObjects.GetColumnIndex("Name");

		printf("%d->%s\n", (int)rsObjects.Columns.Column[objectIdIndex].Data.Buffer,
			(char*)rsObjects.Columns.Column[objectNameIndex].Data.Buffer);

		while (rsColumns.Fetch())
		{
			int columnNameIndex = rsColumns.GetColumnIndex("Name");

			printf("\t%s\n", (char*)rsColumns.Columns.Column[columnNameIndex].Data.Buffer);
		}

		rsColumns.Close();
	}

	rsObjects.Close();

	sql.Disconnect();
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void TestStandardFetch(SQLCONNECTIONSTRING *conString)
{
	CSQL sql;

	sql.ThrowErrors(true);

	if (!sql.Connect(conString))
	{
		return;
	}

	CBoundRecordSet rsObjects;

	sql.Execute("SELECT TOP 100 * FROM sys.objects where type = 'U'", &rsObjects, true);

	char sSQL[1024];

	while (rsObjects.Fetch())
	{
		CBoundRecordSet rsColumns;

		sprintf_s(sSQL, sizeof(sSQL), "SELECT TOP 100 * FROM sys.columns WHERE object_id = %d",
			rsObjects.Values("object_id")->ToIntegerS());

		sql.Execute(sSQL, &rsColumns, true);

		CSQLValue *objectId = rsObjects.Values("object_id");
		CSQLValue *objectName = rsObjects.Values("Name");

		printf("%d->%s\n", objectId->ToIntegerS(), objectName->ToString());

		while (rsColumns.Fetch())
		{
			CSQLValue *columnName = rsColumns.Values("Name");

			printf("\t%s\n", columnName->ToString());

			delete columnName;
		}

		delete objectId;
		delete objectName;

		rsColumns.Close();
	}

	rsObjects.Close();

	sql.Disconnect();
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#endif
