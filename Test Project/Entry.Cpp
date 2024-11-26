#ifndef _ENTRY_CPP
#define _ENTRY_CPP
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include <Windows.H>
#include <Stdio.H>
#include <Stdlib.H>

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include "../CSQLEncapsulation.H"

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void TestStandardFetch(SQLCONNECTIONSTRING* connectionString);
void TestAutoBinding(SQLCONNECTIONSTRING* connectionString);

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

int main()
{
	SQLCONNECTIONSTRING connectionString;
	memset(&connectionString, 0, sizeof(connectionString));
	connectionString.UseTrustedConnection = true;
	connectionString.Driver = "{SQL Server}";
	connectionString.Database = "master";
	connectionString.Server = ".";
	//connectionString.Port = 1433;
	connectionString.ApplicationName = "SQL Test App";

	TestStandardFetch(&connectionString);

	TestAutoBinding(&connectionString);

	system("pause");

	return 0;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

void TestAutoBinding(SQLCONNECTIONSTRING* connectionString)
{
	CSQL sql;

	sql.ThrowErrors(true);

	if (!sql.Connect(connectionString))
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
			rsObjects.Values("object_id")->ToSignedInt());

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

void TestStandardFetch(SQLCONNECTIONSTRING* connectionString)
{
	CSQL sql;

	sql.ThrowErrors(true);

	if (!sql.Connect(connectionString))
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
			rsObjects.Values("object_id")->ToSignedInt());

		sql.Execute(sSQL, &rsColumns, true);

		CSQLValue* objectId = rsObjects.Values("object_id");
		CSQLValue* objectName = rsObjects.Values("Name");

		printf("%d->%s\n", objectId->ToSignedInt(), objectName->ToString());

		while (rsColumns.Fetch())
		{
			CSQLValue* columnName = rsColumns.Values("Name");

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
