///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Copyright © NetworkDLS 2002, All rights reserved
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF 
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO 
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A 
// PARTICULAR PURPOSE.
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#ifndef _CSQLVALUE_H
#define _CSQLVALUE_H
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include <SQL.H>
#include <SQLExt.H>

typedef int(*ErrorHandler)(void *pCaller, const char *sSource, const char *sErrorMsg, const int iErrorNumber);

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

namespace SQLTypes {
	enum SQLType {
		BigInt = 65531,
		Bit = 65529,
		Decimal = 3,
		Float = 6,
		Int = 4,
		Money = 3,
		Numeric = 2,
		Real = 7,
		SmallInt = 5,
		SmallMoney = 3,
		TinyInt = 65530,

		Binary = 65534, //Not (yet) supported.
		Image = 65532,
		VarBinary = 65533,

		NText = 65526,
		Text = 65535,

		DateTime = 93,
		SmallDateTime = 93,
		TimeStamp = 65534,

		Char = 1,
		NChar = 65528,
		NVarChar = 65527,
		Variant = 65386,
		UniqueIdentifier = 65525,
		VarChar = 12
	};
}
namespace CTypes {
	enum CType {
		Char = SQL_C_CHAR,
		SShort = SQL_C_SSHORT,
		UShort = SQL_C_USHORT,
		SLong = SQL_C_SLONG,
		ULong = SQL_C_ULONG,
		Float = SQL_C_FLOAT,
		Double = SQL_C_DOUBLE,
		Bit = SQL_C_BIT,
		STinyInt = SQL_C_STINYINT,
		UTinyInt = SQL_C_UTINYINT,
		SBigInt = SQL_C_SBIGINT,
		UBigInt = SQL_C_UBIGINT,
		Binary = SQL_C_BINARY,
		BookMark = SQL_C_BOOKMARK,
		VarBookMark = SQL_C_VARBOOKMARK,
		Date = SQL_C_TYPE_DATE,
		Time = SQL_C_TYPE_TIME,
		TimeStamp = SQL_C_TYPE_TIMESTAMP,
		Numeric = SQL_C_NUMERIC,
		UniqueIdentifier = SQL_C_GUID
	};
}

typedef struct _TAG_BINDINGDATA {
	LPVOID Buffer;
	SQLLEN Size;
	bool IsNull;
} RECORDSET_SQLVALUE, *LPRECORDSET_SQLVALUE;

typedef struct _TAG_COLUMNINFO {
	RECORDSET_SQLVALUE Data;

	char Name[1024];
	SQLTypes::SQLType DataType;
	SQLULEN MaxSize;
	int DecimalPlaces;
	int Ordinal;
	bool IsBound;
	bool IsNullable;
} RECORDSET_COLUMNINFO, *LPRECORDSET_COLUMNINFO;

typedef struct _TAG_COLUMN {
	RECORDSET_COLUMNINFO *Column;
	int Count;
} RECORDSET_COLUMN, *LPRECORDSET_COLUMN;

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

class CSQLValue {
private:
	LPRECORDSET_COLUMNINFO lpCValue;

	bool bcTrimCharData;
	bool bcReplaceSingleQuotes;
	bool bcThrowErrors;
	bool bcIsValid;

	SQLLEN RTrim(char *sData, SQLLEN iDataSz);
	void ReplaceSingleQuotes(char *sData, SQLLEN iDataSz);

public:
	void *pPublicData;

	void Initialize(LPRECORDSET_COLUMNINFO lpColumnValue,
		bool bTrimCharData, bool bReplaceSingleQuotes, bool bThrowErrors);

	bool ThrowErrors(void);
	bool ThrowErrors(bool bThrowErrors);

	bool TrimCharData(bool bTrimCharData);
	bool TrimCharData(void);

	bool ReplaceSingleQuotes(bool bReplaceSingleQuotes);
	bool ReplaceSingleQuotes(void);

	bool IsValid(void);
	bool IsNull(void);
	SQLLEN Size(void);
	char *ToString(void);
	SQLLEN ToString(char *sOut, int iMaxSz);
	signed int ToIntegerS(void);
	unsigned int ToIntegerU(void);
	double ToDouble(void);
	float ToFloat(void);
	unsigned __int64 ToI64U(void);
	signed __int64 ToI64S(void);
	bool ToBoolean(void);
	unsigned short ToShortU(void);
	signed short ToShortS(void);
};

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#endif
