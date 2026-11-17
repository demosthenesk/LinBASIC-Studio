/* odbc_types.h - the ODBC declarations odbc.c needs (unixODBC, 64-bit), for a computer without
 * the headers of unixODBC (unixodbc-dev). Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#ifndef LB_ODBC_TYPES_H
#define LB_ODBC_TYPES_H
typedef short SQLSMALLINT; typedef unsigned short SQLUSMALLINT; typedef int SQLINTEGER; typedef unsigned int SQLUINTEGER;
typedef long SQLLEN; typedef unsigned long SQLULEN; typedef short SQLRETURN; typedef void *SQLPOINTER;
typedef void *SQLHANDLE; typedef SQLHANDLE SQLHENV, SQLHDBC, SQLHSTMT; typedef void *SQLHWND;
typedef unsigned short SQLWCHAR;
#define SQL_API
#define SQL_SUCCEEDED(rc) (((rc) & (~1)) == 0)
#define SQL_SUCCESS 0
#define SQL_SUCCESS_WITH_INFO 1
#define SQL_NO_DATA 100
#define SQL_NULL_HANDLE 0
#define SQL_HANDLE_ENV 1
#define SQL_HANDLE_DBC 2
#define SQL_HANDLE_STMT 3
#define SQL_ATTR_ODBC_VERSION 200
#define SQL_OV_ODBC3 3UL
#define SQL_NTS (-3)
#define SQL_DRIVER_NOPROMPT 0
#define SQL_ATTR_AUTOCOMMIT 102
#define SQL_AUTOCOMMIT_OFF 0UL
#define SQL_AUTOCOMMIT_ON 1UL
#define SQL_COMMIT 0
#define SQL_ROLLBACK 1
#define SQL_FETCH_NEXT 1
#define SQL_FETCH_FIRST 2
#define SQL_PARAM_INPUT 1
#define SQL_NULL_DATA (-1)
#define SQL_NO_TOTAL (-4)
#define SQL_TRUE 1
#define SQL_ATTR_READONLY 0
#define SQL_C_CHAR 1
#define SQL_C_WCHAR (-8)
#define SQL_C_SBIGINT (-25)
#define SQL_C_DOUBLE 8
#define SQL_C_BINARY (-2)
#define SQL_C_TYPE_TIMESTAMP 93
#define SQL_CHAR 1
#define SQL_NUMERIC 2
#define SQL_DECIMAL 3
#define SQL_INTEGER 4
#define SQL_SMALLINT 5
#define SQL_FLOAT 6
#define SQL_REAL 7
#define SQL_DOUBLE 8
#define SQL_DATE 9
#define SQL_TIME 10
#define SQL_TIMESTAMP 11
#define SQL_VARCHAR 12
#define SQL_TYPE_DATE 91
#define SQL_TYPE_TIME 92
#define SQL_TYPE_TIMESTAMP 93
#define SQL_LONGVARCHAR (-1)
#define SQL_BINARY (-2)
#define SQL_VARBINARY (-3)
#define SQL_LONGVARBINARY (-4)
#define SQL_BIGINT (-5)
#define SQL_TINYINT (-6)
#define SQL_BIT (-7)
#define SQL_WVARCHAR (-9)
#define SQL_DESC_AUTO_UNIQUE_VALUE 11
#define SQL_DESC_TABLE_NAME 15
#define SQL_DESC_UNSIGNED 8
#define SQL_DESC_UPDATABLE 10
#define SQL_DESC_BASE_COLUMN_NAME 22
#define SQL_DESC_BASE_TABLE_NAME 23
#define SQL_DESC_OCTET_LENGTH 1013
typedef struct { SQLSMALLINT year; SQLUSMALLINT month, day, hour, minute, second; SQLUINTEGER fraction; } SQL_TIMESTAMP_STRUCT;
#endif
