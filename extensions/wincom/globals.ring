/*
	WinCOM - Global Constants
*/

WINCOM_VERSION = "1.0.0"

# Modes used by com_dispatch()
COM_GET    = 1
COM_INVOKE = 2
COM_SET    = 3
COM_SETREF = 4

# VARIANT types (used by com_variant() and returned by com_lasttype())
VT_EMPTY    = 0
VT_NULL     = 1
VT_I2       = 2
VT_I4       = 3
VT_R4       = 4
VT_R8       = 5
VT_CY       = 6
VT_DATE     = 7
VT_BSTR     = 8
VT_DISPATCH = 9
VT_ERROR    = 10
VT_BOOL     = 11
VT_VARIANT  = 12
VT_UNKNOWN  = 13
VT_DECIMAL  = 14
VT_I1       = 16
VT_UI1      = 17
VT_UI2      = 18
VT_UI4      = 19
VT_I8       = 20
VT_UI8      = 21
VT_INT      = 22
VT_UINT     = 23
VT_ARRAY    = 8192

# ADO Constants (ADODB)

# CursorTypeEnum
adOpenForwardOnly = 0
adOpenKeyset      = 1
adOpenDynamic     = 2
adOpenStatic      = 3

# LockTypeEnum
adLockReadOnly        = 1
adLockPessimistic     = 2
adLockOptimistic      = 3
adLockBatchOptimistic = 4

# CursorLocationEnum
adUseServer = 2
adUseClient = 3

# CommandTypeEnum and ExecuteOptionEnum
adCmdText          = 1
adCmdTable         = 2
adCmdStoredProc    = 4
adCmdUnknown       = 8
adCmdTableDirect   = 512
adExecuteNoRecords = 128

# DataTypeEnum
adSmallInt       = 2
adInteger        = 3
adSingle         = 4
adDouble         = 5
adCurrency       = 6
adDate           = 7
adBSTR           = 8
adBoolean        = 11
adVariant        = 12
adDecimal        = 14
adTinyInt        = 16
adUnsignedTinyInt = 17
adBigInt         = 20
adGUID           = 72
adBinary         = 128
adChar           = 129
adWChar          = 130
adNumeric        = 131
adDBDate         = 133
adDBTime         = 134
adDBTimeStamp    = 135
adVarChar        = 200
adLongVarChar    = 201
adVarWChar       = 202
adLongVarWChar   = 203
adVarBinary      = 204
adLongVarBinary  = 205

# ParameterDirectionEnum
adParamInput       = 1
adParamOutput      = 2
adParamInputOutput = 3
adParamReturnValue = 4

# ObjectStateEnum
adStateClosed    = 0
adStateOpen      = 1
adStateConnecting = 2
adStateExecuting = 4
adStateFetching  = 8

# SchemaEnum (Connection.OpenSchema)
adSchemaColumns     = 4
adSchemaIndexes     = 12
adSchemaTables      = 20
adSchemaPrimaryKeys = 28

# AffectEnum, SearchDirectionEnum and GetRowsOptionEnum
adAffectCurrent  = 1
adAffectGroup    = 2
adAffectAll      = 3
adSearchForward  = 1
adSearchBackward = -1
adGetRowsRest    = -1

# PersistFormatEnum and SaveOptionsEnum (Recordset.Save, Stream.SaveToFile)
adPersistADTG       = 0
adPersistXML        = 1
adSaveCreateNotExist = 1
adSaveCreateOverWrite = 2

# StreamTypeEnum
adTypeBinary = 1
adTypeText   = 2

# EventStatusEnum (ADO events, the adStatus argument)
adStatusOK             = 1
adStatusErrorsOccurred = 2
adStatusCantDeny       = 3
adStatusCancel         = 4
adStatusUnwantedEvent  = 5

# ADOX Constants (ADOX.Catalog)

# KeyTypeEnum
adKeyPrimary = 1
adKeyForeign = 2
adKeyUnique  = 3

# ColumnAttributesEnum
adColFixed    = 1
adColNullable = 2
