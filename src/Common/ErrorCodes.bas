Attribute VB_Name = "ErrorCodes"
'@Folder("Common")
Option Explicit

Public Const TypeMismatch As Long = 13
Public Const FileNotFound As Long = 53
Public Const PathAccessError As Long = 75
Public Const ArgumentNotNull As Long = 449
Public Const JsonParseError As Long = 10001
Public Const NotSupported As Long = 10005
Public Const NotImplemented As Long = 10006
Public Const ValidationError As Long = 10007
Public Const CsvParseError As Long = 10008
Public Const AppUpdateError As Long = 10009
Public Const RecordDoesNotExist As Long = 10010
Public Const RecordAlreadyExists As Long = 10011
Public Const ExciseValidatorError As Long = 10012
Public Const ConnectionClosed As Long = 10013
Public Const NoActiveTransaction As Long = 10014
Public Const CreateParamError As Long = 10015
Public Const ExecuteCommandError As Long = 10016
Public Const NoDbUser As Long = 10017
Public Const ConnectionAlreadyOpened As Long = 10018
Public Const UserNotActive As Long = 10019
Public Const StartConditionsNotMet As Long = 10020
Public Const UnexpectedValue As Long = 10021
Public Const EmptyFileBin As Long = 10022
Public Const LoggingLevelNotSet As Long = 10023
Public Const LoggingOutputNotSet As Long = 10024
Public Const CommandNotExecuted As Long = 10025
Public Const FileAlreadyImported As Long = 10026
Public Const TooManyRecords As Long = 10027
