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
Public Const BadResponse As Long = 10008

Public Const BadRequest As Long = 10400
Public Const Unauthorized As Long = 10401
Public Const Forbidden As Long = 10403
Public Const NotFound As Long = 10404
Public Const MethodNotAllowed As Long = 10405
Public Const NotAcceptable As Long = 10406
Public Const PreconditionFailed As Long = 10412
Public Const InternalServerError As Long = 10500
