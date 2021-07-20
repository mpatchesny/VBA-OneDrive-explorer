VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RequestSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Models")
Option Explicit

Private Type TFields
    Tenant As String
    ClientId As String
    ResponseType As String
    RedirectUri As String
    ResponseMode As String
    Scope As String
    State As String
    GrantType As String
    ClientSecret As String
End Type
Private this As TFields

Public Property Get Tenant() As String
    Tenant = this.Tenant
End Property
Public Property Let Tenant(ByVal value As String)
    this.Tenant = value
End Property

Public Property Get ClientId() As String
    ClientId = this.ClientId
End Property
Public Property Let ClientId(ByVal value As String)
    this.ClientId = value
End Property

Public Property Get ResponseType() As String
    ResponseType = this.ResponseType
End Property
Public Property Let ResponseType(ByVal value As String)
    this.Tenant = value
End Property

Public Property Get RedirectUri() As String
    RedirectUri = this.RedirectUri
End Property
Public Property Let RedirectUri(ByVal value As String)
    this.RedirectUri = value
End Property

Public Property Get ResponseMode() As String
    ResponseMode = this.ResponseMode
End Property
Public Property Let ResponseMode(ByVal value As String)
    this.ResponseMode = value
End Property

Public Property Get Scope() As String
    Scope = this.Scope
End Property
Public Property Let Scope(ByVal value As String)
    this.Scope = value
End Property

Public Property Get State() As String
    State = this.State
End Property
Public Property Let State(ByVal value As String)
    this.State = value
End Property

Public Property Get GrantType() As String
    GrantType = this.GrantType
End Property
Public Property Let GrantType(ByVal value As String)
    this.GrantType = value
End Property

Public Property Get ClientSecret() As String
    ClientSecret = this.ClientSecret
End Property
Public Property Let ClientSecret(ByVal value As String)
    this.ClientSecret = value
End Property

Public Property Get Self() As RequestSettings
    Set Self = Me
End Property


