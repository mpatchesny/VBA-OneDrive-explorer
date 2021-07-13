# VBA OneDrive explorer

A very simple VBA OneDrive explorer. It lets you navigate through files and folders in OneDrive. Uses Microsoft Graph api.

## Important

In order this to work you have to set Files.Read permissions for Microsoft Graph, please see [How to set Microsoft Graph permissions](## How to set Microsoft Graph permissions).

## Installation

Import all files into macro enabled workbook. Add VBA references:
Microsoft Scripting Runtime
Microsoft VBScript Regular Expressions 5.5
Microsoft WinHTTP Services, version 5.1

It uses VBA-JSON library, you can download latest version from here [VBA-JSON](https://github.com/VBA-tools/VBA-JSON/releases).

## Usage

Add module to your macro and paste following code:

```vb
Public sub Start()

    On Error Goto ErrHandler

    Dim token As String
    token = "" ' paste your token here

    Dim explorer As OneDriveFileExplorer
    Set explorer = New OneDriveFileExplorer
    explorer.Display entryPointPath:="https://graph.microsoft.com/v1.0/me/drive/root/", token:=token, userformTitle:="Select file", allowMultiselect:=True, selectMode:=ESelectModeAll

    If Not explorer.IsCancelled Then
        Dim SelectedItems As Collection
        Set SelectedItems = explorer.SelectedItems
    End If

    Exit Sub
    
ErrHandler:
    MsgBox "Error!" & vbCrLf & vbCrLf & "Error description: " & err.Description & vbCrLf & "Error source: " & err.Source, vbExclamation, "Error!"

End Sub
```

OneDriveFileExplorer.Display arguments:

- entryPointPath - path to root OneDrive folder

- token - Microsoft Graph token (see [How to obtain Microsoft Graph token](## How to obtain Microsoft Graph token))

- userformTitle - title with which the form will be displayed

- allowMultiselect - whether selecting multiple items is allowed or not

- selectMode - if ESelectModeAll then you can select files and folder, if ESelectModeFolder then only folders, if ESelectModeFile then only files

SelectedItems collection is collection of IDriveItem objects. From IDriveItem object you can obtain id and path of drive item (that is file or folder), eg.:

```vb
Private Sub DebugPrintSelectedItems(ByRef col As Collection)

    If Not col Is Nothing Then
        Dim item As IDriveItem
        For Each item In col
            Debug.Print item.Id, item.Path
        Next item
    End If
    
End Sub
```

## How to set Microsoft Graph permissions

Go to [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) and log in. Click three dots menu button next to your account name and click Select permissions -> select Files and click Consent.

## How to obtain Microsoft Graph token

Go to [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) and log in. Click Access Token tab and copy token.
