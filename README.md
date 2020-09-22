<div align="center">

## AutoCompleter \- Class Module


</div>

### Description

This code allows you to have an autocomplete function on any text boxes by creating an instance of the class module below and setting a text control on a form to is CompleteTextbox property. Ideal for those situations when you have multiple autocompletes. (Visual Basic 6 Only - Can easily be modified for 5.0 users)
 
### More Info
 


Dim m_objAutoCompleteUser as clsAutoComplete

Set m_objAutoCompleteUser = New clsAutoComplete

With m_objAutoCompleteUser

.SearchList = m_strUserList

Set .CompleteTextbox = txtUser

.Delimeter = ","

End With

Create a new class module.

Paste all the code below into it.

Rename the module to clsAutoComplete.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[dmbbob](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dmbbob.md)
**Level**          |Unknown
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dmbbob-autocompleter-class-module__1-4073/archive/master.zip)





### Source Code

```
Option Explicit
Private WithEvents m_txtComplete As TextBox
Private m_strDelimeter As String
Private m_strList As String
Private Sub m_txtComplete_KeyUp(KeyCode As Integer, Shift As Integer)
 Dim i As Integer
 Dim strSearchText As String
 Dim intDelimented As Integer
 Dim intLength As Integer
 Dim varArray As Variant
 With m_txtComplete
  If KeyCode <> vbKeyBack And KeyCode > 48 Then
   If InStr(1, m_strList, .Text, vbTextCompare) <> 0 Then
    varArray = Split(m_strList, m_strDelimeter)
    For i = 0 To UBound(varArray)
     strSearchText = Trim(varArray(i))
     If InStr(1, strSearchText, .Text, vbTextCompare) And
      (Left$(.Text, 1) = Left$(strSearchText, 1)) And
      .Text <> "" Then
      .SelText = ""
      .SelLength = 0
      intLength = Len(.Text)
      .Text = .Text & Right$(strSearchText, Len(strSearchText) - Len(.Text))
      .SelStart = intLength
      .SelLength = Len(.Text)
      Exit Sub
     End If
    Next i
   End If
  End If
 End With
End Sub
Public Property Get CompleteTextbox() As TextBox
 Set CompleteTextbox = m_txtComplete
End Property
Public Property Set CompleteTextbox(ByRef txt As TextBox)
 Set m_txtComplete = txt
End Property
Public Property Get SearchList() As String
 SearchList = m_strList
End Property
Public Property Let SearchList(ByVal str As String)
 m_strList = str
End Property
Public Property Get Delimeter() As String
 Delimeter = m_strDelimeter
End Property
Public Property Let Delimeter(ByVal str As String)
 m_strDelimeter = str
End Property
```

