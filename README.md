# VBATests
'Testing for some functionalities in VBA Azure SQL

Function ReadImg(ByVal tbl As String, ByVal tblid As Integer, ByVal FORM As Object, ByVal Ck As Control)

Dim stm As ADODB.Stream, nameexist As String, TempName As String
Dim xtb As Variant
TempName = Environ$("HOMEDRIVE") & "\Temp" & Format(Now, "hh-mm-ss") & ".jpg"

'If Not Dir(TempName) = "" Then
'TempName = "C:\Users\angeandr\Desktop\" & Format(Now, "hh-mm-ss") & ".jpg"
'End If

Call abrirCON
With rsPubs
.ActiveConnection = cnpubs
.Open "Select * from imgs where Tbl='" & tbl & "' and TbID=" & tblid & " and deleted is Null"
'.Open "Select * from imgs where ID=5"

'ActiveCell.CopyFromRecordset rsPubs

If Not .EOF Then
'xtb = .GetRows
'For i = 0 To UBound(xtb, 1)
'For Y = 0 To UBound(xtb, 2)
'MsgBox xtb(i, Y)
'Next
'Next
'MsgBox .Fields("Cnt").value
Else
FORM.Controls(Ck.Name) = False
GoTo finish
End If

Set stm = New ADODB.Stream
    stm.Type = adTypeBinary
    stm.Open

.MoveFirst

    stm.Write .Fields("img").value  ' write bytes to stream
    
    'stm.Write xtb(3, 0)
    stm.Position = 0
    stm.SaveToFile (TempName)
    stm.Close
    
 'Insertar imagen en sheet "logo"
With ThisWorkbook.Sheets(Sheet8.Name)
.Visible = True
.Activate

Dim sh As Shape
For Each sh In .Shapes
If InStr(sh.Name, "Picture") Then sh.Delete
Next

.Pictures.Insert(TempName).Select
PicName = Selection.ShapeRange.Name
.Shapes(PicName).Top = 80
.Shapes(PicName).Left = 0
End With
     
    'FORM.Controls(ctrl.Name).Picture = LoadPicture(TempName) ' load bytes into Image control on form
    FORM.Controls(Ck.Name) = True

End With

Call cerrarCON

Set stm = Nothing
            
If Not Dir(TempName) = "" Then
    Kill (TempName)
End If


finish:
If Not FORM.Controls(Ck.Name) Then
With ThisWorkbook.Sheets(Sheet8.Name)
.Visible = True
.Activate
Dim shx As Shape
For Each shx In .Shapes
If InStr(shx.Name, "Picture") Then shx.Delete
Next
End With
End If

End Function
