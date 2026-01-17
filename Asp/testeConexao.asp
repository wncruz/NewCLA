<%
On Error Resume Next

Dim ConSGA
Set ConSGA = Server.CreateObject("ADODB.Connection")

Response.write "OOO"

'ConSGA.Open "Provider=OraOLEDB.Oracle;Password=ATUALIZA;User ID=SGAV_ATUALIZA;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=seu_host)(PORT=1523))(CONNECT_DATA=(SERVICE_NAME=sgagsid2)));Persist Security Info=True"
'ConSGA.Open "Provider=OraOLEDB.Oracle;Data Source=SGAGSID2;User ID=SGAV_ATUALIZA;Password=ATUALIZA"
ConSGA.Open             "Provider=OraOLEDB.Oracle;Password=ATUALIZA;User ID=SGAV_ATUALIZA;Data Source=SGAGSID2.WORLD;Persist Security Info=True"

If Err.Number <> 0 Then
    Response.Write("<b>Erro na coneXAo:</b> " & Err.Description)
Else
    Response.Write("<b>Conexão bem-sucedida!</b>")
End If

ConSGA.Close
Set ConSGA = Nothing
%>
