<%
' Dados de entrada
Dim identificador1, identificador2, timestamp, dados, md5Hash

identificador1 = "user123"
identificador2 = "abc456"
timestamp = Now ' ou use: timestamp = FormatDateTime(Now, vbGeneralDate)

' Formata a string com separador "|"
dados = identificador1 & "|" & identificador2 & "|" & timestamp

' Cria o hash MD5 usando CAPICOM
Set Hasher = CreateObject("CAPICOM.HashedData")
Hasher.Algorithm = 3 ' 3 = MD5
Hasher.Hash datos

md5Hash = Hasher.Value

' Exibe resultado
Response.Write "MD5: " & md5Hash

' Limpeza
Set Hasher = Nothing
%>
