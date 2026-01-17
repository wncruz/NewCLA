<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: filePost.ASP
'	- Responsável		: Vital
'	- Descrição			: Gravar arquivo enviado para upload no servidor

PathFile = Server.MapPath(Request.Cookies("PATH_FILE"))

Set Upload = Server.CreateObject("Persits.Upload")

Upload.OverwriteFiles = true
Upload.Save

' Check whether a file was selected
Set File = Upload.Files("txtFile")
If Not File Is Nothing Then
   ' Obtain file name
   Filename = file.Filename
   File.SaveAs PathFile
   Response.Write "<script language=javascript>alert('Arquivo atualizado com sucesso!');</script>"
Else ' file not selected
   Response.Write "<script language=javascript>alert('Arquivo não enviado.')</script>"
End If
%>