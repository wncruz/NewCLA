<%
                Response.ContentType = "text/html; charset=utf-8"
                Response.Charset = "UTF-8"
              %>
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ExcelExport.asp
'	- Responsável		: Vital
'	- Descrição			: Grava arquivos de consulta .xls no servidor

Dim strXls
Dim strNomeCons
Dim strPath
Dim objFS
Dim objXls
Dim strNomeArquivo
Dim intIndex
Dim dblRandon  
 
 
 
Set objFS = Server.CreateObject("Scripting.FileSystemObject")

'Remove arquivos com mais de 3 dias
Set objFld  = objFS.GetFolder(Server.MapPath("../Excel/Consultas")) 
For Each objFile in objFld.Files
	if objFile.dateCreated <= (date()-2) then
		if objFile.name <> "Leiame.txt" then
			Call objFS.DeleteFile(Server.MapPath("../Excel/Consultas/" & objFile.name),true)
		End if	
	End if	
Next

strXls = ""


For intIndex = 1 To Request.Form("hdnXls").Count
  strXls = strXls & Request.Form("hdnXls")(intIndex)
Next
strXls = Replace(strXls,"border=0","border=1")
strNomeCons = Request.Form("hdnNomeCons")

'response.write "<script>alert('"& strXls &"')</script>"

Randomize
dblRandon = Round(1000000000*Rnd(1000))
strNomeArquivo = strNomeCons &  Replace(Date,"/","-") & "+" & Replace(Replace(Time,":","-")," ","-") & "-"& dblRandon & ".xls"
strNomeArquivo = strNomeCons & "_" &  Replace(Date,"/","-") & "_" & Replace(Replace(Time,":","-")," ","-") & "_"& dblRandon & ".xls"
strPath = Server.MapPath("../Excel/Consultas/" & strNomeArquivo)



Call objFS.CreateTextFile(strPath,true)

Set objXls = objFS.OpenTextFile(strPath,2)
Call objXls.writeline(strXls)

Set objXls = Nothing
Set objFS = Nothing

Response.Redirect "../Excel/Consultas/" & strNomeArquivo
%>