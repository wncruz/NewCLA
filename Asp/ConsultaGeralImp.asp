<!--#include file="../inc/data.asp"-->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Cla - Controle Local de Acesso - Impressão de Consulta Geral</TITLE>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<SCRIPT LANGUAGE=javascript>
<!--
self.focus() 
function Imprimir()
{
	window.print()
}
-->
</SCRIPT>

</HEAD>
<BODY>
<br>

<%

	strCampos = Request.Form("hdnCampos") 
	strWhere  = Request.Form("hdnWhere") 
	strOrderBy= Request.Form("hdnOrderBy") 
	strGroupBy= Request.Form("hdnGroupBy")
	strConsOrigem = Request.Form("hdnConsOrigem")
	
	if Trim(strCampos) = "" then
		Response.Write "<script language=javascript>alert('Problemas para consultar o registro.')</script>"
		Response.End 
	End if
	
	Vetor_Campos(1)="adWChar,3000,adParamInput," & strCampos
	Vetor_Campos(2)="adWChar,3000,adParamInput," & strWhere
	Vetor_Campos(3)="adWChar,3000,adParamInput," & strGroupBy
	Vetor_Campos(4)="adWChar,3000,adParamInput," & strOrderBy 
	Vetor_Campos(5)="adWChar,1,adParamInput," & strConsOrigem
	
	Call APENDA_PARAM("CLA_sp_sel_consultaGeral",5,Vetor_Campos)
	Set objRS = ObjCmd.Execute'pega dbaction
		
	if Err.number <> 0 then
		Response.Write "<script language=javascript>alert('Problemas para consultar o registro.\nVerifique os dados selecionados/digitados!')</script>"
		Response.End 
	End if		
	'spnConsulta
	Dim objDic
	Set objDic = Server.CreateObject("Scripting.Dictionary")  

	if Not objRS.Eof and Not objRS.bof then
		set objRSDic = db.execute("select * from dicionario")
		While not objRSDic.Eof
			if not objDic.Exists(Trim(Ucase(objRSDic("Dic_Campo"))))  then
				Call objDic.Add (Trim(Ucase(objRSDic("Dic_Campo"))),Trim(objRSDic("Dic_Comentario")))
			End if
			'Next
			objRSDic.MoveNext
		Wend	
			
	Else
		Response.Write "<script language=javascript>alert('Registro não encontrado')</script>"
		Response.End 
	End if

	strHtmlRet = ""
	strHtmlRet = strHtmlRet & "<table class='TableLine' border='1' cellspacing='0' cellpadding='1' bordercolor='black' width='100%' cellspacing='0' cellpadding='3' align='center'>"
	strHtmlRet = strHtmlRet & "<tr class=clsSilver><td colspan=" & objRS.Fields.Count & "><font color=black>&nbsp;•&nbsp;CLA - Consulta Geral - " & Formatar_Data(date) & " " & Time() & "</font></td></tr>"
	strHtmlRet = strHtmlRet & "<tr class=clsSilver>"
	For intIndex=0 to objRS.Fields.Count - 1
		if objDic.Exists(Trim(Ucase(objRS.Fields.Item(intIndex).Name))) then 
			strHtmlRet = strHtmlRet & "<td nowrap><font color=black>&nbsp;" & objDic(Trim(Ucase(objRS.Fields.Item(intIndex).Name))) & "</font></td>"
		End if	
	Next	
	strHtmlRet = strHtmlRet & "</tr>"
		

	While Not objRS.Eof
		strHtmlRet = strHtmlRet & "<tr>"
		dblSolId = objRS("Sol_id")
		For intIndex=0 to objRS.Fields.Count - 1
			if objRS.Fields.Item(intIndex).Name <> "Sol_id" then
				strHtmlRet = strHtmlRet & "<td nowrap><font color=black>"
				strHtmlRet = strHtmlRet & objRS(objRS.Fields.Item(intIndex).Name)
				strHtmlRet = strHtmlRet & "</font></td>"
			End if	
		Next
		strHtmlRet = strHtmlRet & "</tr>"
		objRS.MoveNext
	Wend
	strHtmlRet = strHtmlRet & "</table>"
	
	Response.Write strHtmlRet
%>
<p align=center>
	<input type=button class=button name=btnImprimir value=Imprimir onClick="Imprimir()">&nbsp;
	<input type="button" class="button" name="btnFechar" value="Fechar" onClick="javascript:window.close()">
</p>
</BODY>
</HTML>
