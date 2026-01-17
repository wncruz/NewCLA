<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Recurso.asp
'	- Descrição			: Cadastra/Altera Recurso
%>
<!--#include file="../inc/data.asp"-->
<%
Dim dblId
Dim dblSisId	
Dim dblEscId	
Dim dblDstId	
Dim dblProId 
Dim dblPlaId
Dim objRSRec
Dim strSel

dblId = request("ID")
if Trim(dblId) = "" then
	dblId = Request.Form("hdnId") 
End if

If request("btnGravar")="Gravar" then
	if dblId="" then
		Vetor_Campos(1)="adInteger,2,adParamInput,"
	else
		Vetor_Campos(1)="adInteger,2,adParamInput,"& dblId
	end if
	Vetor_Campos(2)="adInteger,3,adParamInput,"&request("cboSistema")
	Vetor_Campos(3)="adWChar,4,adParamInput,"&request("txtCidSigla")
	Vetor_Campos(4)="adWChar,3,adParamInput,"&request("txtEscSigla")
	Vetor_Campos(5)="adInteger,3,adParamInput,"&request("cboDistribuidor")
	Vetor_Campos(6)="adInteger,3,adParamInput,"&request("cboProvedor")
	Vetor_Campos(7)="adInteger,3,adParamInput,"&request("cboPlataforma")
	Vetor_Campos(8)="adInteger,2,adParamOutput,0"  

	Call APENDA_PARAM("CLA_sp_ins_recurso",8,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value

End if

If dblId<> "" then
	Set objRSRec = db.execute("CLA_sp_view_recurso " & dblId)
	if Not objRSRec.Eof And Not objRSRec.Bof then
		dblSisId	= TratarAspasHtml(objRSRec("Sis_ID"))
		dblEscId	= TratarAspasHtml(objRSRec("Esc_ID"))
		dblDstId	= TratarAspasHtml(objRSRec("Dst_ID"))
		dblProId	= TratarAspasHtml(objRSRec("Pro_ID"))
		dblPlaId	= TratarAspasHtml(objRSRec("Pla_ID"))
 		strCidSigla = TratarAspasHtml(objRSRec("Cid_Sigla"))
		strEscSigla = TratarAspasHtml(objRSRec("Esc_SIgla"))
	End if
Else
	dblSisId	= TratarAspasHtml(Request.Form("cboSistema"))
	dblPlaId	= TratarAspasHtml(Request.Form("cboPlataforma"))
	dblDstId	= TratarAspasHtml(Request.Form("cboDistribuidor"))
	dblProId	= TratarAspasHtml(Request.Form("cboProvedor"))
	strCidSigla = TratarAspasHtml(Request.Form("txtCidSigla"))
	strEscSigla = TratarAspasHtml(Request.Form("txtEscSigla"))
End if

if dblPlaId = "" then
	dblPlaId = 0
end if 
%>
<!--#include file="../inc/header.asp"-->
<form action="recurso.asp" method="post" onSubmit="return checa(this)" name=Form1>
<input type=hidden name=hdnId value="<%=dblId%>">
<SCRIPT LANGUAGE="JavaScript">
function checa(f) 
{
	if (!ValidarCampos(f.cboSistema,"O Sistema")) return false;
	if (!ValidarCampos(f.cboDistribuidor,"A Distribuição")) return false;
	if (!ValidarCampos(f.cboProvedor,"O Provedor")) return false;
	return true;
}
</script>
<tr><td >
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=2><p align=center>Cadastro de Recurso</p></th>
</tr>
<tr class=clsSilver>
<td width="200"><font class="clsObrig">:: </font>Sistema</td>
<td>
<%set objRS = db.execute("CLA_sp_sel_sistema 0")%>
	<select name="cboSistema" onchange= "RetornaCboPlataforma(cboSistema.value,<% = dblPlaId %>)">
		<option value=""></option>
	<%do while not objRS.eof%>
		<%	strSel = ""
			if Trim(dblSisId) = Trim(objRS("Sis_ID")) Then strSel = " selected " End If 
		%>
		<option value="<%=objRS("Sis_ID")%>" <%=strSel%> ><%=objRS("Sis_Desc")%></option>
	<%
	objRS.movenext
	loop
	%>
	</select>
	<span id=spnPlataforma>
	</span>
</td>
</tr>
<tr class=clsSilver>
<td><font class="clsObrig">:: </font>Estação</td>
<td>
	<input type="text" class="text" name="txtCidSigla" value="<%=strCidSigla%>" maxlength="4" size="5">CNL
	<input type="text" class="text" name="txtEscSigla" value="<%=strEscSigla%>" maxlength="3" size="4">Sigla Estação
</td>
</tr>
<tr class=clsSilver>
<td><font class="clsObrig">:: </font>Distribuidor</td>
<td>
<%set objRS = db.execute("CLA_sp_sel_distribuicao 0")%>
	<select name="cboDistribuidor">
		<option value=""></option>
	<%do while not objRS.eof%>
		<%	strSel = ""
			if Trim(dblDstId) = Trim(objRS("Dst_ID")) Then strSel = " selected " End If 
		%>
		<option value="<%=objRS("Dst_ID")%>" <%=strSel%> ><%=objRS("Dst_Desc")%></option>
	<%
	objRS.movenext
	loop
	%>
	</select>
</td>
</tr>
<tr class=clsSilver>
<td><font class="clsObrig">:: </font>Provedor</td>
<td>
<%set objRS = db.execute("CLA_sp_sel_provedor 0")%>
	<select name="cboProvedor">
		<option value=""></option>
	<%do while not objRS.eof%>
		<%	strSel = ""
			if Trim(dblProId) = Trim(objRS("Pro_ID")) Then strSel = " selected " End If 
		%>
		<option value="<%=objRS("Pro_ID")%>" <%=strSel%> ><%=objRS("Pro_Nome")%></option>
	<%
	objRS.movenext
	loop
	%>
	</select>
</td>
</tr>
</table>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
<td colspan=2 align="center">
	<input type="submit" class="button" name="btnGravar" value="Gravar" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
	<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnId.value = '';LimparForm();setarFocus('cboSistema');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
	<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('recurso_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
	<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
</td>
</tr>
<tr>
</tr>
</table>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
<td>
<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
</td>
</tr>
</table>
</td>
</tr>
</table>
</body>
<SCRIPT LANGUAGE=javascript>
<!--
setarFocus('cboSistema')
//-->

function RetornaCboPlataforma(sisID, PlaID)
{
	if (sisID != 1) {
		spnPlataforma.innerHTML = ""
		return 
	}
	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
	var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
	var strXML
	
	strXML = "<root>"
	strXML = strXML + "<plaid>" +  PlaID + "</plaid>"
	strXML = strXML + "<funcao></funcao>"
	strXML = strXML + "</root>" 
	
	xmlDoc.loadXML(strXML);
	xmlhttp.Open("POST","RetornaPlataforma.asp" , false);
	xmlhttp.Send(xmlDoc.xml);
	
	strXML = xmlhttp.responseText;
	spnPlataforma.innerHTML = strXML
}
RetornaCboPlataforma(document.forms[0].cboSistema.value,<% = dblPlaId %>)
</SCRIPT>

</html>
<%
Set objRSRec = Nothing
DesconectarCla()
%>
