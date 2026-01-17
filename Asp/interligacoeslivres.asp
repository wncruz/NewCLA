<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: InterligacoesLives.asp
'	- Responsável		: Vital
'	- Descrição			: Busca de interligações livres
%>
<!--#include file="../inc/data.asp"-->
<%if request("rec_id") <> "" then%>
	<html>
	<head>
	<title>CLA - Controle Local de Acesso</title>
	<link rel=stylesheet type="text/css" href="../css/cla.css">
	<script language='javascript' src="../javascript/cla.js"></script>
	<script language='javascript' src="../javascript/claMsg.js"></script>
	</head>
	<body leftmargin="0" topmargin="0" class=TA>
	<table valign="top" width="100%" align=center>
	<tr>
	<td>
	<center>
	<%
	dblRecId = request("rec_id")
	if Trim(dblRecId) <> "" then
		Set objRSRec = db.execute("CLA_sp_view_recurso " & dblRecId)
		if Not objRSRec.Eof And Not objRSRec.Bof then
			dblDstId		= objRSRec("Dst_Id")
			dblEscId		= objRSRec("Esc_Id")
		End if
	End if	
	Set objRSRec = Nothing

	Set cor = db.execute("CLA_sp_sel_InterligacaoLivre " & dblEscId & "," &  dblDstId & ","  & request("qtd"))
	if not cor.eof then
	%>
	<table border="0" cellspacing="1" cellpadding=1 width="100%">
	<Form name=Form1 method=Post>
	<input type=hidden name=hdnRecId value="<%=request("rec_id")%>">
	<input type=hidden name=hdnSistema value="<%=Request("hdnRede")%>">
	<tr>
	<th colspan=2><p align=center>Posições Livres</p></th>
	</tr>
	<tr>
		<%if Request("hdnRede") = "3" then %>
			<td>PADE/PAC</td>
		<%Else%>
			<td>Coordenada</td>
		<%End if%>	
		<td>
			<input type=text class=text size=20 maxlength=20 name=txtCoordenada>
			<input type=button class=button name=btnProcurar value=Procurar onClick="ProcurarInter()" accesskey="P" onmouseover="showtip(this,event,'Procurar (Alt+P)');">
		</td>
	</tr>
</table>
<span id=spnPosicoes>
	<table border="0" cellspacing="1" cellpadding=1 width="100%">
	<tr>
		<th>&nbsp;Origem</th>
		<th>&nbsp;Destino</th>
	</tr>

	<%
		do while not cor.eof
			if strClass = "clsSilver" then strClass = "clsSilver2" else strClass = "clsSilver" end if
	%>
			<tr class="<%=strClass%>">
				<td style="cursor:hand" align=center>&nbsp;<a onClick="javascript:copyPaste('<%=cor("Int_CorOrigem")%>,<%=cor("Int_CorDestino")%>');"><%=cor("Int_CorOrigem")%></a></td>
				<td style="cursor:hand" align=center>&nbsp;<a onClick="javascript:copyPaste('<%=cor("Int_CorOrigem")%>,<%=cor("Int_CorDestino")%>');"><%=cor("Int_CorDestino")%></a></td>
			</tr>
	<%
			cor.movenext
		loop
	%>

	</tr>
	</table>
</span>
	<%Else%>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" valign=center>
		<tr>
			<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Não há posições livres para este recurso.</font></td>
		</tr>		
		</table>
	<%End If%>
	</td>
	</tr>
	<tr>
		<td align=center height=30px>
			<input type=button class=button name=btnSair value=Sair onClick="Sair()" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">&nbsp;
		</td>
	</tr>
	</table>
<script language="JavaScript">

function Sair()
{
	window.returnValue = objAryFac
	window.close()
}

var objAryFac = window.dialogArguments

function copyPaste()
{
	var indice = 0;
	var blnCheio = true
 	if (VerificarFacExistente(arguments[0]))
 	{
		<%if Request("hdnRede") = "3" then %>
 			alert("PADE/PAC já foi selecionada!")
		<%Else%>
 			alert("Coordenada já foi selecionada!")
		<%End if%>
		return
 	}
			
	for (var intIndex=0;intIndex<objAryFac.length;intIndex++)
	{
		if (objAryFac[intIndex] =="")
		{
			objAryFac[intIndex] = arguments[0];
			alert("Facilidade preenchida com sucesso!")
			blnCheio = false
			return
		}
	}
	if (blnCheio) {alert("Já existem "+objAryFac.length+" interligações selecionadas.");return}
}				   

function VerificarFacExistente(strChave)
{
	var blnRet = false 
	for (var intIndex=0;intIndex<objAryFac.length;intIndex++)
	{
		if (strChave == objAryFac[intIndex])
		{
			blnRet = true
		}
	}
	return blnRet
}
</script>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub window_onunload
	window.returnValue = objAryFac
End Sub
-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
function ProcurarInter(){

	with (document.forms[0]){
		<%if Request("hdnRede") = "3" then %>
			if (!ValidarCampos(txtCoordenada,"PADE/PAC")) return
		<%Else%>
			if (!ValidarCampos(txtCoordenada,"Coordenada")) return
		<%End if%>	
		hdnAcao.value = "ConsultarCoordenadasLivresAlocacao"
		target = "IFrmProcesso"
		action = "ProcessoConsFac.asp"
		submit()
	}
}
//-->
</SCRIPT>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnNomeCons value="ConsInter">
</Form>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "0" 
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
</body>
</html>
<%end if%>
