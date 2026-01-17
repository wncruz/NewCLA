<%@ CodePage=65001 %>
<%
	Response.ContentType = "text/html; charset=utf-8"
	Response.Charset = "UTF-8"
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<head>
    
</head>
<%
'•EMBRATEL - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: SolicitacaoAprov_Main.asp
'	- Descrição			: Solicitação vinda do sistema Aprovisionador

Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
if Request.Form("hdnXmlReturn") <> "" then
	objXmlDados.loadXml(Request.Form("hdnXmlReturn"))
Else
	objXmlDados.loadXml("<xDados/>")
End if

%>
<script language='javascript' src="../javascript/help.js"></script>
<SCRIPT LANGUAGE="JavaScript">
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")
function checa(f) {
	return true;
}

function AtId(solid,acfid,vall,AprovisiId){
	with (document.forms[0])
		{
				hdn678.value = vall
				hdnSolId.value = solid
				hdnAcfId.value = acfid
				hdnAprovisiId.value = AprovisiId
		}
	}

function ConsultarPedidosPend()
{
	with (document.forms[0])
	{
		if (cboAcao.value == "" && txt_oe_numero.value == "")
		{
			alert("Selecione uma ação.")
			return;
		}
		
		if (cboAcao.value == "REENV" && txt_oe_numero.value == "")
		{
			alert("Informe o número da Order Entry.");
			txt_oe_numero.focus();
			return;
		}
	
		hdnAcao.value = "SEL"
		target = self.name
		action = "SolicitacaoReenvioAprov_main.asp?Consulta=1"
		submit()
	}
}

function DetalharItem(dblSolId)
{
	with (document.forms[0])
	{
		PopularXml()
		hdnSolId.value = dblSolId
		DetalharFac()
	}
}
</script>

<form action="SolicitacaoReenvioAprov_main.asp" name="Form1" method="post" onsubmit="return checa(this)">
<input type=hidden name=hdnPedId>
<input type=hidden name=hdnSolId>
<input type=hidden name=hdnAprovisiId>
<input type=hidden name=hdnXmlReturn>
<input type=hidden name=hdnOEOrigem value="Aprov">
<input type=hidden name=hdnAcaoApg>
<input type=hidden name=cboAcao value="REENV">
<tr>
<td >

<table border=0 cellspacing="1" cellpadding="0" width="760" >
<tr >
	<th colspan=2 ><table width="760"><tr><th><center>Solicitação de Reenvio - Aprovisionador</th>
			<th width="26">
<!--
  Good Início				
			<div id="ajuda1" style="position:absolute; left:-1000px; top:-1000px; width:234; height:223; z-index:1;visibility: hidden; border: 0px none;">
		    <object name="button" classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="https://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="234" height="223">
            	<param name="movie" value="../imagens/Help/help.swf?arquivo=../imagens/Help/SolicitacaoAprov_main.txt">
    			<param name="quality" value="high">
    		<embed src="../imagens/Help/help.swf?arquivo=../imagens/Help/SolicitacaoAprov_main.txt" quality="high" pluginspage="https://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="234" height="223"></embed>
 			</object>
			</div>
						
			<div id="espaco_left" title="Ajuda">
			<a href="javascript:;" onMouseOver="showAjudaImg('duvida1','ajuda1',10,20 );" onClick="hideAjuda('ajuda1');">
			<img src="../imagens/question_b.gif" border="0" id="duvida1" style="cursor:pointer"/></a>
			</div>			
  Good fim
-->  			
			</th></tr></table>
		</th>
</tr>
<tr class=clsSilver>
 <td colspan="2"><center>ATENÇÃO: <br>Ao solicitar um reenvio de uma OE do Aprovisionador SGA Voz para o CLA, o sistema irá gerar automaticamente um processo de cancelamento do ID Lógico (678 ou 677) anterior.
 <br>
 </td>
</tr>

<tr class=clsSilver>
	<td>Aprovisionador</td>
	<td>
	    <select name="Cbo_OrisolAprov">
		  <%
		  set rsCboOrisolAprov = db.execute("select Orisol_ID,OriSol_Alias from CLA_OrigemSolicitacao where Orisol_ID in( 7 ) order by orisol_alias ")
		  do while not rsCboOrisolAprov.eof
		  %>
			<option value="<%=rsCboOrisolAprov("Orisol_ID")%>"
		  <%
		    'if strGrupo <> "" then
			  'if trim(strGrupo) = trim(gr("GCli_ID")) then
			    'response.write "selected"
			  'end if
			'end if
		  %>
			><%=ucase(rsCboOrisolAprov("OriSol_Alias"))%></option>
		  <%
			rsCboOrisolAprov.movenext
			loop
		  %>
		</select>
	</td>
</tr>
<tr class=clsSilver>
	<td>OE</td>
	<td>
        <input id="txt_oe_numero" type="text" title="Número" maxlength="7" size="8" 
            class=text onKeyUp="ValidarTipo(this,0)" 
            name="txt_oe_numero" value='<%=request("txt_oe_numero")%>'>&nbsp;/
        <input id="txt_oe_ano" type="text" title="Ano" maxlength="4" size="5" 
            class=text onKeyUp="ValidarTipo(this,0)" name="txt_oe_ano" 
            value='<%=request("txt_oe_ano")%>'>&nbsp;item&nbsp;
        <input id="txt_oe_item" type="text" title="Item" maxlength="3" size="4" 
            class=text onKeyUp="ValidarTipo(this,0)" name="txt_oe_item" 
            value='<%=request("txt_oe_item")%>'>
	</td>
</tr>
<%
Dim IntNumOE_id
IntNumOE_id = null
IntNumOE_id = Request.form("txtNumOE")
%>

<tr>
	<td colspan=2 align=center height=35px>
		<input type="button" name="btconsulta" value="Consultar" class="button" onClick="ConsultarPedidosPend()">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px">
        </td>
</tr>
<tr>
	<td colspan=2 align=center height=35px>
	<iframe	id			= "IFrmProcesso"
			name        = "IFrmProcesso"
			width       = "0"
			height      = "0"
			frameborder = "0"
			scrolling   = "no"
			align       = "left">
	</iFrame>
	</td>
</tr>
</table>
<span id=spnLinks></span>
<%
Dim strClass
Dim dblEstId
Dim intIndex
Dim strSql
Dim intCount
Dim strSel
Dim strXls
Dim strLink
Dim strHtml

if Request.ServerVariables("CONTENT_LENGTH") = 0  then
	dblUsuIdFac = dblUsuId
End If

Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
if Request.Form("hdnXMLReturn") <> "" then
	objXmlDados.loadXml(Request.Form("hdnXMLReturn"))
	set objNodeAux = objXmlDados.getElementsByTagName("cboUsuario")
	if objNodeAux.length > 0 then dblUsuIdFac = objNodeAux(0).childNodes(0).text
End if

if dblUsuIdFac = "" then dblUsuIdFac = "null" End if

'inseri teste para verificar se deve ou não realizar a consulta
if Request.QueryString ("Consulta") = "1" or Request.QueryString ("btn") <> "" then
	strResponse = RetornaTabela(dblUsuId) 
	Response.Write strResponse 
	%>
	<!--#include file="../inc/ControlesPaginacao.asp"-->
	<%
end if
%>
<input type=hidden name=hdnXls value="<%=strXls%>">
<input type="hidden" name="hdnPaginaOrig"	value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnOrigem>
<input type=hidden name=hdnNomeCons value="PedidosPend">
<input type=hidden name="hdn678" value =<%=Request.Form("hdn678")%>>
<input type=hidden name=hdnAcfId value=<%=Request.Form("hdnAcfId")%>>
</form>

</body>
</html>
<SCRIPT LANGUAGE=javascript>
spnLinks.innerHTML = '<%=TratarAspasJS(strLink)%>'

function EditarFac(dblSolId,AprovisiId,strAcao,StrOrigem)
{
	with (document.forms[0])
	{
		PopularXml()
		hdnSolId.value = dblSolId
		hdnAprovisiId.value = AprovisiId
		hdnOEOrigem.value = "Aprov"
		hdnAcao.value = strAcao					
		hdnOrigem.value = StrOrigem			
		action = "SolicitacaoReenvioAprov.asp"		
		submit()		
	}
}
</SCRIPT>

<form name="form_confirma" method="post" action="ProcessoReenvioAprov.asp" target="IFrmProcesso">
	<input type="hidden" name="hdnAprovisiID">
	<input type="hidden" name="hdnAcao">
        <input type="hidden" name="hdnOrisol_ID">
	<input type=hidden name="proc_SN" value="N">         
</form>

<script language="VBscript">
function Confirmar(Aprovisi_ID,Acao,Acl_IDAcessoLogico,OE,orisol_id)

  if form_confirma.proc_SN.value = "S" then
  	MsgBox "Aguarde, o cancelamento do ID lógico '" & Acl_IDAcessoLogico & "' (OE " & OE & ") está em processamento.", 48, "CLA - Reenvio de Solicitação"
  else
		returnvalue=MsgBox ("Confirma o cancelamento do ID lógico " & Acl_IDAcessoLogico & " e reenvio da OE "&OE&" para o CLA? Atenção: Esta ação não poderá ser desfeita.",547,"Confirmação definitiva de cancelamento de ID Lógico e reenvio de OE.")               
		If returnvalue=6 Then 
			form_confirma.proc_SN.value = "S"                                                                                                                                 
			form_confirma.hdnAprovisiID.value = Aprovisi_ID
			form_confirma.hdnAcao.value = Acao
		        form_confirma.hdnOrisol_ID.value = orisol_id
			form_confirma.submit()    
			     
		End If
	End if
	
end function
</script>

<%
Set objRSPag = Nothing
DesconectarCla()

function RetornaTabela(Usuario)
   	Vetor_Campos(1)="adInteger,10,adParamInput," 
	Vetor_Campos(2)="adInteger,10,adParamInput,"
	Vetor_Campos(3)="adWChar,30,adParamInput,null"		
	Vetor_Campos(4)="adWChar,5,adParamInput,"	& request("cboAcao")
	Vetor_Campos(5)="adInteger,10,adParamInput," & Usuario
	Vetor_Campos(6)="adWChar,7,adParamInput," & request("txt_Oe_numero")
	Vetor_Campos(7)="adWChar,4,adParamInput," & request("txt_Oe_Ano")
	Vetor_Campos(8)="adInteger,4,adParamInput," & request("txt_Oe_Item")
	Vetor_Campos(9)="adWChar,60,adParamInput," & request("txt_Cliente")
	Vetor_Campos(10)="adWChar,30,adParamInput," & request("txt_Designacao")
	Vetor_Campos(11)="adInteger,10,adParamInput," & request("Cbo_OrisolAprov")

    strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_Aprovisionador ",11,Vetor_Campos)

	call paginarRS(1,strSql)
	intCount=1

	if not objRSPag.Eof and not objRSPag.Bof then
		'Link Xls/Impressão
		strLink =	"<table border=0 width=760><tr><td colspan=2 align=right>" & _
						"<a href='javascript:AbrirXls()' onmouseover=""showtip(this,event,'Consulta em formato Excel...')""><img src='../imagens/excel.gif' border=0></a>&nbsp;" & _
						"<a href='javascript:TelaImpressao(800,600,""Pedidos Pendentes - " & date() & " " & Time() & " "")' onmouseover=""showtip(this,event,'Tela de Impressão...')""><img src='../imagens/impressora.gif' border=0></a></td></tr>" & _
						"</table>"

		strHtml = "<table border=0 cellspacing=1 width=760>" & _
						"<tr>" & _
							"<th align=center>+</th>" & _
							"<th width=120>&nbsp;Numero/Ano OE</th>" & _
							"<th>&nbsp;Item OE</th>" & _
							"<th>&nbsp;Cliente</th>" & _
							"<th>&nbsp;Ação</th>" & _
							"<th nowrap>&nbsp;Serviço</th>" & _
							"<th>&nbsp;ID Lógico</th>" & _
							"<th>&nbsp;Solicitação</th>" & _
							"<th>&nbsp;Designação <br>&nbsp;do Serviço</th>" & _
						"</tr>"
		strXls = strHtml

		For intIndex = 1 to objRSPag.PageSize

			if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

			if isnull(objRSPag("sol_id")) Then strsolicitacao = "0" Else strsolicitacao = objRSPag("sol_id") End if

			If trim(objRSPag("Acao")) = "ATV" Then StrAcao = "ATV" 	End if 

			If trim(objRSPag("Acao")) = "ALT" Then StrAcao = "ALT" 	End if 
			
			If trim(objRSPag("Acao")) = "CAN" Then StrAcao = "CAN" End if 
			
			If trim(objRSPag("Acao")) = "DES" Then StrAcao = "DES" 	End if
			
			If trim(request("cboAcao")) = "REENV" Then StrOrigem = "REENV" 	End if
			
			If trim(objRSPag("Acao")) = "CAN" or trim(objRSPag("Acao")) = "DES" then	
				
				Set objRS = db.Execute("SELECT acf_id FROM CLA_VIEW_SOLICITACAOMIN WHERE sol_id = '" & strsolicitacao & "'")
				if Not objRS.Eof and Not objRS.Bof then  acf_id = objRS("acf_id") end if
				
			End IF
			strHtml = strHtml & "<tr class='" & strClass & "'>" '& _
			
			strHtml = strHtml & "<td>&nbsp;&nbsp;&nbsp;</td>"
			
			strHtml = strHtml & "<td ><a href='javascript:Confirmar(" & objRSPag("Aprovisi_id") & ",""" & StrAcao & """"  & ",""" & objRSPag("Acl_IDAcessoLogico") & """,""" & right("00000" & Trim(objRSPag("OE_Numero")),5) & "/" & objRSPag("OE_Ano") & ""","  & objRSPag("orisol_id") & " )'>" & "OE:" & right("00000" & Trim(objRSPag("OE_Numero")),5) & "/" & objRSPag("OE_Ano") & "</a></td>"
			
			if Trim(objRSPag("OriSol_Descricao")) = "SGAV" then
			  strOriSol_Descricao = "SGA Voz"
			elseif Trim(objRSPag("OriSol_Descricao")) = "SGAP" then
			  strOriSol_Descricao = "SGA Plus"
			else
			  strOriSol_Descricao = Trim(objRSPag("OriSol_Descricao"))
			end if
																
			strHtml = strHtml & "<td >" & objRSPag("OE_Item") & "</td>" & _
								"<td >" & objRSPag("Cli_Nome") & "</td>" & _
								"<td >" & objRSPag("Acao") & "</td>" & _
								"<td >" & objRSPag("Ser_Desc") & "</td>" & _
								"<td >" & objRSPag("Acl_IDAcessoLogico") & "</td>" & _
								"<td >" & objRSPag("Sol_ID") & "</td>" & _
								"<td >" & Trim(objRSPag("Acl_designacaoServico")) & "</td>" & "</tr>"

			strXls = strXls & "<tr class='" & strClass & "'>" & _
								"<td></td>"& _
								"<td >" & "OE:" & right("00000" & objRSPag("OE_Numero"),5) & "/" & objRSPag("OE_Ano") & "</td>" & _
								"<td >" & objRSPag("OE_Item") & "</td>" & _
								"<td >" & objRSPag("Cli_Nome") & "</td>" & _
								"<td >" & objRSPag("Acao") & "</td>" & _
								"<td >" & objRSPag("Acl_NContratoServico") & "</td>" & _
								"<td >" & objRSPag("Acl_IDAcessoLogico") & "</td>" & _
								"<td >" & objRSPag("Sol_ID") & "</td>" & _
								"<td >" & Trim(objRSPag("Acl_designacaoServico")) & "</td>" & _
							"</tr>"

			intCount = intCount+1
			objRSPag.MoveNext
			if objRSPag.EOF then Exit For
		Next
		strHtml = strHtml & "</table>"
		strXls = strXls & "</table>"
		RetornaTabela = strHtml
	Else
		strHtml ="<table width= 760 border= 0 cellspacing= 0 cellpadding= 0 valign=top>"
		strHtml = strHtml + "<tr>"
		strHtml = strHtml + "<td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td>"
		strHtml = strHtml + "</tr>"
		strHtml = strHtml + "</table>"
		RetornaTabela = strHtml
	End if

End function
%>
