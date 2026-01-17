<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<head>
    
</head>
<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: Solicitacao_Main.asp
'	- Descrição			: Consulta de Pedidos Pendentes

'Response.Write "<br><br>"
'Response.Write "<th colspan=2 ><p align=center><font size=5>Estamos em manutenção previsão 30 Minutos.</font></p></th>"
'Response.end

Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
if Request.Form("hdnXmlReturn") <> "" then
	objXmlDados.loadXml(Request.Form("hdnXmlReturn"))
Else
	objXmlDados.loadXml("<xDados/>")
End if

%>
<SCRIPT LANGUAGE="JavaScript">
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")
function checa(f) {
	return true;
}

function AtId(solid,acfid,vall,solApgId){
	with (document.forms[0])
		{
				hdn678.value = vall
				hdnSolId.value = solid
				hdnAcfId.value = acfid
				hdnSolAPGId.value = solApgId
		}
	}

function ConsultarPedidosPend()
{
	with (document.forms[0])
	{
		hdnAcao.value = "SEL"
		target = self.name
		action = "Solicitacao_main.asp?Consulta=1"

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

<form action="Solicitacao_main.asp" name="Form1" method="post" onsubmit="return checa(this)">
<input type=hidden name=hdnPedId>
<input type=hidden name=hdnSolId>
<input type=hidden name=hdnSolAPGId>
<input type=hidden name=hdnXmlReturn>
<input type=hidden name=hdnOEOrigem value="APG">
<input type=hidden name=hdnAcaoApg>
<tr>
<td >

<table border=0 cellspacing="1" cellpadding="0" width="760" >
<tr >
	<th colspan=2 ><p align=center>Solicitações Aprovisionamento APG</p></th>
</tr>
<tr class=clsSilver>
	<td>
		Processo
	</td>
	<td>
		<select name="cboProcesso">
		  <option value = "" > Selecione um Processo </option>	
		  <option value="ATV" <%if request("cboProcesso") = "ATV" then%>selected<%end if%>>ATIVAÇÃO</option>
		  <option value="DES" <%if request("cboProcesso") = "DES" then%>selected<%end if%>>DESATIVAÇÃO</option>
		</select>
	</td>
</tr>
<tr class=clsSilver>
	<td>
		Acao
	</td>
	<td>
		<select name="cboAcao">
			  <option value="" > Selecione uma Ação </option>	
			  <option value="ATV" <%if request("cboAcao") = "ATV" then%>selected<%end if%>>ATIVAÇÃO</option>
			  <option value="CAN" <%if request("cboAcao") = "CAN" then%>selected<%end if%>>CANCELAMENTO</option>
			  <option value="ALT" <%if request("cboAcao") = "ALT" then%>selected<%end if%>>ALTERACAO</option>
			  <option value="DES" <%if request("cboAcao") = "DES" then%>selected<%end if%>>DESATIVAÇÃO</option>
		</select>
	</td>
</tr>

<tr class=clsSilver>
	<td>
		OE 
	</td>
	<td>
        <input id="txt_oe_numero" type="text" title="Número" maxlength="7" size="7" 
            class=text onKeyUp="ValidarTipo(this,0)" onclick="return Text1_onclick()" 
            name="txt_oe_numero" value='<%=request("txt_oe_numero")%>'>&nbsp;/
        <input id="txt_oe_ano" type="text" title="Ano" maxlength="4" size="5" 
            class=text onKeyUp="ValidarTipo(this,0)" name="txt_oe_ano" 
            value='<%=request("txt_oe_ano")%>'>&nbsp;item&nbsp;
        <input id="txt_oe_item" type="text" title="Item" maxlength="4" size="5" 
            class=text onKeyUp="ValidarTipo(this,0)" name="txt_oe_item" 
            value='<%=request("txt_oe_item")%>'>
	</td>
</tr>
<tr class=clsSilver>
	<td>
		
	    Cliente</td>
	<td>
	    <input id="txt_cliente" type="text" class="text" maxlength="60" 
            name="txt_cliente" size="61" value='<%=request("txt_cliente")%>' /></td>
</tr>
<tr class=clsSilver>
	<td>
		
	    Designação</td>
	<td>
	    <input id="txt_designacao" type="text" class="text" maxlength="11" 
            name="txt_designacao" size="12" onKeyUp="ValidarTipo(this,0)" value='<%=request("txt_designacao")%>'></td>
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


'dblUsuIdFac = Request.Form("cboUsuario")
if Request.ServerVariables("CONTENT_LENGTH") = 0  then
	dblUsuIdFac = dblUsuId
	'response.write "<script>alert('"&dblUsuIdFac&"')</script>"
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
	'response.write "<script>alert('"&dblUsuId&"')</script>"
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
<input type=hidden name=hdnNomeCons value="PedidosPend">
<input type=hidden name="hdn678" value =<%=Request.Form("hdn678")%>>
<input type=hidden name=hdnAcfId value=<%=Request.Form("hdnAcfId")%>>
</form>

</body>
</html>
<SCRIPT LANGUAGE=javascript>
<!--
spnLinks.innerHTML = '<%=TratarAspasJS(strLink)%>'
//setarFocus('cboUsuario')
//-->



function EditarFac(dblSolId,dblSolIdAPG,strAcao)
{
	with (document.forms[0])
	{
		<%if dblUsuIdFac <> "999999999" then%>
			PopularXml()
			hdnSolId.value = dblSolId
			//hdnPedId.value = dblPedId
			hdnSolAPGId.value = dblSolIdAPG
			hdnOEOrigem.value = "APG"
			hdnAcao.value = strAcao
			
			var strNome = "Solicitacao " + dblSolId 
			var objJanela = window.open()
			objJanela.name = strNome
			target = strNome
			//target = "IFrmProcesso_"
			action = "Solicitacao.asp"
			
			submit()
		<%else%>
			hdnSolId.value = dblSolId
			//hdnPedId.value = dblPedId
			hdnSolAPGId.value = dblSolIdAPG
			hdnOEOrigem.value = "APG"
			hdnAcao.value = strAcao

			target = "IFrmProcesso"
			action = "Solicitacao.asp"
			submit()
		<%End if%>
	}
}

function ContinuaAlocacao(dblSolId,dblPedId)
{
	with (document.forms[0])
	{
		PopularXml()
		hdnSolId.value = dblSolId
		hdnPedId.value = dblPedId
		hdnSolAPGId.value = dblSolId

		target = self.name
		action = "Solicitacao.asp"
		submit()
	}
}
function Text1_onclick() {

}

function Reset1_onclick() {

}

</SCRIPT>

<%
Set objRSPag = Nothing
DesconectarCla()
%>
<%

function RetornaTabela(Usuario)


  'response.write "<script>alert('"&request("cboProcesso")&"')</script>"
   
	Vetor_Campos(1)="adWChar,50,adParamInput,"	
	Vetor_Campos(2)="adInteger,1,adParamInput," 
	Vetor_Campos(3)="adWChar,50,adParamInput,"	
	Vetor_Campos(4)="adInteger,1,adParamInput," 
	Vetor_Campos(5)="adInteger,1,adParamInput," 
	Vetor_Campos(6)="adWChar,5,adParamInput,"	& request("cboProcesso")
	Vetor_Campos(7)="adWChar,5,adParamInput,"	& request("cboAcao")
	Vetor_Campos(8)="adInteger,10,adParamInput," & Usuario
	Vetor_Campos(9)="adWChar,7,adParamInput," & request("txt_Oe_numero")
	Vetor_Campos(10)="adInteger,4,adParamInput," & request("txt_Oe_Ano")
	Vetor_Campos(11)="adInteger,4,adParamInput," & request("txt_Oe_Item")
	Vetor_Campos(12)="adWChar,60,adParamInput," & request("txt_Cliente")
	Vetor_Campos(13)="adWChar,100,adParamInput," & request("txt_Designacao")

    strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_tarefas_APG ",13,Vetor_Campos)

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
							"<th nowrap>&nbsp;Nº do Contrato</th>" & _
							"<th>&nbsp;Endereço</th>" & _
							"<th>&nbsp;Designação </th>" & _
						"</tr>"
		strXls = strHtml

		For intIndex = 1 to objRSPag.PageSize

			if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

			if isnull(objRSPag("Solicitacao")) Then strsolicitacao = "0" Else strsolicitacao = objRSPag("Solicitacao") End if

					If trim(objRSPag("Acao")) = "ATV" Then
				    
						StrAcao = "Ativacao"
					
					End if 


					If trim(objRSPag("Acao")) = "ALT" and trim(objRSPag("Processo")) <> "DES" Then
						
						StrAcao = "Alteracao"
					
					End if 
					
					If trim(objRSPag("Acao")) = "CAN" Then
						
						StrAcao = "Cancelamento"
					
					End if 
					
					If trim(objRSPag("Acao")) = "DES" OR (trim(objRSPag("Processo")) = "DES" and trim(objRSPag("Acao")) = "ALT")Then
						
						StrAcao = "Desativacao"
					
					End if
					
					
					Set objRS = db.Execute("SELECT acf_id FROM CLA_VIEW_SOLICITACAOMIN WHERE sol_id = '" & strsolicitacao & "'")
					if Not objRS.Eof and Not objRS.Bof then 
					  acf_id = objRS("acf_id")
					end if
					
					strHtml = strHtml & "<tr class='" & strClass & "'>" '& _
					
					if isnull(objRSPag("Solicitacao")) then 
						strHtml = strHtml & "<td>&nbsp;&nbsp;&nbsp;</td>"
					else
						strHtml = strHtml & "<td ><a href='javascript:DetalharItem(" & trim(objRSPag("solicitacao")) & ")'>...&nbsp;</a></td>"
					end if 
					
					'Alterado PRSS - 09/04/2007 - INICIO
					select case trim(objRSPag("Acao"))
					Case "DES"
					   strHtml = strHtml & "<td ><a href=" &chr(34)& "javascript:AtId('" & strsolicitacao & "','" & acf_id & "','" & objRSPag("ID_logico") &  "', '" & objRSPag("Sol_Acesso_ID") & "');AlterarSolicitacao(" & strsolicitacao & ",2)" &chr(34)&">" & "OE:" & right("00000" & Trim(objRSPag("OE_Numero")),5) & "/" & objRSPag("OE_Ano") & "</a></td>"
					Case "CAN"
					   strHtml = strHtml & "<td ><a href=" &chr(34)& "javascript:AtId('" & strsolicitacao & "','" & acf_id & "','" & objRSPag("ID_logico") &  "', '" & objRSPag("Sol_Acesso_ID") & "');AlterarSolicitacao(" & strsolicitacao & ",4)" &chr(34)&">" & "OE:" & right("00000" & Trim(objRSPag("OE_Numero")),5) & "/" & objRSPag("OE_Ano") & "</a></td>"
					Case else
						if(trim(objRSPag("Processo")) = "DES" and trim(objRSPag("Acao")) = "ALT") then
							strHtml = strHtml & "<td ><a href=" &chr(34)& "javascript:AtId('" & strsolicitacao & "','" & acf_id & "','" & objRSPag("ID_logico") &  "','" & objRSPag("Sol_Acesso_ID") &  "');AlterarSolicitacao(" & strsolicitacao & ",2)" &chr(34)&">" & "OE:" & right("00000" & Trim(objRSPag("OE_Numero")),5) & "/" & objRSPag("OE_Ano") & "</a></td>"
						else
					 	 strHtml = strHtml & "<td ><a href='javascript:EditarFac(" & strsolicitacao & "," & objRSPag("Sol_Acesso_ID") & ",""" & StrAcao & """)'>" & "OE:" & right("00000" & Trim(objRSPag("OE_Numero")),5) & "/" & objRSPag("OE_Ano") & "</a></td>"
						end if
					End select
										'PRSS - 09/04/2007 - FIM
										
					strHtml = strHtml & "<td >" & objRSPag("OE_Item") & "</td>" & _
										"<td >" & objRSPag("Razao_Social") & "</td>" & _
										"<td >" & objRSPag("Acao") & "</td>" & _
										"<td >" & objRSPag("Numero_Contrato_Cliente") & "</td>" & _
										"<td >" & Trim(objRSPag("Tipo_Logradouro")) & " " & Trim(objRSPag("Nome_Logradouro")) & ", " & _
											Trim(objRSPag("Numero_Predio")) & " " & Trim(objRSPag("Bairro")) & " " & _
											Trim(objRSPag("Cidade")) & " " & Trim(objRSPag("UF"))  & "</td>" & _
										"<td >" & Trim(objRSPag("Designacao_Servico")) & "</td>" & "</tr>"


					strXls = strXls & "<tr class='" & strClass & "'>" & _
										"<td ></td>" & _
										"<td >" & "OE:" & right("00000" & objRSPag("OE_Numero"),5) & "/" & objRSPag("OE_Ano") & "</td>" & _
										"<td >" & objRSPag("OE_Item") & "</td>" & _
										"<td >" & objRSPag("Razao_Social") & "</td>" & _
										"<td >" & objRSPag("Acao") & "</td>" & _
										"<td >" & objRSPag("Numero_Contrato_Cliente") & "</td>" & _

										"<td >" & Trim(objRSPag("Tipo_Logradouro")) & " " & Trim(objRSPag("Nome_Logradouro")) & ", " & _
											Trim(objRSPag("Numero_Predio")) & " " & Trim(objRSPag("Bairro")) & " " & _
											Trim(objRSPag("Cidade")) & " " & Trim(objRSPag("UF"))  & "</td>" & _

										"<td >" & Trim(objRSPag("Designacao_Servico")) & "</td>" & _
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