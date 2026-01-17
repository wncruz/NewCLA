<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
'EMBRATEL - Sistema CLA
'Arquivo: SolicitacaoAprov_Main.asp
'Descricao: Solicitacao vindas do sistemas Aprovisionador

Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
if Request.Form("hdnXmlReturn") <> "" then
    objXmlDados.loadXml(Request.Form("hdnXmlReturn"))
Else
    objXmlDados.loadXml("<xDados/>")
End if
%>

<script language='javascript' src="../javascript/help.js"></script>
<script language="JavaScript">
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")

function checa(f) { return true; }

function AtId(solid,acfid,vall,AprovisiId){
    with (document.forms[0]) {
        hdn678.value = vall
        hdnSolId.value = solid
        hdnAcfId.value = acfid
        hdnAprovisiId.value = AprovisiId
    }
}

function ConsultarPedidosPend() {
    with (document.forms[0]) {
        if (cboAcao.value == "" && txt_oe_numero.value == "") {
            alert("Selecione uma acao.")
            return;
        }
        if (cboAcao.value == "REC" && txt_oe_numero.value == "") {
            alert("Informe o numero da Order Entry.");
            txt_oe_numero.focus();
            return;
        }
        hdnAcao.value = "SEL"
        target = self.name
        action = "SolicitacaoAprov_main.asp?Consulta=1"
        submit()
    }
}

function DetalharItem(dblSolId) {
    with (document.forms[0]) {
        PopularXml()
        hdnSolId.value = dblSolId
        DetalharFac()
    }
}

function CompletarCampoIA(obj) {
    if (obj.value != "" && obj.value != 0) {
        var intLen = parseInt(obj.size) - parseInt(obj.value.length)
        switch (obj.TIPO.toUpperCase()) {
            case "N":
                for (var intIndex=0; intIndex<intLen; intIndex++) {
                    obj.value = "0" + obj.value
                }
                break
            default:
                for (var intIndex=0; intIndex<intLen; intIndex++) {
                    obj.value = obj.value + " "
                }
        }
    }
}

function EditarFac(dblSolId, AprovisiId, strAcao, StrOrigem) {
    with (document.forms[0]) {
        hdnSolId.value = dblSolId
        hdnAprovisiId.value = AprovisiId
        hdnOEOrigem.value = "Aprov"
        hdnAcao.value = strAcao
        hdnOrigem.value = StrOrigem
        action = "SolicitacaoAprov.asp"
        submit()
    }
}
</script>

<!-- TITULO DA PAGINA -->
<h2 style="color:#003366; margin-bottom:20px; padding-bottom:10px; border-bottom:2px solid #003366;">
    Solicitacao de Acesso - Aprovisionador
</h2>

<form action="SolicitacaoAprov_main.asp" name="Form1" method="post" onsubmit="return checa(this)">
<input type="hidden" name="hdnPedId">
<input type="hidden" name="hdnSolId">
<input type="hidden" name="hdnAprovisiId">
<input type="hidden" name="hdnXmlReturn">
<input type="hidden" name="hdnOEOrigem" value="Aprov">
<input type="hidden" name="hdnAcaoApg">

<!-- FORMULARIO DE FILTROS -->
<table style="width:100%; max-width:800px; margin-bottom:20px;">
    <tr class="clsSilver">
        <td style="width:150px; padding:10px;"><strong>Aprovisionador</strong></td>
        <td style="padding:10px;">
            <select name="Cbo_OrisolAprov" style="min-width:200px;">
                <option value=""></option>
                <%
                set rsCboOrisolAprov = db.execute("select Orisol_ID,OriSol_Alias from CLA_OrigemSolicitacao where Orisol_InterfAprov = 1 and Orisol_ID= 10 order by orisol_alias")
                do while not rsCboOrisolAprov.eof
                %>
                <option value="<%=rsCboOrisolAprov("Orisol_ID")%>" <%if request("Cbo_OrisolAprov") = rsCboOrisolAprov("Orisol_ID") then%>selected<%end if%>><%=ucase(rsCboOrisolAprov("OriSol_Alias"))%></option>
                <%
                rsCboOrisolAprov.movenext
                loop
                %>
            </select>
        </td>
    </tr>
    <tr class="clsSilver2">
        <td style="padding:10px;"><strong>Acao</strong></td>
        <td style="padding:10px;">
            <select name="cboAcao" style="min-width:200px;">
                <option value="">Selecione uma Acao</option>
                <option value="ATV" <%if request("cboAcao") = "ATV" then%>selected<%end if%>>ATIVACAO</option>
               <!-- <option value="CAN" <%if request("cboAcao") = "CAN" then%>selected<%end if%>>CANCELAMENTO</option> -->
                <option value="ALT" <%if request("cboAcao") = "ALT" then%>selected<%end if%>>ALTERACAO</option>
                <option value="DES" <%if request("cboAcao") = "DES" then%>selected<%end if%>>DESATIVACAO</option>
            </select>
        </td>
    </tr>
	<!--
    <tr class="clsSilver">
        <td style="padding:10px;"><strong>OE</strong></td>
        <td style="padding:10px;">
            <input id="txt_oe_numero" type="text" title="Numero" maxlength="10" size="11" name="txt_oe_numero" value='<%=request("txt_oe_numero")%>'> /
            <input id="txt_oe_ano" type="text" title="Ano" maxlength="4" size="5" name="txt_oe_ano" value='<%=request("txt_oe_ano")%>'> item
            <input id="txt_oe_item" type="text" title="Item" maxlength="3" size="4" name="txt_oe_item" value='<%=request("txt_oe_item")%>'>
        </td>
    </tr>
	-->
    <tr class="clsSilver2">
        <td style="padding:10px;"><strong>Solicitacao - CFD</strong></td>
        <td style="padding:10px;">
            <input id="txt_variavel" type="text" title="XXXXXXXX" maxlength="8" size="8" name="txt_variavel" value='<%=request("txt_variavel")%>' onblur="CompletarCampoIA(this)" TIPO="A">
            <input id="txt_ss" type="text" title="IA" maxlength="2" size="2" name="txt_ss" value='<%=request("txt_ss")%>' onblur="CompletarCampoIA(this)" TIPO="A">
            <input id="txt_num_sol" type="text" title="Numero" maxlength="4" size="4" name="txt_num_sol" value='<%=request("txt_num_sol")%>' onblur="CompletarCampoIA(this)" TIPO="N"> /
            <input id="txt_ano_sol" type="text" title="Ano" maxlength="4" size="4" name="txt_ano_sol" value='<%=request("txt_ano_sol")%>' onblur="CompletarCampoIA(this)">
        </td>
    </tr>
    <tr class="clsSilver">
        <td style="padding:10px;"><strong>Cliente</strong></td>
        <td style="padding:10px;">
            <input id="txt_cliente" type="text" maxlength="60" name="txt_cliente" size="50" value='<%=request("txt_cliente")%>' style="width:100%; max-width:400px;">
        </td>
    </tr>
    <tr class="clsSilver2">
        <td style="padding:10px;"><strong>Designacao</strong></td>
        <td style="padding:10px;">
            <input id="txt_designacao" type="text" maxlength="60" name="txt_designacao" size="50" value='<%=request("txt_designacao")%>' style="width:100%; max-width:400px;">
        </td>
    </tr>
</table>

<!-- BOTOES -->
<div style="margin:20px 0;">
    <input type="button" name="btconsulta" value="Consultar" class="button" onclick="ConsultarPedidosPend()">
    <input type="button" class="button" name="btnSair" value="Sair" onclick="javascript:window.location.replace('main.asp')" style="margin-left:10px;">
</div>

<iframe id="IFrmProcesso" name="IFrmProcesso" width="0" height="0" frameborder="0" scrolling="no"></iframe>

<span id="spnLinks"></span>

<%
Dim strClass, dblEstId, intIndex, strSql, intCount, strSel, strXls, strLink, strHtml

if Request.ServerVariables("CONTENT_LENGTH") = 0 then
    dblUsuIdFac = dblUsuId
End If

Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
if Request.Form("hdnXMLReturn") <> "" then
    objXmlDados.loadXml(Request.Form("hdnXMLReturn"))
    set objNodeAux = objXmlDados.getElementsByTagName("cboUsuario")
    if objNodeAux.length > 0 then dblUsuIdFac = objNodeAux(0).childNodes(0).text
End if

if dblUsuIdFac = "" then dblUsuIdFac = "null" End if

if Request.QueryString("Consulta") = "1" or Request.QueryString("btn") <> "" then
    strResponse = RetornaTabela(dblUsuId)
    Response.Write "<div style='margin-top:20px; overflow-x:auto;'>" & strResponse & "</div>"
    %>
    <!--#include file="../inc/ControlesPaginacao.asp"-->
    <%
end if
%>

<input type="hidden" name="hdnXls" value="<%=strXls%>">
<input type="hidden" name="hdnPaginaOrig" value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type="hidden" name="hdnAcao">
<input type="hidden" name="hdnOrigem">
<input type="hidden" name="hdnNomeCons" value="PedidosPend">
<input type="hidden" name="hdn678" value="<%=Request.Form("hdn678")%>">
<input type="hidden" name="hdnAcfId" value="<%=Request.Form("hdnAcfId")%>">
</form>

<script language="javascript">
document.getElementById('spnLinks').innerHTML = '<%=TratarAspasJS(strLink)%>'
</script>

<%
Set objRSPag = Nothing
DesconectarCla()

function RetornaTabela(Usuario)

	txt_ia = request("txt_variavel") + request("txt_ss") + request("txt_num_sol") + request("txt_ano_sol")   
	
	'response.write "<script>alert('"&txt_ia&"')</script>"
	
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
	Vetor_Campos(12)="adWChar,40,adParamInput," & txt_ia ' request("txt_ia")

    strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_Aprovisionador ",12,Vetor_Campos)
	
	'Response.Write 	strSql

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
							"<th >+</th>" & _
							"<th width=120>&nbsp;Numero/Ano OE</th>" & _
							"<th>&nbsp;Item OE</th>" & _
							"<th >&nbsp;Ponta</th>" & _
							"<th>&nbsp;Cliente</th>" & _
							"<th>&nbsp;Ação</th>" & _
							"<th nowrap>&nbsp;Serviço</th>" & _
							"<th>&nbsp;Endereço</th>" & _
							"<th>&nbsp;Designação <br>&nbsp;do Serviço</th>" & _
							"<th>&nbsp;Aprovisionador </th>" & _
						"</tr>"
		strXls = strHtml

		For intIndex = 1 to objRSPag.PageSize

			if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if

			if isnull(objRSPag("sol_id")) Then strsolicitacao = "0" Else strsolicitacao = objRSPag("sol_id") End if

			If trim(objRSPag("Acao")) = "ATV" Then StrAcao = "Ativacao" 	End if 

			If trim(objRSPag("Acao")) = "ALT" Then StrAcao = "Alteracao" 	End if 
			
			If trim(objRSPag("Acao")) = "CAN" Then StrAcao = "Cancelamento" End if 
			
			If trim(objRSPag("Acao")) = "DES" Then StrAcao = "Desativacao" 	End if
			
			If trim(request("cboAcao")) = "REC" Then 
			
				StrOrigem = "Recadastro" 	
			
			End if
			
			If trim(objRSPag("Acao")) = "CAN" or trim(objRSPag("Acao")) = "DES" then	
				
				Set objRS = db.Execute("SELECT acf_id FROM CLA_VIEW_SOLICITACAOMIN WHERE sol_id = '" & strsolicitacao & "'")
				if Not objRS.Eof and Not objRS.Bof then  acf_id = objRS("acf_id") end if
				
			End IF
			strHtml = strHtml & "<tr class='" & strClass & "'>" '& _
			
			if isnull(objRSPag("sol_id")) then 
				strHtml = strHtml & "<td>&nbsp;&nbsp;&nbsp;</td>"
			else
				strHtml = strHtml & "<td ><a href='javascript:DetalharItem(" & trim(objRSPag("sol_id")) & ")'>...&nbsp;</a></td>"
			end if 
			
			select case trim(objRSPag("Acao"))
			Case "DES"
				if Trim(objRSPag("OriSol_Descricao")) = "CFD" then
			  	 strHtml = strHtml & "<td ><a href=" &chr(34)& "javascript:AtId('" & strsolicitacao & "','" & acf_id & "','" & objRSPag("acl_idacessologico") &  "', '" & objRSPag("Aprovisi_id") & "');DesCanSolicitacao(" & strsolicitacao & ",2)" &chr(34)&">"  &  objRSPag("IA") &  "</a></td>"
			    else
			   strHtml = strHtml & "<td ><a href=" &chr(34)& "javascript:AtId('" & strsolicitacao & "','" & acf_id & "','" & objRSPag("acl_idacessologico") &  "', '" & objRSPag("Aprovisi_id") & "');DesCanSolicitacao(" & strsolicitacao & ",2)" &chr(34)&">" & "OE:" & right("00000000" & Trim(objRSPag("OE_Numero")),8) & "/" & objRSPag("OE_Ano") & "</a></td>"
				end if 	
			Case "CAN"
			   strHtml = strHtml & "<td ><a href=" &chr(34)& "javascript:AtId('" & strsolicitacao & "','" & acf_id & "','" & objRSPag("acl_idacessologico") &  "', '" & objRSPag("Aprovisi_id") & "');DesCanSolicitacao(" & strsolicitacao & ",4)" &chr(34)&">" & "OE:" & right("00000000" & Trim(objRSPag("OE_Numero")),8) & "/" & objRSPag("OE_Ano") & "</a></td>"
			
			Case else
				if Trim(objRSPag("OriSol_Descricao")) = "CFD" then
			   			strHtml = strHtml & "<td ><a href='javascript:EditarFac(" & strsolicitacao & "," & objRSPag("Aprovisi_id") & ",""" & StrAcao & """,""" & StrOrigem & """)'>" &  objRSPag("IA") & "</a></td>"
				ELSE
			   strHtml = strHtml & "<td ><a href='javascript:EditarFac(" & strsolicitacao & "," & objRSPag("Aprovisi_id") & ",""" & StrAcao & """,""" & StrOrigem & """)'>" & "OE:" & right("00000000" & Trim(objRSPag("OE_Numero")),8) & "/" & objRSPag("OE_Ano") & "</a></td>"
				END IF 
			
			End select
			
			if Trim(objRSPag("OriSol_Descricao")) = "SGAV" then
			  strOriSol_Descricao = "SGA Voz"
			elseif Trim(objRSPag("OriSol_Descricao")) = "SGAP" then
			  strOriSol_Descricao = "SGA Plus"
			else
			  strOriSol_Descricao = Trim(objRSPag("OriSol_Descricao"))
			end if
																
			set objRSserv = db.execute("CLA_sp_sel_servico null,null,'"& objRSPag("Ser_Desc") &"',null")
			if Not objRSserv.Eof and not objRSserv.Bof then
				Servico = objRSserv("Ser_Desc")
			end if
																
			strHtml = strHtml & "<td >" & objRSPag("OE_Item") & "</td>" & _
								"<td >" & objRSPag("Id_Endereco") & "</td>" & _
								"<td >" & objRSPag("Cli_Nome") & "</td>" & _
								"<td >" & objRSPag("Acao") & "</td>" & _
								"<td >" & objRSPag("Ser_Desc") & "</td>" & _
								"<td >" & Trim(objRSPag("Tpl_Sigla")) & " " & Trim(objRSPag("End_NomeLogr")) & ", " & _
									Trim(objRSPag("End_NroLogr")) & " " & Trim(objRSPag("End_Bairro")) & " " & _
									Trim(objRSPag("Cidade")) & " " & Trim(objRSPag("Est_Sigla"))  & "</td>" & _
								"<td >" & Trim(objRSPag("Acl_designacaoServico")) & "</td>" & _
								"<td >" & Ucase(strOriSol_Descricao) & "</td>" & "</tr>"

			strXls = strXls & "<tr class='" & strClass & "'>" & _
								"<td></td>"& _
								
								"<td >" & "OE:" & objRSPag("OE_Numero") & "/" & objRSPag("OE_Ano") & "</td>" & _
								
								"<td >" & objRSPag("OE_Item") & "</td>" & _
								"<td >" & objRSPag("Id_Endereco") & "</td>" & _
								"<td >" & objRSPag("Cli_Nome") & "</td>" & _
								"<td >" & objRSPag("Acao") & "</td>" & _
								"<td >" & objRSPag("Acl_NContratoServico") & "</td>" & _

								"<td >" & Trim(objRSPag("Tpl_Sigla")) & " " & Trim(objRSPag("End_NomeLogr")) & ", " & _
									Trim(objRSPag("End_NroLogr")) & " " & Trim(objRSPag("End_Bairro")) & " " & _
									Trim(objRSPag("Cidade")) & " " & Trim(objRSPag("Est_Sigla"))  & "</td>" & _

								"<td >" & Trim(objRSPag("Acl_designacaoServico")) & "</td>" & _
								"<td >" & Ucase(strOriSol_Descricao) & "</td>" & _
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

<!--#include file="../inc/footer.asp"-->