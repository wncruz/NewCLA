<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<table width="780" border="0" cellspacing="0" cellpadding="0">
<tr><td>
<SCRIPT LANGUAGE=javascript>
<!--
function ResgatarDistribuicao(obj,intDistrib)
{
	with (document.forms[0])
	{
		if (obj.value != "")
		{
			hdnAcao.value = "ResgatarDistribuicao"
			hdnEstacaoAtual.value = obj.value
			hdnNomeLocal.value = "DistLocalInstala"
			target = "IFrmProcesso"
			action = "ProcessoCla.asp"
			submit()
		}	
	}
}
function LimparRdi()
{
	with (document.forms[0])
	{
		for (var intIndex=0;intIndex<objAryFac.length;intIndex++)
		{
			objAryFac[intIndex] = ""
		}
		spnCoordenada1.innerHTML = "";
		spnCoordenada2.innerHTML = "";
		spnCoordenada3.innerHTML = "";
		spnCoordenada4.innerHTML = "";
		hdnRdiId.value = ""
	}
	LimparForm()
}

function GravarRdi() {
	with (document.forms[0]){

		if (!ValidarCampos(cboLocalInstala,"Estação")) return
		if (!ValidarCampos(cboDistLocalInstala,"Distribuidor")) return
		if (!ValidarCampos(cboProvedor,"Provedor")) return

		if (!ValidarCampos(txtNomeCli,"Cliente")) return
		if (!ValidarCampos(txtEndereco,"Endereço")) return

		if (!ValidarTipoInfo(txtDtInicioTemp,1,"Data início temporário")) return false;
		if (!ValidarTipoInfo(txtDtFimTemp,1,"Data fim temporário")) return false;

		if (IsEmpty(txtCoordenada1.value) && IsEmpty(txtCoordenada2.value) && IsEmpty(txtCoordenada3.value) && IsEmpty(txtCoordenada4.value)) {
			alert("Informe pelo menos uma PADE/PAC.")
			return false;
		}

		hdnAcao.value = "GravarRedeInterna"
		target = "IFrmProcesso"
		action = "ProcessoRedeInterna.asp"
		submit()
		return true;
	}	
}

function apaga(indice) {
	objAryFac[parseInt(indice-1)] = ""
	if (indice == 1) {
		Form1.txtCoordenada1.value = "";
		spnCoordenada1.innerHTML = "";
	}
	if (indice == 2) {
		Form1.txtCoordenada2.value = "";
		spnCoordenada2.innerHTML = "";
	}
	if (indice == 3) {
		Form1.txtCoordenada3.value = "";
		spnCoordenada3.innerHTML = "";
	}
	if (indice == 4) {
		Form1.txtCoordenada4.value = "";
		spnCoordenada4.innerHTML = "";
	}
}

var objAryFac = new Array(4)
for (var intIndex=0;intIndex<objAryFac.length;intIndex++)
{
	objAryFac[intIndex] = ""
}

var objAryFacRet

function PosicoesLivre(strPagina,intRecId)
{
	for (var intIndex=0;intIndex<objAryFac.length;intIndex++)
	{
		objAryFac[intIndex] = ""
	}

	with (document.forms[0])
	{
		switch (strPagina)
		{
			case "L": //livres
				objAryFacRet = window.showModalDialog('interligacoeslivres.asp?rec_id='+intRecId+'&qtd=10',objAryFac,'dialogHeight: 200px; dialogWidth: 350px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
				break
			case "O": //Ocupada
				objAryFacRet = window.showModalDialog('consultainterocupadas_main.asp?rec_id='+intRecId,objAryFac,'dialogHeight: 350px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
				break
		}

			//Preencha as coordenadas
		try{	
			for (var intIndex=0;intIndex<objAryFacRet.length;intIndex++)
			{
				if (objAryFacRet[intIndex] != ""){
					eval("document.forms[0].txtCoordenada"+parseInt(intIndex+1)+".value = '"+objAryFacRet[intIndex].split(",")[0]+"'")
					eval("spnCoordenada"+parseInt(intIndex+1)+".innerHTML = '"+objAryFacRet[intIndex].split(",")[0]+" > "+objAryFacRet[intIndex].split(",")[1] +"'")
				}	
			}
		}catch(e){}	
	}	
}

function ListaInterligacaoLivre()
{
	with (document.forms[0]){
		if (!ValidarCampos(cboLocalInstala,"Estação")) return
		if (!ValidarCampos(cboDistLocalInstala,"Distribuidor")) return
		if (!ValidarCampos(cboProvedor,"Provedor")) return

		hdnAcao.value = "ListaInterligacaoLivre"
		target = "IFrmProcesso"
		action = "ProcessoRedeInterna.asp"
		submit()
	}	
}

function EditarRdi(obj){
	with (document.forms[0]){
		hdnRdiId.value = obj.value
		hdnAcao.value = "EditarRdi"
		target = "IFrmProcesso"
		action = "ProcessoRedeInterna.asp"
		submit()
	}
}
function RemoverRdi()
{
	with (document.forms[0]){
	 if (hdnRdiId.value == ""){
		alert("Favor editar a rede interna primeiro.")
		return
	 }else{
		hdnAcao.value = "RemoverRdi"
		target = "IFrmProcesso"
		action = "ProcessoRedeInterna.asp"
		submit()
	 }
	}	
}
//-->
</SCRIPT>

<form method="post" name=Form1>
<input type="hidden" name="hdnEstacaoAtual">
<input type="hidden" name="hdnNomeLocal">
<input type="hidden" name="cboSistema" value="<%=TroncoPar%>">
<input type="hidden" name="hdnAcao">
<input type="hidden" name="hdnRdiId">
<input type="hidden" name="hdnDstId">
<input type="hidden" name="hdnDtPedido" value="<%=right("0" & day(date), 2) & "/" & right("0" & month(date), 2) & "/" & year(date)%>">

<table border="0" width="760" cellspacing="1" cellpadding="0">
<tr>
	<th colspan=2><p align="center">Cadastro de Rede Interna</p></th>
</tr>
<tr class=clsSilver>
	<td width=150px nowrap><font class="clsObrig">:: </font>Estação</td>
	<td  >
		<select name="cboLocalInstala" onChange="ResgatarDistribuicao(this,1)">
			<option value=""></option>
			<%set objRS = db.execute("CLA_sp_sel_estacao 0")
				While not objRS.Eof 
					strItemSel = ""
					if Trim(strLocalInstala) = Trim(objRS("Esc_ID")) then strItemSel = " Selected " End if
					Response.Write "<Option value=" & objRS("Esc_ID") & strItemSel & ">" & objRS("Cid_Sigla") & "  " & objRS("Esc_Sigla") & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
			%>
		</select>
	</td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Distribuidor</td>
	<td >
		<span id="spnDistLocalInstala">
		<select name="cboDistLocalInstala" >
			<option value=""></option>
			<%	if Trim(strLocalConfig) <> "" then 
					set objRS = db.execute("CLA_sp_view_recursodistribuicao " & strLocalConfig)
					While not objRS.Eof 
						strItemSel = ""
						if Trim(strDistrib) = Trim(objRS("Dst_ID")) then strItemSel = " Selected " End if
						Response.Write "<Option value=" & objRS("Dst_ID") & strItemSel & ">" & objRS("Dst_Desc") & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
				End if	
			%>
		</select>
		</span>
	</td>
</tr>
<tr class=clsSilver>
	<td><font class="clsObrig">:: </font>Provedor</td>
	<td>
		<select name="cboProvedor">
		<%
		Set pro = db.execute("CLA_sp_view_escpro " & dblUsuId)
		if not pro.eof then
			%>
			<option value=""></option>
			<%
			do while not pro.eof
			%>
				<option value="<%=pro("Pro_ID")%>"><%=pro("Pro_Nome")%></option>
			<%
				pro.movenext
			loop
		end if
		%>
		</select>
	</td>
</tr>
<tr class=clsSilver>
	<td width="160">&nbsp;Data do Pedido</td>
	<td class="lightblue"><span id=spnDtPedido><%Response.write right("0" & day(date), 2) & "/" & right("0" & month(date), 2) & "/" & year(date)%></span>
	</td>
</tr>

	<tr class=clsSilver>
		<td><font class="clsObrig">:: </font>Nome do Cliente</td>
		<td><input type="text" class="text" name="txtNomeCli" size="50" maxlength="50">
		</td>
	</tr>
				
	<tr class=clsSilver>
		<td><font class="clsObrig">:: </font>Endereço do Cliente</td>
		<td><input type="text" class="text" name="txtEndereco" size="60" maxlength="60">
		</td>
	</tr>
		
	<tr class=clsSilver>
		<td>&nbsp;Designação</td>
		<td><input type="text" class="text" name="txtDesignacao" size="30" maxlength="30">
		</td>
	</tr>
		
	<tr class=clsSilver>
		<td>&nbsp;Nº Contrato</td>
		<td><input type="text" class="text" name="txtNroContrato" size="30" maxlength="50">
		</td>
	</tr>
				
	<tr class=clsSilver>
		<td>&nbsp;Data Início Temporário</td>
		<td><input type="text" class="text" name="txtDtInicioTemp" size="10" maxlength="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;(dd/mm/aaa)
		</td>
	</tr>
				
	<tr class=clsSilver>
		<td>&nbsp;Data Fim Temporário</td>
		<td><input type="text" class="text" name="txtDtFimTemp" size="10" maxlength="10" onKeyPress="OnlyNumbers();AdicionaBarraData(this)">&nbsp;(dd/mm/aaa)
		</td>
	</tr>
				
	<tr class=clsSilver>
		<td>&nbsp;Observação</td>
		<td><textarea name="txtObs" cols="60" rows="3" onkeydown="MaxLength(this,300);" ></textarea>A(300)
		</td>
	</tr>
	<tr class=clsSilver>
		<th colspan="2">&nbsp;•&nbsp;Facilidades</th>
	</tr>
	<tr align=center width=100% >
		<td colspan="2" class="lightblue">
		<table border="0" width=100% cellspacing="1" cellpadding="0">
		<tr>
			<th>&nbsp;Coordenada</th>
			<th>&nbsp;Ação</th>
			<th>&nbsp;Origem > Destino</th>

		</tr>	
		<%For intIndex=1 to 4%>
			<tr class=clsSilver>
				<td width=250 align=right ><%=intIndex%>&nbsp;•&nbsp;
					<input type="text" class="text" name="txtCoordenada<%=intIndex%>" size="20" maxlength="20">
				</td>
				<td width=80 align=center>
					<input type="button" class="button" style="width:30px" name="btnProcurar<%=intIndex%>" value="..." onclick="ListaInterligacaoLivre()">
					<input type="button" class="button" style="width:30px" name="apaga<%=intIndex%>" value="X" onClick="javascript:apaga(<%=intIndex%>);">
				</td>
				<td><span id=spnCoordenada<%=intIndex%>></span></td>	
			</tr>
		<%next%>
		</table>	
	</td>
</tr>
<tr>
	<td colspan=2 align=center height=30>
		<input type="button" class="button" name="btnGravar" value="Gravar" onClick="GravarRdi()">&nbsp;
		<input type="button" class="button" name="btnRemover" value="Remover" onClick="RemoverRdi()">&nbsp;
		<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="LimparRdi()">&nbsp;
		<input type="button" class="button" name="sair" value="Sair" onClick="javascript:window.location.replace('main.asp');">&nbsp;
	</td>	
</tr>	
</td>
</tr>
</table>
<table border="0" cellspacing="1" cellpadding=0 width="760">
<tr>
	<th colspan=8><p align="center">Rede Interna Cadastradas</p></th>
</tr>
<tr>
	<th>&nbsp;</th>
	<th>&nbsp;Estação</th>
	<th>&nbsp;Distribuidor</th>
	<th>&nbsp;Provedor</th>
	<th>&nbsp;Data</th>
	<th>&nbsp;Cliente</th>
	<th>&nbsp;Designação</th>
	<th>&nbsp;Nº Contrato</th>
</tr>
<%
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
strSql = "CLA_sp_sel_redeInterna"
Call PaginarRS(0,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize
		if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
		%>
		<tr class="<%=strClass%>">
			<td ><input type=Radio name=rdoCoordenada value="<%=objRSPag("Rdi_Id")%>" onClick="EditarRdi(this)"></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("Esc_Sigla")) & " " & TratarAspasHtml(objRSPag("cid_Sigla"))%></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("Dst_Desc"))%></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("Pro_Nome"))%></td>
			<td >&nbsp;<%=Formatar_Data(objRSPag("Rdi_DataPed"))%></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("Rdi_CliNome"))%></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("Rdi_Designacao"))%></td>
			<td >&nbsp;<%=TratarAspasHtml(objRSPag("Rdi_NroContrato"))%></td>
		</tr>
		<%
		intCount = intCount+1
		objRSPag.MoveNext
		if objRSPag.EOF then Exit For
	Next
End if
%>
</table>
<!--#include file="../inc/ControlesPaginacao.asp"-->
</form>
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