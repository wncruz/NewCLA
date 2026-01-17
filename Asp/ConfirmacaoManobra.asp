<!--#include file="../inc/data.asp"-->

<!--#include file="../inc/header.asp"-->
<%
dblPedId = Request.Form("hdnPedId")

set fac = db.execute("CLA_sp_sel_facilidade " & dblPedId)

if fac.Eof then%>
	<table border=0 align=center>
		<tr>
			<td align=center>
				<br>Registros não encontrado<br>
			</td>
		</tr>
		<tr>
			<td align=center>
				<input type=button name=btnValotar class=button value=Voltar onClick="javascript:window.location.replace('PendenciaManobra.asp')">&nbsp;
				<input type=button name=btnSair class=button value=Sair onClick="javascript:window.location.replace('main.asp')">
			</td>
		</tr>
	<table>
<%	
	Response.End 
End if
%>
<form name=Form1 method="post">
<input type="hidden" name="hdnAcao">

<table border="0"  cellspacing="1" cellpadding="0" width=760 >
<tr>
	<th colspan=2><p align="center">Pendência de Manobra</p></th>
</tr>
<tr class=clsSilver>
	<th colspan=2>&nbsp;•&nbsp;Recurso</th>
</tr>
<tr class=clsSilver>
	<td width=150px>Estação</td>
	<td>
		<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="20%" >
			<tr><td class="lightblue">&nbsp;
				<%	dblEscId = fac("Esc_Id")
					dblDisId = fac("Dst_id")
					dblSisId = fac("Sis_id")
					dblProId = fac("Pro_id")
					dblRecId = fac("Rec_IDEntrega")
					set objRS = db.execute("CLA_sp_sel_estacao 0")
					While not objRS.Eof 
						if Cdbl("0" & dblEscId) = Cdbl("0" & objRS("Esc_ID")) then 
							Response.Write objRS("Cid_Sigla") & "  " & objRS("Esc_Sigla")
						End if	
						objRS.MoveNext
					Wend
				%>
			</td></tr>	
		</table>	
	</td>
</tr>
<tr class=clsSilver>
	<td>Distribuidor</td>
	<td>
		<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="20%" >
			<tr><td class="lightblue">&nbsp;
			<%	if Trim(dblEscId) <> "" then 
					set objRS = db.execute("CLA_sp_view_recursodistribuicao " & dblEscId)
					While not objRS.Eof 
						if Trim(dblDisId) = Trim(objRS("Dst_ID")) then 
							Response.Write objRS("Dst_Desc")
						End if	
						objRS.MoveNext
					Wend
				End if	
			%>
			</td></tr>	
		</table>	
	</td>
</tr>
<tr class=clsSilver>
	<td>Rede</td>
	<td>
		<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="30%" >
			<tr><td class="lightblue">&nbsp;
			<%
				Set objRS = db.execute("CLA_sp_sel_sistema 0")
				While not objRS.Eof 
					if Trim(dblSisId) = Trim(objRS("Sis_ID")) then 
						Response.Write Ucase(objRS("Sis_Desc"))
					End if	
					objRS.MoveNext
				Wend
			%>
			</td></tr>	
		</table>	
	</td>
</tr>
<tr class=clsSilver>
	<td>Provedor</td>
	<td>
		<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="2" bordercolorlight="#003388" bordercolordark="#ffffff" width="30%" >
			<tr><td class="lightblue">&nbsp;
			<%	set objRS = db.execute("CLA_sp_sel_provedor 0")
				While not objRS.Eof 
					if Trim(dblProId) = Trim(objRS("Pro_ID")) then 
						Response.Write objRS("Pro_Nome")
					End if
					objRS.MoveNext
				Wend
			%>
			</td></tr>	
		</table>	
	</td>
</tr>
</table>

<table rules="groups" border=0 cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="760">
		<tr class=clsSilver>
			<th colspan=3>&nbsp;•&nbsp;Informe a Nova PADE/PAC</th>
		</tr>
		<tr class=clsSilver>
			<td width=150>&nbsp;</td>
			<td width=150>Nº de Acesso</td>
			<td>PADE/PAC</td>		
		</tr>
		<%
		while not fac.eof

			strNroAcesso = fac("Acf_nroacessoPtaEbt")
			dblFacId = fac("Fac_Id")

			Vetor_Campos(1)="adInteger,2,adParamInput,"
			Vetor_Campos(2)="adInteger,2,adParamInput,"
			Vetor_Campos(3)="adWchar,1,adParamInput,T"
			Vetor_Campos(4)="adInteger,2,adParamInput," & dblFacId
			Vetor_Campos(5)="adWChar,25,adParamInput," & strNroAcesso

			strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_interligacaolib",5,Vetor_Campos)
			Set objRSLib =  db.Execute(strSqlRet)
			intCount = 1
			if Not objRSLib.Eof and Not objRSLib.Bof then
				While not objRSLib.Eof 
					strInterLigOig  = objRSLib("Int_CorOrigem") '& "&nbsp;&nbsp;>&nbsp;&nbsp;" & 
					strInterLigDest = objRSLib("Int_CorDestino")
					%>
					<tr class=clsSilver>
						<td>De:</td>
						<td><input type="text" class="text" name="numeroacessode<%=intCount%>"	value="<%=strNroAcesso%>" style="background-color:#eeeeee" readOnly tabindex=-1>
						<td>
							<input type="text" class="text" name="coordenadade<%=intCount%>" onClick="document.forms[0].coordenadapara<%=intCount%>.value=this.value;" style="cursor:hand"	value="<%=strInterLigOig%>" style="background-color:#eeeeee" maxlength="20" size="15" ReadOnly tabindex=-1>
							<input type=hidden name=fac<%=intCount%> value="<%=fac("Fac_Id")%>">
						</td>
					</tr>
		
					<tr class=clsSilver>
						<td>Para:</td>
						<td>&nbsp;</td>
						<td>
							<input type="text" class="text" name="coordenadapara<%=intCount%>" maxlength="20" size="15">
							<input type="button" class="button" style="width:30px" name="btnProcurar<%=intCount%>" value="..." onclick="javascript:PosicoesLivre('L',<%=dblRecId%>)">
							<input type="button" class="button" style="width:30px" name="apaga<%=intCount%>" value="X" onClick="javascript:apaga(<%=intCount%>);">
						</td>
					</tr>
					<%
					intCount = intCount + 1
					objRSLib.MoveNext
				Wend 
			End if
			fac.MoveNext
		Wend
		%>
</table>

<table width="760" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td align=center height=35px>
		<input type="button" class="button" name="btnGravar" value="Gravar" onclick="Gravar()">
		<input type=button name=btnValotar class=button value=Voltar onClick="javascript:window.location.replace('PendenciaManobra.asp')">&nbsp;
		<input type=button name=btnSair class=button value=Sair onClick="javascript:window.location.replace('main.asp')">
	</td>
</tr>
</table>
<input type="hidden" name="estacao"			value="<%=dblEscID%>">
<input type="hidden" name="distribuicao"	value="<%=dblDisID%>">
<input type="hidden" name="sistema"			value="<%=dblSisID%>">
<input type="hidden" name="provedor"		value="<%=dblProID%>">
<input type="hidden" name="hdnCount"		value="<%=intCount-1%>">
<input type="hidden" name="hdnRecId"		value="<%=dblRecID%>">
<input type="hidden" name="tipo"			value="I">
</form>
<SCRIPT LANGUAGE=javascript>
<!--
function Gravar()
{
	with (document.forms[0])
	{
		for(var intIndex=0;intIndex<<%=intCount-1%>;intIndex++)
		{
			var objCood = new Object(eval("coordenadapara"+parseInt(intIndex+1)))
			if (!ValidarCampos(objCood,intIndex+"º Coordenação \"Para\".")) return
		}	
		hdnAcao.value = "TrocarInterLigacao"
		target = "IFrmProcesso"
		action = "ProcessoManobra.ASP"
		submit()
	}	
}

var objAryFac = new Array(<%=intCount-1%>)
for (var intIndex=0;intIndex<objAryFac.length;intIndex++)
{
	objAryFac[intIndex] = ""
}

var objAryFacRet

function PosicoesLivre(strPagina,intRecId)
{
	var intCont = 0
	for (var intIndex=0;intIndex<objAryFac.length;intIndex++)
	{
		var objForm = new Object(eval("document.forms[0].coordenadapara"+parseInt(intIndex+1)))
		if (objForm.value != ""){
			objAryFac[intIndex] = objForm.value
			intCont += 1
		}	
	}
	if (intCont == objAryFac.length){
		alert("Todas as posições estão preenchidas.\nA nova seleção substituirá as atuais.")
		for (var intIndex=0;intIndex<objAryFac.length;intIndex++)
		{
			objAryFac[intIndex] = ""
		}
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
					eval("document.forms[0].coordenadapara"+parseInt(intIndex+1)+".value = '"+objAryFacRet[intIndex].split(",")[0]+"'")
				}	
			}
		}catch(e){}	
	}	
}

function apaga(indice) {
	eval("document.forms[0].coordenadapara"+indice+".value = ''")
	objAryFac[parseInt(indice-1)] = ''
}
-->
</SCRIPT>

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
