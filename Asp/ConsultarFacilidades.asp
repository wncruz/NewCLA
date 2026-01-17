<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ConsultarFacilidade.asp
'	- Responsável		: Vital
'	- Descrição			: Consulta de Facilidades na Alocação
%>
<!--#include file="../inc/data.asp"-->
<%

strStsFac = Request("strStsFac")
if strStsFac = "O" then
	strTitulo =  "Consulta de Posições Ocupadas"
Else
	strTitulo = "Consulta de Posições Disponíveis"
End if

Function ResgatarCamposFacilidade(strRede)

	strRet = ""
	Select Case int("0" & strRede)

		Case 1

			strRet = strRet & "<table align=center cellspacing=1 cellpadding=1 width=540 >"
			strRet = strRet & "<tr>"
			strRet = strRet & "<th colspan=9>"& strTitulo & "</th>"
			strRet = strRet & "</tr>"
			strRet = strRet & "<tr class=clsSilver2>"
			strRet = strRet & "<td>&nbsp;Fila</td>"
			strRet = strRet & "<td>&nbsp;Bastidor</td>"
			strRet = strRet & "<td>&nbsp;Régua</td>"
			strRet = strRet & "<td>&nbsp;Posição</td>"
			strRet = strRet & "<td>&nbsp;Domínio</td>"
			strRet = strRet & "<td>&nbsp;Nó</td>"
			strRet = strRet & "<td>&nbsp;Slot</td>"
			strRet = strRet & "<td>&nbsp;Porta</td>"
			strRet = strRet & "<td>&nbsp;Timeslot</td>"
			strRet = strRet & "	</tr>"


			strRet = strRet & "<tr class=clsSilver>"
			strRet = strRet & "	<td><input type=text class=text name=txtFila		maxlength=5	onKeyUp=""ValidarTipo(this,2)"" size=6></td>"
			strRet = strRet & "	<td><input type=text class=text name=txtBastidor	maxlength=3 onKeyUp=""ValidarTipo(this,0)"" size=4></td>"
			strRet = strRet & "	<td><input type=text class=text name=txtRegua		maxlength=2 onKeyUp=""ValidarTipo(this,2)"" size=3></td>"
			strRet = strRet & "	<td><input type=text class=text name=txtPosicao		maxlength=3 onKeyUp=""ValidarTipo(this,2)"" size=4></td>"
			
			dim objRSPlataforma , strPlataforma
			
			set objRSPlataforma  = db.execute("Cla_sp_Sel_Plataforma " & Request("cboPlataforma") )
			
			if not objRSPlataforma.eof and not objRSPlataforma.bof then 
				strPlataforma	= trim(objRSPlataforma("Pla_TipoPlataforma"))
			end if 
	
			set objRSPlataforma  = nothing 
			
			if  strPlataforma <> "LUCENT" then 
			
					strRet = strRet & "	<td><input type=text class=text name=txtDominio		maxlength=20 onKeyUp=""ValidarTipo(this,2)"" size=10></td>"
					strRet = strRet & "	<td><input type=text class=text name=txtNO			maxlength=20 onKeyUp=""ValidarTipo(this,2)"" size=10></td>"
					strRet = strRet & "	<td><input type=text class=text name=txtSlot		maxlength=2   size=2 onBlur='SearchDom(this,""A1,A2,A3,A4,A5,A6,A7,A8,B1,B2,B3,B4,B5,B6,B7,B8"")'>(A2)</td>"
					strRet = strRet & "	<td><input type=text class=text name=txtPorta		maxlength=1   size=1 onBlur='SearchDom(this,""A,B"")'>(A1)</td>"

			else
					strRet = strRet & "	<td><input type=text class=text name=txtDominio		maxlength=20  size=10>(A10)</td>"
					strRet = strRet & "	<td><input type=text class=text name=txtNO			maxlength=2   size=2 onKeyUp=""ValidarTipo(this,0)"" >(N2)</td>"
					strRet = strRet & "	<td><input type=text class=text name=txtSlot		maxlength=2   size=2 onBlur='SearchDom(this,""01,02,03,04,05,06,07,08,09,10,11,12"")' onKeyUp=""ValidarTipo(this,2)"">(N2)</td>"
					strRet = strRet & "	<td><input type=text class=text name=txtPorta		maxlength=2   size=2 onBlur='SearchDom(this,""01,02,03,04,05,06,07,08,09,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63"")' onKeyUp=""ValidarTipo(this,2)"">(N2)</td>"			
			
			end if 

			strRet = strRet & "	<td><input type=text class=text name=txtTimeSlot	maxlength=9 onKeyUp=""ValidarNTipo(this,0,4,4,1,0,4)"" size=10></td>"
			strRet = strRet & "</tr>"
			
			strRet = strRet & "<tr class=clsSilver>"
			strRet = strRet & "	<td align=center>(A5)</td>"
			strRet = strRet & "	<td align=center>(N3)</td>"
			strRet = strRet & "	<td align=center>(A2)</td>"
			strRet = strRet & "	<td align=center>(A3)</td>"
			
			if  strPlataforma <> "LUCENT" then 
				strRet = strRet & "	<td align=center>(A20)</td>"
				strRet = strRet & "	<td align=center>(A20)</td>"
				strRet = strRet & "	<td align=center>(A2)</td>"
				strRet = strRet & "	<td align=center>(A1)</td>"
			else
				strRet = strRet & "	<td align=center>(A10)</td>"
				strRet = strRet & "	<td align=center>(N2)</td>"
				strRet = strRet & "	<td align=center>(N2)</td>"
				strRet = strRet & "	<td align=center>(N2)</td>"
			end if 
			
			strRet = strRet & "	<td align=center>(N4-N4)</td>"
			strRet = strRet & "</tr>"
			strRet = strRet & "</table>"
			
		Case 2
		
			'Não deterministico
			strRet = ""
			strRet = strRet & "<table align=center cellspacing=1 cellpadding=1 width=540>"
			strRet = strRet & "	<tr>"
			strRet = strRet & "		<th colspan=2>"& strTitulo & "</th>"
			strRet = strRet & "	</tr>"
			strRet = strRet & "	<tr >"
			strRet = strRet & "	<tr class=clsSilver2>"
			strRet = strRet & "		<td>&nbsp;Tronco</td>"
			strRet = strRet & "		<td>&nbsp;Par</td>"
			strRet = strRet & "	</tr>"
			strRet = strRet & "<tr class=clsSilver>"
			strRet = strRet & "	<td><input type=text class=text name=txtTronco  maxlength=20 size=10>(A20)</td>"
			strRet = strRet & "	<td><input type=text class=text name=txtPar  maxlength=9 size=10 onKeyUp=""if(this.value.length > 1) ValidarNTipo(this,0,4,4,1,0,4)"" TIPO=N >(N4-N4) ou (A1)</td>"
			strRet = strRet & "</tr>"

			strRet = strRet & "</table>"

		Case 3

			strRet = ""
			strRet = strRet & "<table align=center cellspacing=1 cellpadding=1 width=540>"
			strRet = strRet & "	<tr>"
			strRet = strRet & "		<th colspan=5>"& strTitulo & "</th>"
			strRet = strRet & "	</tr>"
			strRet = strRet & "	<tr class=clsSilver2>"
			strRet = strRet & "		<td>&nbsp;Cabo</td>"
			strRet = strRet & "		<td>&nbsp;Par</td>"
			strRet = strRet & "		<td nowrap>&nbsp;PADE</td>"
			strRet = strRet & "		<td>&nbsp;Derivação</td>"
			strRet = strRet & "		<td nowrap>&nbsp;Tipo do Cabo</td>"
			strRet = strRet & "	</tr>"
			strRet = strRet & "<tr class=clsSilver>"
			strRet = strRet & "<td><input type=text class=text name=txtTronco maxlength=2 size=2 onKeyUp=""parent.ValidarTipo(this,0)""  TIPO=N >(N2)</td>"
			strRet = strRet & "<td><input type=text class=text name=txtPar 	maxlength=9 size=9 onKeyUp=""parent.ValidarNTipo(this,0,4,4,1,0,4)"">(N4-N4)</td>"
			strRet = strRet & "<td><input type=text class=text name=txtCaixaEmenda onKeyUp=""parent.ValidarTipo(this,2)""  maxlength=3 size=5 >(A3)</td>"
			strRet = strRet & "<td><input type=text class=text name=txtLateral  onKeyUp=""parent.ValidarTipo(this,2)""	maxlength=4 size=5>(A4)</td>"
			strRet = strRet & "<td><select name=cboTipoCabo >"
			strRet = strRet & "			<option value=""""></option>"
			strRet = strRet & "			<option value=H >H</option>"
			strRet = strRet & "			<option value=X >X</option>"
			'strRet = strRet & "			<option value=XDSL >XDSL</option>"
			strRet = strRet & "			<option value=C >C</option>"
			strRet = strRet & "		</select></td>"
			strRet = strRet & "</tr>"

			strRet = strRet & "</table>"

	End Select

	strRet = strRet & "<table align=center rules=groups cellspacing=0 cellpadding=0 bordercolorlight=#003388 bordercolordark=#ffffff width=540 >"
	strRet = strRet & "<tr class=clsSilver>"
	strRet = strRet & "	<td width=100x >Nº de Registros</td>"
	strRet = strRet & "	<td ><input type=text name=txtQtdeRegistros size=5 class=text value=10 onKeyUp=""ValidarTipo(this,0)"" ></td>"
	strRet = strRet & "	<td align=right>"
	strRet = strRet & "	<input type=button class=button name=btnProcurar value=Procurar onclick=""ProcurarPosicoes(" & Request("cboSistema")  & ")"" accesskey=""P"" onmouseover=""showtip(this,event,'Procurar (Alt+P)');"">&nbsp;"
	strRet = strRet & "	<input type=button class=button name=btnSair value=Sair onclick=""Sair()""  accesskey=""X"" onmouseover=""showtip(this,event,'Sair (Alt+X)');"">&nbsp;"
	strRet = strRet & "	</td>"
	strRet = strRet & "</tr>"
	strRet = strRet & "</table>"
	
	ResgatarCamposFacilidade = strRet

End Function

Function ValidarRecurso(intRede)

	Vetor_Campos(1)="adInterger,2,adParamInput," & request("cboLocalInstala")
	Vetor_Campos(2)="adInterger,2,adParamInput," & request("cboDistLocalInstala")
	Vetor_Campos(3)="adInterger,2,adParamInput," & request("cboProvedor")
	Vetor_Campos(4)="adInterger,2,adParamInput," & request("cboSistema")
	Vetor_Campos(5)="adInterger,2,adParamInput," & request("cboPlataforma")
	Vetor_Campos(6)="adInteger,2,adParamOutput,0"

	strSqlRet = APENDA_PARAMSTR("CLA_sp_check_recurso2",6,Vetor_Campos)
	Set objRS = db.Execute(strSqlRet)
	DBAction = objRS("ret")
	
	

	if DBAction = 0 then
		if Not objRS.Eof And Not objRS.Bof then
			dblRecId = objRS("Rec_id")
		End if	
	Else	
		Response.Write "<script language=javascript>alert('Recurso não encontrado!');window.close();</script>"
		Response.End 
	End if
End Function


Select Case Trim(Request("hdnAcao"))
	Case "Posicoes"
		Dim dblRecId
		ValidarRecurso(Request("cboSistema"))
		strRede = ResgatarCamposFacilidade(Request("cboSistema"))

	Case "ConsultarFacilidade"
		
End Select
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>CLA - Controle Local de Acesso</TITLE>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<SCRIPT language=javascript>
<!--
function ProcurarPosicoes(intTipoRede)
{
	with (document.forms[0])
	{
		switch (parseInt("0" +intTipoRede))
		{
			case 1:
				if (txtFila.value == "" && txtBastidor.value == "" && txtRegua.value == "" && txtPosicao.value == "" && txtTimeSlot.value == "" && txtDominio.value == "" && txtNO.value == "")
				{
					alert("Preencha pelo menos um campo!")
					return
				}
				break
			case 2:
				if (txtTronco.value == "" && txtPar.value == "")
				{
					alert("Preencha pelo menos um campo!")
					return
				}
				break
			case 3:
				if (txtTronco.value == "" && txtPar.value == "" && cboTipoCabo.value == "" && txtLateral.value == "" && txtCaixaEmenda.value == "")
				{
					alert("Preencha pelo menos um campo!")
					return
				}
				break
		}

		if (txtQtdeRegistros.value != "")
		{
			if (txtQtdeRegistros.value > 100){alert("Quantidade de registros maior que o permitido 100!");return}
		}
		else
		{
			txtQtdeRegistros.value = 10
		}	

		if (hdnStsFac.value == "O")
		{
			hdnAcao.value = "ResgatarPosicoesOcupadas"
		}
		else
		{
			hdnAcao.value = "ResgatarPosicoesLivres"
		}	
		hdnRede.value = intTipoRede
		target = "IFrmProcesso"
		action = "ProcessoConsFac.asp"
		submit()
	}
}

function Sair()
{
	window.returnValue = objAryFac
	window.close()
}

function VerificarFacExistente(strChave)
{
	var blnRet = false 
	var intRede = document.forms[0].hdnRede.value

	switch (intRede)
	{ 
		case "1": //Deterministico
			for (var intIndex=0;intIndex<objAryFac.length;intIndex++)
			{
				if (strChave == objAryFac[intIndex][3])
				{
					blnRet = true
				}
			}
			break
		
		default: //Ndet e ADE
			for (var intIndex=0;intIndex<objAryFac.length;intIndex++)
			{
				if (strChave == objAryFac[intIndex][1])
				{
					blnRet = true
				}
			}
	}
	return blnRet
}

function VerificarOcupacao()
{
	var intRede = document.forms[0].hdnRede.value
	var intTotalFac = 0

	for (var intIndex=0;intIndex<objAryFac.length;intIndex++)
	{
		if (objAryFac[intIndex][0] != '')
		{
			intTotalFac += 1
		}
	}
	if (parseInt(intTotalFac) >= 4)
	{
		return true
	}
	else
	{
		return false
	}
}

function DetalharSolicitacao(dblSolId)
{
	//Enviar para pagina de Detalhamento do pedido
	objAryFac[0][0] = "detalhar"
	objAryFac[0][1] = 	dblSolId
	window.close() 
}

var objAryFac = window.dialogArguments

function copyPaste() 
{
	var indice = 0;
	var intRede = document.forms[0].hdnRede.value

	switch (parseInt(intRede))
	{ 
		case 1: //Deterministico
 			
			for (var intIndex=0;intIndex<objAryFac.length;intIndex++)
			{
				objAryFac[intIndex][0] = arguments[0];
				objAryFac[intIndex][1] = arguments[1];
				objAryFac[intIndex][2] = arguments[2];
				objAryFac[intIndex][3] = arguments[3];
				objAryFac[intIndex][4] = arguments[4];
				alert("Facilidade preenchida com sucesso!")
				return
			}
			break
		case 2: //Não deterministico
 			
			for (var intIndex=0;intIndex<objAryFac.length;intIndex++)
			{
				objAryFac[intIndex][0] = arguments[0];
				objAryFac[intIndex][1] = arguments[1];
				alert("Facilidade preenchida com sucesso!")
				return
			}
			break

		case 3: //ADE

			for (var intIndex=0;intIndex<objAryFac.length;intIndex++)
			{
				objAryFac[intIndex][0] = arguments[0];
				objAryFac[intIndex][1] = arguments[1];
				objAryFac[intIndex][2] = arguments[2];
				objAryFac[intIndex][3] = arguments[3];
				objAryFac[intIndex][4] = arguments[4];
				objAryFac[intIndex][5] = arguments[5];
				alert("Facilidade preenchida com sucesso!")
				window.close()
				return
				
			}
			break
	}
}
   
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")

function DetalharFacilidade(intFacId){
	var objNode = objXmlGeral.selectNodes("//Facilidade[@Fac_Id="+intFacId+"]")
	if (objNode.length>0){
		var intRet = window.showModalDialog('MessageConsFac.asp',objNode,'dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
	}
}
//-->
</SCRIPT>
</HEAD>
<BODY >
<Form name=Form1 method=Post>
<input type=hidden name=hdnRede value="<%=request("cboSistema")%>">
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnRecId		value="<%=dblRecId%>">
<input type=hidden name=hdnStsFac		value="<%=Request("strStsFac")%>">

<input type=hidden name=cboLocalInstala		value="<%=request("cboLocalInstala")%>">
<input type=hidden name=cboDistLocalInstala	value="<%=request("cboDistLocalInstala")%>">
<input type=hidden name=cboProvedor			value="<%=request("cboProvedor")%>">
<input type=hidden name=cboSistema			value="<%=request("cboSistema")%>">
<input type=hidden name=hdnPlataforma       value="<%=request("cboPlataforma")%>">
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnXls>
<input type=hidden name=hdnNomeCons value="ConsFac">
<%
if strStsFac = "O" then
	Response.Write "<input type=hidden name=rdoStatusFac value=1>"
Else
	Response.Write "<input type=hidden name=rdoStatusFac value=0>"
End if
Response.Write strRede
%>
<table width="500" align=center border=0>
	<tr>
		<td align=center><span id=spnPosicoes></span></tr>
	</tr>
</table>
</Form>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "0" 
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
</BODY>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub window_onunload
	window.returnValue = objAryFac
End Sub
-->
</SCRIPT>
</HTML>