<!--#include file="../inc/data.asp"-->

<%
Dim dblAcf_IDAcessoFisico

Dim strCNL1
Dim strPredio1
Dim strComplemento1
Dim strCNL2
Dim strPredio2
Dim strComplemento2
Dim strCodFuncao
Dim strSequencial
Dim strAcf_IDAcessoFisicoAux
Dim StrDesig_Id
Dim StrAcf_Id
Dim Designacao
Dim retOTS
Dim retRC
Dim retMensagem
Dim Tronco

Dim dblacao

dblAcf_IDAcessoFisico = Request.Form("hdnAcf_IDAcessoFisico")
dblAcf_ID = Request.Form("hdnAcfID")
dblacao = Request.QueryString("acao")

 
if Trim(dblAcf_IDAcessoFisico) = "" then
	 dblAcf_IDAcessoFisico  = Request.Form("hdnAcf_IDAcessoFisico")
End if	 

If request("hdnAcao")="Gravar" then

	strCNL1 					=	Espaco(Request.Form("txtCNL1"),4)
	strPredio1					=	Espaco(Request.Form("txtPredio1"),3)
	strComplemento1				=	Espaco(Request.Form("txtComplemento1"),3)
	strCNL2						=	Espaco(Request.Form("txtCNL2"),4)
	strPredio2					=	Espaco(Request.Form("txtPredio2"),3)
	strComplemento2				=	Espaco(Request.Form("txtComplemento2"),3)
	strCodFuncao				=	Espaco(Request.Form("txtCodFuncao"),6)
	strSequencial				=	Espaco(Request.Form("txtSequencial"),3)
	
	'Designacao = Espaco(Request.Form("txtCNL1"),4) + Espaco(Request.Form("txtPredio1"),3) + Espaco(Request.Form("txtComplemento1"),3) + Espaco(Request.Form("txtCNL2"),4)  
	'Designacao = Designacao + Espaco(Request.Form("txtPredio2"),3) + Espaco(Request.Form("txtComplemento2"),3) + Espaco(Request.Form("txtCodFuncao"),6) + Espaco(Request.Form("txtSequencial"),3)
	
	Designacao = strCNL1 + strPredio1 + strComplemento1 + strCNL2 + strPredio2 + strComplemento2 + strCodFuncao + strSequencial
	
	Set Tronco = CreateObject("EOL.CFMYS")	
	Tronco.ServerAddress = "ETBPRDBA.NT.EMBRATEL.COM.BR:1971@RPC/SRVPCFM/CALLNAT" '10.4.2.1:5026@RPC/SRVPCLE/CALLNAT' //Prod  //Desenv 10.2.4.83:5024@RPC/SRVCLE/CALLNAT
	Tronco.Logon
	Tronco.CFMYS193 Designacao, retOTS , retRC , retMensagem	
	
	'Deve ser considerada como resposta positiva, de existência da designação,
	'quando uma das respostas abaixo ocorrer:

	'RC = 00 ou
	'RC = 02 e Mensagem diferente de 'CIRCUITO JA DESATIVADO NO CFM'  ou
	'RC = 02 e mensagem diferente de 'CIRCUITO EM CANCELAMENTO DA ATIVACAO NO CFM ou
	'RC = 03
		
	if (  ( retRC = 0 ) or (retRC = 2 and retMensagem <> "CIRCUITO JA DESATIVADO NO CFM") or (retRC = 2 and retMensagem <> "CIRCUITO EM CANCELAMENTO DA ATIVACAO NO CFM")  or (retRC = 3) ) then	
		
		Vetor_Campos(1)="adWChar,15,adParamInput,"	& dblAcf_IDAcessoFisico 'Request.Form("hdnAcfID")	
		Vetor_Campos(2)="adInteger,2,adParamInput,null" 
		Vetor_Campos(3)="adInteger,2,adParamInput,null" 
		Vetor_Campos(4)="adWChar,15,adParamInput,null"				
		
		strSql = APENDA_PARAMSTR("CLA_sp_sel_designacaoTronco2m",4,Vetor_Campos)
		
		set objRS = db.execute(strSql)
		If Not objRS.eof or Not objRS.bof Then
		
			StrAcf_Id = TratarAspasHtml(Trim(objRS("acf_id")))
		end if
		
		Vetor_Campos(1)="adInteger,2,adParamInput, "  
		Vetor_Campos(2)="adInteger,2,adParamInput," & StrAcf_Id 'Request.Form("hdnAcfId") 
		Vetor_Campos(3)="adWChar,30,adParamInput,"& Designacao
		Vetor_Campos(4)="adInteger,3,adParamInput, 1 " ' 1 - inclusão , 2 - alteração 
		Vetor_Campos(5)="adInteger,3,adParamOutput, null" 
			
		Call APENDA_PARAM("CLA_sp_ins_Designacao",5,Vetor_Campos)
				
		ObjCmd.Execute'pega dbaction
		DBAction = ObjCmd.Parameters("RET").value
		if 	DBAction = 1 then
			response.write "<script>alert('Incluído com Sucesso')</script>"
		end if
	else
		if trim(retMensagem) = "" then
		  msg_aux = "<script>alert('"&retRC&" - Designação inválida no CFM')</script>"
		else
		  msg_aux = "<script>alert('"&retRC&" - Designação inválida no CFM ("&trim(retMensagem)&")')</script>"
		end if
		
		response.write msg_aux
	end if 

End if

dblacao = ""


%>
<!--#include file="../inc/header.asp"-->
<form action="IncluirDesignacaoTronco2M.asp" method="post" >
<SCRIPT LANGUAGE="JavaScript">


function incluir()
{
	with (document.forms[0])
	{
		if (txtCNL1.value != "" || txtPredio1.value != "" || txtCNL2.value != "" || txtPredio2.value != "" || txtCodFuncao.value != "" || txtSequencial.value != "")
		{
			hdnAcao.value = "Gravar"
			hdnAcf_id.value = txtFisico.value
			submit()
		}
		else
		{
			alert("Informe os campos obrigatórios :: ")
			txtCNL1.focus()
			return
		}		
	}
}


</script>


<input type=hidden name=hdnAcao>
<input type=hidden name=hdnAcf_IDAcessoFisico value="<%=dblAcf_IDAcessoFisico%>" >
<input type=hidden name=hdnAcf_id value="<%=StrAcf_Id%>" >
<input type=hidden name=txtFisico value="<%=Server.HTMLEncode(Request.Form("txtFisico"))%>">
<input type=hidden name=txtSolID value="<%=Server.HTMLEncode(Request.Form("txtSolID"))%>">
<input type=hidden name=txtLogico value="<%=Server.HTMLEncode(Request.Form("txtLogico"))%>">


<tr>
	<td >
		<table border="0" cellspacing="1" cellpadding=0 width="760">
			<tr>
				<th colspan=2><p align=center>Associação de Acesso Físico com Designação de Tronco 2M</p></td>
			</tr>

			<tr class=clsSilver>
				<td><font class="clsObrig">:: </font>ID Acesso Físico</td>
				<td><input type="text" class="text" name="txtFisico" value="<%=dblAcf_IDAcessoFisico%>" maxlength="60" size="40" disabled=true></td>
			</tr>
			<table border="0" cellspacing="1" cellpadding=0 width="760">
				<tr>
					<td>
						<tr class=clsSilver>
							<th colspan=8><p align=center>Designação Tronco 2M</p></td>
						</tr>
						<tr class=clsSilver>
							<td><font class="clsObrig">:: </font>CNL(A4)</td>
							<td><font class="clsObrig">:: </font>Prédio(A3)</td>
							<td>Complemento(A3)</td>
							<td><font class="clsObrig">:: </font>CNL(A4)</td>
							<td><font class="clsObrig">:: </font>Prédio(A3)</td>
							<td>Complemento(A3)</td>
							<td><font class="clsObrig">:: </font>Código de Função(A6)</td>
							<td><font class="clsObrig">:: </font>Sequencial(N3)</td>
						</tr>
						
						
					   	<tr class=clsSilver>
					 		<td><input type="text" class="text" name="txtCNL1" value="<%=strCNL1%>" maxlength="4" size="10" ></td>
							<td><input type="text" class="text" name="txtPredio1" value="<%=strPredio1%>" maxlength="3" size="10" ></td>
							<td><input type="text" class="text" name="txtComplemento1" value="<%=strComplemento1%>" maxlength="3" size="10" ></td>
							<td><input type="text" class="text" name="txtCNL2" value="<%=strCNL2%>" maxlength="4" size="10" ></td>
							<td><input type="text" class="text" name="txtPredio2" value="<%=strPredio2%>" maxlength="3" size="10" ></td>
							<td><input type="text" class="text" name="txtComplemento2" value="<%=strComplemento2%>" maxlength="3" size="10"></td>
							<td><input type="text" class="text" name="txtCodFuncao" value="<%=strCodFuncao%>" maxlength="6" size="10" ></td>
							<td><input type="text" class="text" name="txtSequencial" value="<%=strSequencial%>" maxlength="3" size="10" onKeyUp="ValidarTipo(this,0)"></td>
						</tr>
																		
					</td>
				</tr>
			</table>

		</table>
		<table width="760" border=0>
		<tr>
			<td colspan=2 align="center"><br>
				<input type="submit" class="button" name="Incluir" value="Incluir" onclick="incluir()" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
				<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="LimparForm();setarFocus('txtCNL1');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
				<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('DesignacaoTronco2M_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
				<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
			</td>
		</tr>
		</table>
		<table width="760" border=0>
		<tr>
			<td>
				<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
			</td>
		</tr>
		<tr>
			<td>
				<b> Formato Designação de Tronco 2M: Estação 1 + Estação 2 + Código de Função + Sequencial </b>
			</td>
		</tr>
		<tr>
			<td>
				<b> CNL(A4) + Prédio(A3) + Complemento(A3) + Código de Função(A6) + Sequencial(N3) </b>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</body>

</html>
<%
Set objRS = Nothing
DesconectarCla()
%>

