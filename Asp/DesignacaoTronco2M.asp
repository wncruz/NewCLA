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
Dim dblDesig_id


dblAcf_IDAcessoFisico = Request.QueryString("Acf_IDAcessoFisico")
dblDesig_id = Request.QueryString("desig_id")
dblacao = Request.QueryString("acao")

'response.write dblAcf_IDAcessoFisico
'response.write dblDesig_id
'response.write dblacao
'response.end


 
if Trim(dblAcf_IDAcessoFisico) = "" then
	 dblAcf_IDAcessoFisico  = Request.Form("hdnAcf_IDAcessoFisico")
End if	 

If request("action")="Gravar" and dblacao <> "registro" then


	Designacao = Espaco(Request.Form("txtCNL1"),4) + Espaco(Request.Form("txtPredio1"),3) + Espaco(Request.Form("txtComplemento1"),3) + Espaco(Request.Form("txtCNL2"),4)  
	Designacao = Designacao + Espaco(Request.Form("txtPredio2"),3) + Espaco(Request.Form("txtComplemento2"),3) + Espaco(Request.Form("txtCodFuncao"),6) + Espaco(Request.Form("txtSequencial"),3)
		
	'response.write Designacao
	Set Tronco = CreateObject("EOL.CFMYS") 
		
	
	Tronco.ServerAddress = "ETBPRDBA.NT.EMBRATEL.COM.BR:1971@RPC/SRVPCFM/CALLNAT" '10.4.2.1:5026@RPC/SRVPCLE/CALLNAT' //Prod  //Desenv 10.2.4.83:5024@RPC/SRVCLE/CALLNAT
	Tronco.Logon
		
	Tronco.CFMYS193 Designacao, retOTS , retRC , retMensagem  
	
	'response.write retOTS
	'response.write retrc
	'response.write retMensagem
	'response.end
	'Deve ser considerada como resposta positiva, de existência da designação,
	'quando uma das respostas abaixo ocorrer:

	'RC = 00 ou
	'RC = 02 e Mensagem diferente de 'CIRCUITO JA DESATIVADO NO CFM'  ou
	'RC = 02 e mensagem diferente de 'CIRCUITO EM CANCELAMENTO DA ATIVACAO NO CFM ou
	'RC = 03
	
	if ( ( retRC = 0 ) or (retRC = 2 and retMensagem <> "CIRCUITO JA DESATIVADO NO CFM") or (retRC = 2 and retMensagem <> "CIRCUITO EM CANCELAMENTO DA ATIVACAO NO CFM")  or (retRC = 3) ) then
	

		Vetor_Campos(1)="adInteger,2,adParamInput," & Request.Form("hdndesig_id") 
		Vetor_Campos(2)="adInteger,2,adParamInput," & Request.Form("hdnAcf_id") 
		Vetor_Campos(3)="adWChar,30,adParamInput,"& Designacao
		Vetor_Campos(4)="adInteger,3,adParamInput, 2"  '1- inclusão , 2 - alteração
		Vetor_Campos(5)="adInteger,3,adParamOutput, null" 
			
		Call APENDA_PARAM("CLA_sp_ins_Designacao",5,Vetor_Campos)
		ObjCmd.Execute'pega dbaction
		DBAction = ObjCmd.Parameters("RET").value
		if 	DBAction = 1 then
			response.write "<script>alert('Gravado com Sucesso')</script>"
		end if
		
	else
		response.write "<script>alert('Designação inválida no CFM')</script>"
	
	end if 
	
	Set Tronco = Nothing

End if

dblacao = ""

if dblAcf_IDAcessoFisico <> "" then

	Vetor_Campos(1)="adWChar,15,adParamInput,"	& dblAcf_IDAcessoFisico	
	Vetor_Campos(2)="adInteger,8,adParamInput,null" 
	Vetor_Campos(3)="adInteger,8,adParamInput,null" 
	Vetor_Campos(4)="adWChar,15,adParamInput,null"				
	Vetor_Campos(5)="adInteger,8,adParamInput, " & dblDesig_id
	
	strSql = APENDA_PARAMSTR("CLA_sp_sel_designacaoTronco2m",5,Vetor_Campos)
	
	set objRS = db.execute(strSql)
	If Not objRS.eof or Not objRS.bof Then
	
		if (objRS("desig_designacao") <> "" ) then
		
			strCNL1 					=	left(TratarAspasHtml(Trim(objRS("desig_designacao"))),4)
			strAcf_IDAcessoFisicoAux	=	right(TratarAspasHtml(Trim(objRS("desig_designacao"))),25)
			strPredio1					=	left(strAcf_IDAcessoFisicoAux,3)
			strAcf_IDAcessoFisicoAux	=	right(strAcf_IDAcessoFisicoAux,22)
			strComplemento1				=	left(strAcf_IDAcessoFisicoAux,3)
			strAcf_IDAcessoFisicoAux	=	right(strAcf_IDAcessoFisicoAux,19)
			strCNL2						=	left(strAcf_IDAcessoFisicoAux,4)
			strAcf_IDAcessoFisicoAux	=	right(strAcf_IDAcessoFisicoAux,15)
			strPredio2					=	left(strAcf_IDAcessoFisicoAux,3)
			strAcf_IDAcessoFisicoAux	=	right(strAcf_IDAcessoFisicoAux,12)
			strComplemento2				=	left(strAcf_IDAcessoFisicoAux,3)
			strAcf_IDAcessoFisicoAux	=	right(strAcf_IDAcessoFisicoAux,9)
			strCodFuncao				=	left(strAcf_IDAcessoFisicoAux,6)
			strAcf_IDAcessoFisicoAux	=	right(strAcf_IDAcessoFisicoAux,3)
			strSequencial				=	strAcf_IDAcessoFisicoAux
			
			StrDesig_Id = TratarAspasHtml(Trim(objRS("desig_id")))
			
		end if
		StrAcf_Id = TratarAspasHtml(Trim(objRS("acf_id")))
				
	end if
end if
%>
<!--#include file="../inc/header.asp"-->
<form action="DesignacaoTronco2M.asp" method="post" >
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnAcf_IDAcessoFisico value="<%=dblAcf_IDAcessoFisico%>" >
<input type=hidden name=hdndesig_id value="<%=StrDesig_Id%>" >
<input type=hidden name=hdnAcf_id value="<%=StrAcf_Id%>" >


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
				<input type="submit" class="button" name="action" value="Gravar" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
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

