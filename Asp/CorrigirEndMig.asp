<%
'	- Sistema	: CLA
'	- Arquivo	: CorrigirEndMig.asp
'	- Descrição	: Corrige um endereço migrado
%>

<!--#include file="../inc/data.asp"-->

<%
REM Dados do endereço migrado
Dim intIdEndMig
Dim intIdCompMig
Dim strNomeFantasiaMig
Dim strContaCorrenteMig
Dim strIdLogicoMigradoMig
Dim strIdFisicoMigradoMig
Dim	strDesignacaoServicoMig
Dim strLogradouroMig
Dim strNumeroMig
Dim strBairroMig
Dim strCidadeMig
Dim strCEPMig
Dim	strUFMig
Dim strComplementoMig
Dim strErroMig
Dim DM

REM Dados do endereço corrigido
Dim dblCepID
Dim strUf			
Dim strLocalidade	
Dim strCep			
Dim strBairro		
Dim strCepConsulta
Dim strNumero
Dim strComplemento
Dim strRuaCompleta
Dim strCidade
Dim strCNL
Dim strTipoLogr
Dim strNomeLogr
Dim strPreposicao
Dim strTitulo
Dim strRua
Dim FiltroDM

REM Obtem os argumentos enviados pela página ListarEndMig.asp -------------
intIdEndMig  = request("EndId")
if trim(intIdEndMig) = "" then intIdEndMig = request.form("hdnIdEndMig")

intIdCompMig = request("AecId")
if trim(intIdCompMig) = "" then intIdCompMig = request.form("hdnIdCompMig")
REM -----------------------------------------------------------------------

REM Obtém o CEP escolhido -------------------------------------------------
dblCepID = request("ID")
if Trim(dblCepID) = "" then dblCepID = Request.Form("hdnId") 
REM -----------------------------------------------------------------------


REM Obtém o FiltroDM -------------------------------------------------
FiltroDM = Request.Form("hdnFiltroDM")
REM -----------------------------------------------------------------------


REM Monta uma combo de endereços com base no CEP escolhido ----------------
strCboRet = ""

REM Recuperando os dados do endereço cadastrado na base dos coreios por esta página.
if Request.Form("EndGravadoMig") = "OK" then
  strCepConsulta = request.Form("txtCep")
end if
REM -----------------------------------------------------------------------

REM Função ProcurarCEP() - Retorna se o cep foi encontrado ou não. --------
REM Liberar a inclusão de endereços por esta página caso o cep não seja encontrado.

if Trim(Request.Form("hdnAcao"))= "ProcurarCEP" then
  set objRS = db.execute("CLA_SP_VIEW_CEP '" & Request.Form("hdnCEP") & "'")
  strCepConsulta = Request.Form("hdnCEP")
  strUfMig = TratarAspasHtml(Trim(Request.Form("txtUFMig")))
  If Not objRS.eof and Not objRS.bof then
    if objRS("Est_Sigla") <> strUfMig then
	  Response.Write "<script language=javascript>alert('Não é permitido um CEP cuja UF seja diferente da atual!')</script>"
	else
	  strCboRet = "<td>:: Endereços</td><td><Select name=cboCEPS onchange = 'ResgatarCEP(this)'>"
	  strCboRet = strCboRet & "<Option value="""">SELECIONE UM ENDEREÇO</Option>"
	  While Not objRS.Eof
	    strCboRet = strCboRet & "<Option value=" & Trim(objRS("Cep_ID")) & ">" & TratarAspasJS(Trim(objRS("RuaCompleta"))) & " - " & TratarAspasJS(Trim(objRS("Cep"))) & "</Option>"
		objRS.MoveNext
	  Wend
		strCboRet = strCboRet & "</Select> <br>(Ou preencha os campos e grave outro endereço para o mesmo CEP genérico)</td>"
    end if
		
	REM Liberar a inclusão de endereços por esta página caso o cep seja genérico: ----
	'Response.Write "<script>alert('Últimos 3 digitos do CEP: " & mid(strCepConsulta,7,3) & "')</script>"
	if trim(mid(strCepConsulta,7,3)) = "" or mid(strCepConsulta,7,3) = "000" then
	  'Response.Write "<script>alert('CEP Genérico')</script>"
	  var_liberaredicao = true
	else  
	  var_liberaredicao = false
	end if
	REM -----------------------------------------------------------------------
  Else      	
    Response.Write "<script language=javascript>alert('CEP não encontrado.')</script>"
	var_liberaredicao = true
  End if
End if 
REM -----------------------------------------------------------------------

REM GRAVA O ENDEREÇO ------------------------------------------------------
if Trim(Request.Form("hdnAcao")) = "Gravar" then
	Vetor_Campos(1)="adInteger,1,adParamInput,"	&	request.form("hdnIdEndMig")
	Vetor_Campos(2)="adWChar,2,adParamInput," 	&	request.form("txtUF")
	Vetor_Campos(3)="adWChar,4,adParamInput,"	&	request.Form("txtCNL")
	Vetor_Campos(4)="adWChar,15,adParamInput,"	& 	request.Form("hdnTipoLogr")
	Vetor_Campos(5)="adWChar,10,adParamInput,"	& 	request.Form("txtNumero")
	Vetor_Campos(6)="adWChar,60,adParamInput,"	& 	request.Form("hdnNomeLogr")
	Vetor_Campos(7)="adWChar,30,adParamInput,"	& 	request.Form("txtBairro")
	Vetor_Campos(8)="adWChar,9,adParamInput,"	&	request.Form("txtCEP")
	Vetor_Campos(9)="adInteger,1,adParamInput,"	& 	request.Form("hdnIdCompMig")
	Vetor_Campos(10)="adWChar,30,adParamInput,"	& 	request.form("txtComplemento")
	Vetor_Campos(11)="adInteger,1,adParamInput,"	& request.form("txtIdFisicoMigradoMig")
	Vetor_Campos(12)="adWChar,15,adParamInput,"	& 	strUserName
	Vetor_Campos(13)="adInteger,1,adParamOutput,0"
	
	var_habilita_response = false
	if var_habilita_response = true then
	  Response.Write "<b>CLA_sp_corrigir_end_mig</b><br><br>"
	  Response.Write "<font color='blue'>SET </font><b>@End_ID</b>=<font color='red'>'" & request.form("hdnIdEndMig") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Est_Sigla </b>=<font color='red'>'" & request.form("txtUF") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Cid_Sigla </b>=<font color='red'>'" & request.Form("txtCNL") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Tpl_Sigla</b>=<font color='red'>'" & request.Form("hdnTipoLogr") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@End_NroLogr</b>=<font color='red'>'" & request.Form("txtNumero") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@End_NomeLogr </b>=<font color='red'>'" & request.Form("hdnNomeLogr") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@End_Bairro</b>=<font color='red'>'" & request.Form("txtBairro") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@End_CEP</b>=<font color='red'>'" & request.Form("txtCEP") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Aec_ID</b>=<font color='red'>'" & request.Form("hdnIdCompMig") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Aec_Complemento</b>=<font color='red'>'" & request.form("txtComplemento") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Acf_ID</b>=<font color='red'>'" & request.form("txtIdFisicoMigradoMig") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@RespAtualizacao</b>=<font color='red'>'" & strUserName & "'</font><br>"
	  Response.Write "<br><a href='menumig.asp'>Voltar</a>"
	  Response.end
	end if  
	
	Call APENDA_PARAM("CLA_sp_corrigir_end_mig ", 13, Vetor_Campos)
	
 	ObjCmd.Execute 'pega dbaction

	DBAction = ObjCmd.Parameters("RET").value
	

	if DBAction = 0 then
		Response.Write "<script language=javascript> alert('Endereço gravado com sucesso!'); </script>"
	else
		Response.Write "<script language=javascript> alert('Não foi possível gravar o endereço!'); </script>"
	end if

End if
REM -----------------------------------------------------------------------


REM MANIPULA AS INFORMAÇÕES REFERENTES AO ENDEREÇO MIGRADO COM ERRO -------
if intIdEndMig <> "" then
	
	REM Se tiver sido chamada através da ListarEndMig.asp
	
	' Set ObjRsMig = db.execute("CLA_SP_Obter_Dados_End_Acesso_Mig " & intIdEndMig)
	Set ObjRsMig = db.execute("CLA_SP_Obter_Dados_Acesso_Mig " & intIdEndMig)	
	if Not ObjRsMig.Eof and not ObjRsMig.Bof then
		strNomeFantasiaMig 			= TratarAspasHtml(Trim(ObjRsMig("NomeFantasia")))
		strContaCorrenteMig 		= TratarAspasHtml(Trim(ObjRsMig("ContaCorrente")))
		strIdLogicoMigradoMig 		= TratarAspasHtml(Trim(ObjRsMig("IdLogicoMigrado")))
		strIdFisicoMigradoMig 		= TratarAspasHtml(Trim(ObjRsMig("IdFisicoMigrado")))		
		strDesignacaoServicoMig		= TratarAspasHtml(Trim(ObjRsMig("DesignacaoServico")))
		strLogradouroMig 			= TratarAspasHtml(Trim(ObjRsMig("TipoLogradouro"))) + " " + TratarAspasHtml(Trim(ObjRsMig("Logradouro")))
		strNumeroMig 				= TratarAspasHtml(Trim(ObjRsMig("Numero")))
		strBairroMig 				= TratarAspasHtml(Trim(ObjRsMig("Bairro")))
		strCidadeMig 				= TratarAspasHtml(Trim(ObjRsMig("Cidade")))
		strCEPMig 					= TratarAspasHtml(Trim(ObjRsMig("CEP")))
		strUFMig 					= TratarAspasHtml(Trim(ObjRsMig("UF")))
		strBairroMig 				= TratarAspasHtml(Trim(ObjRsMig("Bairro")))
		strComplementoMig 			= TratarAspasHtml(Trim(ObjRsMig("Complemento")))
		strErroMig 					= TratarAspasHtml(Trim(ObjRsMig("Erro")))
		DM 							= TratarAspasHtml(Trim(ObjRsMig("DM")))
	End If

else
	strNomeFantasiaMig 			= TratarAspasHtml(Trim(Request.Form("txtNomeFantasiaMig")))
	strContaCorrenteMig 		= TratarAspasHtml(Trim(Request.Form("txtContaCorrenteMig")))
	strIdLogicoMigradoMig 		= TratarAspasHtml(Trim(Request.Form("txtIdLogicoMigradoMig")))
	strIdFisicoMigradoMig 		= TratarAspasHtml(Trim(Request.Form("txtIdFisicoMigradoMig")))		
	strDesignacaoServicoMig		= TratarAspasHtml(Trim(Request.Form("txtDesignacaoServicoMig")))
	strLogradouroMig 			= TratarAspasHtml(Trim(Request.Form("txtLogradouroMig")))
	strNumeroMig 				= TratarAspasHtml(Trim(Request.Form("txtNumeroMig")))
	strBairroMig 				= TratarAspasHtml(Trim(Request.Form("txtBairroMig")))
	strCidadeMig 				= TratarAspasHtml(Trim(Request.Form("txtCidadeMig")))
	strCEPMig 					= TratarAspasHtml(Trim(Request.Form("txtCEPMig")))
	strUFMig 					= TratarAspasHtml(Trim(Request.Form("txtUFMig")))
	strBairroMig 				= TratarAspasHtml(Trim(Request.Form("txtBairroMig")))
	strComplementoMig 			= TratarAspasHtml(Trim(Request.Form("txtComplementoMig")))
	strErroMig 					= TratarAspasHtml(Trim(Request.Form("txtErroMig")))
	DM 							= TratarAspasHtml(Trim(Request.Form("hdnDM")))
end if



REM MANIPULA AS INFORMAÇÕES REFERENTES AO ENDEREÇO CORRIGIDO
if dblCepID <> "" then

	' SE O USUÁRIO ESCOLHER ALGUM ENDEREÇO A PARTIR DA COMBO
	
	Set objRS = db.execute("CLA_sp_view_cep null," & dblCepID)
	if Not objRS.Eof and not objRS.Bof then
		strCNL			= TratarAspasHtml(Trim(objRS("Cnl")))
		strLocalidade	= TratarAspasHtml(Trim(objRS("Cid_Desc")))
		strCidade		= TratarAspasHtml(Trim(objRS("Cidade")))
		strBairro		= TratarAspasHtml(Trim(objRS("Bairroinicial")))
		strRuaCompleta	= TratarAspasHtml(Trim(objRS("RuaCompleta")))
		strCep			= TratarAspasHtml(Trim(objRS("Cep")))
		strUF			= TratarAspasHtml(Trim(objRS("Est_Sigla")))
		strCidade		= TratarAspasHtml(Trim(objRS("Cidade")))
		strTipoLogr		= TratarAspasHtml(Trim(objRS("Logradouro")))
		strPreposicao	= TratarAspasHtml(Trim(objRS("Preposicao")))
		strTitulo		= TratarAspasHtml(Trim(objRS("Titulo"))) 
		strRua			= TratarAspasHtml(Trim(objRS("Rua")))

		REM Monta o nome do logradouro
		if strPreposicao = "" and strTitulo = "" then
			REM "JOSE MARIANO" (CEP 54730-090)
			strNomeLogr = strRua									
		else
			if strPreposicao <> "" and strTitulo = "" then
				REM "DOS" + " " + "ABACATEIROS" (CEP 13474-372)
				strNomeLogr = strPreposicao + " " + strRua	
			else
				if strPreposicao = "" and strTitulo <> "" then
					REM "PREFEITO" + " " + "ABDO NAJAR" (CEP 13466-615)
					strNomeLogr = strTitulo + " " + strRua	
				else
					REM "NOSSA SENHORA" + " " + "DO" + " " + "BOM CONSELHO" (CEP 05763-915)
					strNomeLogr = strTitulo + " " + strPreposicao + " " + strRua 	
				end if
			end if
		end if

	End if
Else
	strCNL			= TratarAspasHtml(Trim(Request.Form("txtCnl")))
	strLocalidade	= TratarAspasHtml(Trim(Request.Form("txtLocalidade")))
	strCidade		= TratarAspasHtml(Trim(Request.Form("txtCidade")))
	strBairro		= TratarAspasHtml(Trim(Request.Form("txtBairro")))
	strRuaCompleta	= TratarAspasHtml(Trim(Request.Form("txRuaCompleta")))
	strCep			= TratarAspasHtml(Trim(Request.Form("txtCep")))
	strUF			= TratarAspasHtml(Trim(Request.Form("txtUf")))
	strCidade		= TratarAspasHtml(Trim(Request.Form("txtCidade")))
	strTipoLogr		= TratarAspasHtml(Trim(Request.Form("hdnTipoLogr")))
	strNomeLogr		= TratarAspasHtml(Trim(Request.Form("hdnNomeLogr")))
End if

strNumero			= TratarAspasHtml(Trim(Request.Form("txtNumero")))	
strComplemento 		= TratarAspasHtml(Trim(Request.Form("txtComplemento")))

if (trim(mid(strCepConsulta,7,3)) = "" AND Trim(Request.Form("hdnAcao"))= "ProcurarCEP") or (mid(strCepConsulta,7,3) = "000" AND Trim(Request.Form("hdnAcao"))= "ProcurarCEP") then
  'Response.Write "<script>alert('CEP Genérico')</script>"
  strRuaCompleta = ""
end if
%>

<!--#include file="../inc/headermig.asp"-->

<form action="CorrigirEndMig.asp" method="post" name="Form_1">
<input type=hidden name=hdnId value="<%=dblCepID%>" >

<script language="VBScript" src="../VBScript/estadocep.vbs"></script>
<script language=VBScript>
Function recuperaestadocep()
 cep = Form_1.txtCepConsulta.value
 estadovalida = Form_1.txtUFMig.value
 'alert estadovalida 
 if estadocep(cep,estadovalida) then
    Procurarcep
 end if 
End function

Sub btnIncluir_OnClick()
  cep = Form_1.txtCepConsulta.value
  estadovalida = Form_1.txtUFMig.value
  'alert estadovalida 
  if estadocep(cep,estadovalida) then
    GravarEnd
  end if
End sub
</script>

<SCRIPT LANGUAGE="JavaScript">
function ResgatarCEP(objIDCEP)
{
	with (document.forms[0])
	{
		hdnId.value = objIDCEP.value
		target = self.name 
		action = "CorrigirEndMig.asp"
		submit()
	}
}

function GravarInfo()
{
	with (document.forms[0])
	{
		if (txtCEP.value == "" || txtBairro.value == "" || txtNumero.value == "")
		{
			alert("Os campos CEP, bairro e número são obrigatórios!");
			return false;
		}	
		hdnAcao.value = "Gravar"
		target = self.name 
		action = "CorrigirEndMig.asp"
		submit()
	}
}

function ProcurarCEP()
{
	//document.forms[0].spnCeps.innerHTML = ''
	with (document.forms[0])
	{
		try {
				spnCeps.innerHTML = ''
			}
		catch(e){
				alert(e.description)
		}
			
		hdnAcao.value = "ProcurarCEP"
		hdnCEP.value = txtCepConsulta.value
		if (hdnCEP.value.length < 5 )
		{
			alert("O CEP deve ter, pelo menos, cinco números!")
			return
		}
		target = self.name 
		action = "CorrigirEndMig.asp"
		submit()
	}	
}

function GravarEnd()
{

	with (document.forms[0])
	{
		if (txtUFMig.value == "")
		{
			alert("Verifique campo obrigatório (UF).");
			txtUFMig.focus();
			return false;
		}
		if (txtCNL.value == "")
		{
			alert("Verifique campo obrigatório (CNL).");
			txtCNL.focus();
			return false;
		}
		if (cboTipoLogr.value == "")
		{
			alert("Verifique campo obrigatório (Tipo Logradouro).");
			cboTipoLogr.focus();
			return false;
		}
		if (txtRuaCompleta.value == "")
		{
			alert("Verifique campo obrigatório (Nome Logradouro).");
			txtRuaCompleta.focus();
			return false;
		}
		if (txtCEP.value == "")
		{
			alert("Verifique campo obrigatório (CEP).");
			txtCEP.focus();
			return false;
		}
		if (txtBairro.value == "")
		{
			alert("Verifique campo obrigatório (Bairro).");
			txtBairro.focus();
			return false;
		}
		
		hdnAcao.value = "Gravar_mig"
		target = self.name 
		action = "CorrigirEndMig_CadEnd.asp?EndId=<%=intIdEndMig%>&AecId=<%=intIdCompMig%>"
		submit()
	}
}
</script>

<input type=hidden name=hdnAcao>
<input type=hidden name=hdnCEP>
<input type=hidden name=hdnDM value="<%=DM%>">
<input type=hidden name=hdnFiltroDM value="<%=FiltroDM%>">


<input type=hidden name=hdnIdEndMig value="<%=intIdEndMig%>">
<input type=hidden name=hdnIdCompMig value="<%=intIdCompMig%>">

<input type=hidden name=hdnTipoLogr value="<%=strTipoLogr%>">
<input type=hidden name=hdnNomeLogr value="<%=strNomeLogr%>">

<tr><td>
<table border=0 cellspacing="1" cellpadding="0" width="760">
<tr>
	<th colspan=2><p align="center">Correção de Endereço</p></th>
</tr>

<!-- Exibe o endereço migrado, juntamente com os dados do cliente e acesso -->
<tr class=clsSilver>
<td>Cliente</td>
<td><input type="text" class="text" name="txtNomeFantasiaMig"  maxlength="40" size="46" readonly value="<%=strNomeFantasiaMig%>"></td>
</tr>

<tr class=clsSilver>
<td>Conta-Corrente</td>
<td><input type="text" class="text" name="txtContaCorrenteMig"  maxlength="40" size="46" readonly value="<%=strContaCorrenteMig%>"></td>
</tr>

</tr>
<tr class=clsSilver>
<td>Designação do Serviço</td>
<td><input type="text" class="text" name="txtDesignacaoServicoMig" maxlength="40" size="46" readonly value="<%=strDesignacaoServicoMig%>"></td>
</tr>

</tr>
<tr class=clsSilver>
<td>Designação do Acesso Lógico</td>
<td><input type="text" class="text" name="txtIdLogicoMigradoMig" maxlength="40" size="46" readonly value="<%=strIdLogicoMigradoMig%>"></td>
</tr>

</tr>
<tr class=clsSilver>
<td>Designação do Acesso Físico</td>
<td><input type="text" class="text" name="txtIdFisicoMigradoMig" maxlength="40" size="46" readonly value="<%=strIdFisicoMigradoMig%>"></td>
</tr>

</tr>
<tr class=clsSilver>
<td>Logradouro</td>
<td><input type="text" class="text" name="txtLogradouroMig" maxlength="40" size="46" readonly value="<%=strLogradouroMig%>"></td>
</tr>

</tr>
<tr class=clsSilver>
<td>Número</td>
<td><input type="text" class="text" name="txtNumeroMig" maxlength="40" size="46" readonly value="<%=strNumeroMig%>"></td>
</tr>

</tr>
<tr class=clsSilver>
<td>Complemento</td>
<td><input type="text" class="text" name="txtComplementoMig" maxlength="40" size="46" readonly value="<%=strComplementoMig%>"></td>
</tr>

</tr>
<tr class=clsSilver>
<td>Bairro</td>
<td><input type="text" class="text" name="txtBairroMig" maxlength="40" size="46" readonly value="<%=strBairroMig%>"></td>
</tr>

</tr>
<tr class=clsSilver>
<td>Cidade</td>
<td><input type="text" class="text" name="txtCidadeMig" maxlength="40" size="46" readonly value="<%=strCidadeMig%>"></td>
</tr>

</tr>
<tr class=clsSilver>
<td>UF</td>
<td><input type="text" class="text" name="txtUFMig" maxlength="40" size="46" readonly value="<%=strUFMig%>"></td>
</tr>

</tr>
<tr class=clsSilver>
<td>CEP</td>
<td><input type="text" class="text" name="txtCEPMig" maxlength="40" size="46" readonly value="<%=strCEPMig%>"></td>
</tr>

</tr>
<tr class=clsSilver>
<td>Erro</td>
<td><input type="text" class="text" name="txtErroMig" maxlength="40" size="46" readonly value="<%=strErroMig%>"></td>
</tr>
<!-- Permite a correção do Endereço -->
<tr class=clsSilver>
	<td>:: CEP</td>
	<td><input type="text" class="text" name="txtCepConsulta" value =  '<% = strCepConsulta %>'  maxlength="9" size="10" onKeyPress="OnlyNumbers();AdicionaBarraCep(this)" >(99999-999)
		<input type=button name=btnProcurarCepInstala value="Procurar CEP" class="button" onclick="<%if var_liberaredicao <> true then%>ProcurarCEP()<%else%>recuperaestadocep()<%end if%>" onmouseover="showtip(this,event,'Procurar por CEP exato ou pelos 5 primeiros dígitos (Alt+D)');" accesskey="D">
	</td>
</tr>

<tr class=clsSilver>
	<span id='spnCeps'> 
			<%
				if strCboRet <> "" then 
					Response.Write(strCboRet)
				end if 
			%>
	</span>
</tr>

</tr>
<%if var_liberaredicao = true then%>
<tr class=clsSilver>
<td>:: Tipo de Logradouro</td>
<td>
<%
'seleciona os tipos de logradouro
set objRS = db.execute("CLA_sp_sel_tplogradouro")
%>
	<select name="cboTipoLogr">
	<Option value=""></Option>
	<%do while not objRS.eof%>
		<option value="<%=trim(objRS("Tpl_Sigla"))%>" 
		<%If Trim(strLogr) = Trim(objRS("Tpl_Sigla")) then Response.write "selected" End if
		%>><%=trim(objRS("Tpl_Sigla"))%></option>
	<%
		objRS.movenext
	loop
	%>
	</select>
</td>
</tr>
<%end if%>
<tr class=clsSilver>
<td><%if var_liberaredicao = true then%>:: <%end if%>Logradouro</td>
<td><input type="text" class="text" name="txtRuaCompleta" maxlength="40" size="46" <%if var_liberaredicao <> true then%>Readonly<%end if%> value="<%=strRuaCompleta%>"></td>
</tr>

<%if var_liberaredicao <> true then%>
<tr class=clsSilver>
<td>:: Número</td>
<td><input type="text" class="text" name="txtNumero" value="<%=strNumero%>" maxlength="10" size="12"></td>
</tr>

<tr class=clsSilver>
<td>:: Complemento</td>
<td><input type="text" class="text" name="txtComplemento" value="<%=strComplemento%>" maxlength="30" size="35"></td>
</tr>
<%end if%>
<tr class=clsSilver>
<td>:: Bairro</td>
<td><input type="text" class="text" name="txtBairro" value="<%=strBairro%>" maxlength="30" size="35" ></td>
</tr>

<tr class=clsSilver>
<td>Cidade</td>
<td><input type="text" class="text" name="txtCidade" maxlength="40" size="46" Readonly value="<%=strCidade%>"></td>
</tr>

<tr class=clsSilver>
<td><%if var_liberaredicao = true then%>:: <%end if%>CNL</td>
<td><input type="text" class="text" name="txtCNL" maxlength="4" onblur="ResgatarCidadeLocal()" size="5" <%if var_liberaredicao <> true then%>Readonly<%end if%> value="<%=strCNL%>">
&nbsp;<input type="text" class="text" name="txtLocalidade"  maxlength="40" size="46" Readonly value="<%=strLocalidade%>">&nbsp;
</td>
</tr>
<tr class=clsSilver>
<td><%if var_liberaredicao = true then%>:: <%end if%>UF</td>
<td><input type="text" class="text" name="txtUF" maxlength="2" size="3" <%if var_liberaredicao <> true then%>Readonly<%end if%> value="<%=strUF%>"></td>
</tr>

<tr class=clsSilver>
<td>CEP</td>
<td><input type="text" class="text" name="txtCEP" maxlength="9" size="10" readonly value="<%if strCepConsulta <> "" then%><% = strCepConsulta %><%else%><%=strCEP%><%end if%>"></td>
</tr>

</table>

<table width="760" border=0>
<tr >
	<td colspan=2 align="center"><br>
	    <%if var_liberaredicao = true then%>
		<input type="button" class="button" name="btnIncluir" value="Incluir Endereço" accesskey="C" onmouseover="showtip(this,event,'Incluir (Alt+C)');">&nbsp;
		<%end if%>
		<input type="button" class="button" name="btnGravar" value="Gravar" onclick="GravarInfo()" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');" <%if var_liberaredicao = true then%>disabled<%end if%>>&nbsp;
		<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('ListarEndMig.asp?uf=<%=strUFMig%>&hdnDM=<%=DM%>&hdnFiltroDM=<%=FiltroDM%>')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
	</td>
</tr>
</table>

<table width="760" border=0>
<tr>
<td>
<font class="clsObrig">:: </font> Campos editáveis.
</td>
</tr>
</table>

</td>
</tr>
</table>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnCidSel>

<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "100%" 
	    height      = "100%"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
<SCRIPT LANGUAGE=javascript>
<!--
function ResgatarCidadeLocal()
{
    //alert ("OK")
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarCidadeLocal"
		hdnUFAtual.value = txtUFMig.value
		//alert (hdnUFAtual.value)
		hdnCidSel.value = txtCNL.value
		//alert (hdnCidSel.value)
		target = "IFrmProcesso"
		action = "ProcessoCla.asp"
		submit()
	}
}
//-->
</Script>

</form>
<%
if Request.Form("EndGravadoMig") = "OK" then
  %>
  <script>
    ProcurarCEP()
  </script>
  <%
end if
%>

</body>
</html>
