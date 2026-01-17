<%
'•EXPERT INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: interface_main.asp
'	- Responsável		: PRSS
'	- Descrição			: Lista/Remove interfaces no sistema utilizadas pela Solicitacao.asp
%>
<!--#include file="../inc/data.asp"-->
<%
Dim dblCepID
Dim strUf			
Dim strLocalidade	
Dim strLogr			
Dim strNomeLogr		
Dim strCep			
Dim strBairro		
Dim strEmail				
Dim strSel
Dim strCboRet
Dim intCount
Dim strCepConsulta

dblCepID = request("ID")
if Trim(dblCepID) = "" then
	dblCepID = Request.Form("hdnId") 
End if

strCboRet = ""

if Trim(Request.Form("hdnAcao"))="Gravar_mig" then 
    strAcao = "INS"
	
	Vetor_Campos(1)="adInteger,2,adParamInput,"
	Vetor_Campos(2)="adWChar,2,adParamInput,"	&	request.Form("txtufmig")
	Vetor_Campos(3)="adWChar,4,adParamInput,"	&	request.Form("txtCnl")
	Vetor_Campos(4)="adWChar,15,adParamInput,"	& 	request.Form("cboTipoLogr")
	Vetor_Campos(5)="adWChar,60,adParamInput,"	& 	request.Form("txtTitulo")
	Vetor_Campos(6)="adWChar,3,adParamInput,"	& 	request.Form("txtPreposicao")
	Vetor_Campos(7)="adWChar,60,adParamInput,"	& 	request.Form("txtRuaCompleta")
	Vetor_Campos(8)="adWChar,9,adParamInput,"	&	request.Form("txtCep")
	Vetor_Campos(9)="adWChar,60,adParamInput,"	& 	request.Form("txtBairro")
	Vetor_Campos(10)="adWChar,3,adParamInput,"	& 	strAcao
	Vetor_Campos(11)="adInteger,3,adParamOutput,0"
	
	var_habilita_response = FALSE
	if var_habilita_response = true then
	  Response.Write "<b>CLA_sp_basecorreio</b><br><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Cep_ID</b>=<font color='red'>'" & "" & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Est_Sigla </b>=<font color='red'>'" & request.Form("txtufmig") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Cid_Sigla </b>=<font color='red'>'" & request.Form("txtCnl") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Tpl_Sigla</b>=<font color='red'>'" & request.Form("cboTipoLogr") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Titulo</b>=<font color='red'>'" & request.Form("txtTitulo") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Preposicao </b>=<font color='red'>'" & request.Form("txtPreposicao") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@NomeLogr</b>=<font color='red'>'" & request.Form("txtRuaCompleta") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Cep</b>=<font color='red'>'" & request.Form("txtCep") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Bairro</b>=<font color='red'>'" & request.Form("txtBairro") & "'</font><br>"
	  Response.Write "<font color='blue'>SET </font><b>@Acao</b>=<font color='red'>'" & strAcao & "'</font><br>"
	  'Response.Write "<br><a href='menumig.asp'>Voltar</a>"
	  Response.end
	end if  
	
	Call APENDA_PARAM("CLA_sp_basecorreio ",11,Vetor_Campos)
	
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
End if


REM Obtem os argumentos enviados pela página ListarEndMig.asp -------------
intIdEndMig  = request("EndId")
if trim(intIdEndMig) = "" then intIdEndMig = request.form("hdnIdEndMig")

intIdCompMig = request("AecId")
if trim(intIdCompMig) = "" then intIdCompMig = request.form("hdnIdCompMig")
REM -----------------------------------------------------------------------

action = "CorrigirEndMig.asp?EndId=" &intIdEndMig& "&AecId=" &intIdCompMig &""

%>
<Form name="Form_1" method="post" action="<%=action%>">
  <input type="hidden" name="txtufmig" value="<%=request.Form("txtufmig")%>">
  <input type="hidden" name="txtCnl" value="<%=request.Form("txtCnl")%>">
  <input type="hidden" name="cboTipoLogr" value="<%=request.Form("cboTipoLogr")%>">
  <input type="hidden" name="txtTitulo" value="<%=request.Form("txtTitulo")%>">
  <input type="hidden" name="txtPreposicao" value="<%=request.Form("txtPreposicao")%>">
  <input type="hidden" name="txtRuaCompleta" value="<%=request.Form("txtRuaCompleta")%>">
  <input type="hidden" name="txtCep" value="<%=request.Form("txtCep")%>">
  <input type="hidden" name="txtBairro" value="<%=request.Form("txtBairro")%>">  
  <input type="hidden" name="EndGravadoMig" value="OK">
</form>

<script>
  Form_1.submit()
</script>


