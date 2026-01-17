<!--#include file="../inc/Data.asp"-->
<%
if strLoginRede = "EDAR" or strLoginRede = "SCESAR" or strLoginRede = "T3FRRP" then

	%>
	<html>
		<title>Gerar Nome do Acesso Fisico.</title>
		<center>
			<br>
			<br>
			<%
	
				strCid_Sigla 	= Request("txt_Cid_Sigla")
				strSol_ID 		= Request("txt_Sol_ID")
				strAcf_ID 		= Request("txt_Acf_ID")
	
			%>
	
			<form name="Form_1" method="post" action="RelaServico11.asp">

				<p> Gera o nome do acesso fisico quando não conseguiu gerar automaticamento através do sistema</p>
	   
	   			<input type="text" name="txt_Cid_Sigla" value="<%=strCid_Sigla%>"> -> Cid_Sigla
	   			<br><br>
	   
	   			<input type="text" name="txt_Sol_ID" value="<%=strSol_ID%>"> -> Sol_ID
	   			<br><br>

				<input type="text" name="txt_Acf_ID" value="<%=strAcf_ID%>"> -> Acf_ID
				<br><br>

				<input type="submit" name="btnok" value="Gerar Nome Fisico">	   
	
			</form>
	
	<%
	if strAcf_ID <> "" then

		if strCid_Sigla = "" then
			response.write "<font color=red>Favor informar o Cid_Sigla.</font>"
			response.end
		end if

		if strSol_ID = "" then
			response.write "<font color=red>Favor informar o Sol_ID.</font>"
			response.end
		end if

		if strAcf_ID = "" then
			response.write "<font color=red>Favor informar o ACF_ID.</font>"
			response.end
		end if

  		'response.write "<script>alert('"&strIDLogico&"')</script>"

		'@Cid_Sigla 		varchar(4)	,
		'@Sol_ID			int			,
		'@Acf_ID			int			,
		'@ret				varchar(15)	= null OUTPUT
		Vetor_Campos(1)="adWChar,4,adParamInput, " & strCid_Sigla 	'request("Cid_Sigla")
		Vetor_Campos(2)="adInteger,4,adParamInput," & strSol_ID 	'request("Sol_ID")
		Vetor_Campos(3)="adInteger,4,adParamInput," & strAcf_ID		'request("Acf_ID")
		Vetor_Campos(4)="adwchar,15,adParamOutput,"

		'on error resume next
		Call APENDA_PARAM("CLA_sp_ins_NumeroIDAcessoFisico_2",4,Vetor_Campos)
		ObjCmd.Execute'pega dbaction
		DBAction = ObjCmd.Parameters("RET1").value

		response.write "<font color=red>Nome do Acesso Fisico -> </font>" & DBAction

	else
		response.write "<font color=red>Favor informar os campos da tela.</font>"
	End if
	%>
	<br>
<%
else
	response.write "Pagina em Manutenção"
end if
%>