<!--#include file="../inc/Data.asp"-->
<%
if strLoginRede = "PRSSILV" or strLoginRede = "EDAR" or strLoginRede = "MSCAPRI" then
	%>
	<html>
	<title>Checagem de Serviço 0800.</title>
	<center>
	<br>
	<br>
	<%
	OriSol_ID = Request("OriSol_ID")
	Strxml1 = Request("Strxml1")
	id_tarefa = Request("id_tarefa")
	
	if id_tarefa = "" then
		id_tarefa = "1"
	end if
	%>
	<form name="Form_1" method="post" action="RelaServico7.asp">
		<br>
		<input type="text" name="id_tarefa" value="<%=id_tarefa%>">
		<br>
	   <select name="orisol_id">
	   		<option value=""></option>
			<option value="7" <%if OriSol_ID = 7 then%>selected<%end if%>>SGAV</option>
			<option value="6" <%if OriSol_ID = 6 then%>selected<%end if%>>SGAP</option>
	   </select>
	   <br><br>
	   <textarea name="Strxml1" cols="100" rows="15"><%=Strxml1%></textarea>
	   <br><br>
	   <input type="submit" name="btnok" value="Retorno Entregar">	   
	    <input type="button" name="btnlimp" value="Limpar" onclick="Form_1.Strxml1.value=''">
	</form>
	<%
	if OriSol_ID <> "" AND Strxml1 <> ""  then

		set ConSGA = Server.CreateObject("ADODB.Command")
		
		If Ucase(Request.ServerVariables("SERVER_NAME")) = "NTSPO913" or  Ucase(Request.ServerVariables("SERVER_NAME")) = "10.100.1.21" or Ucase(Request.ServerVariables("SERVER_NAME")) = "NTSPO912" or  Ucase(Request.ServerVariables("SERVER_NAME")) = "10.100.1.17" then
					StrSQL = "select Conn_Desc from CLA_ConexaoInterf where Conn_Tipo = 'PRD' and OriSol_ID = " & OriSol_ID
		else
					StrSQL = "select Conn_Desc from CLA_ConexaoInterf where Conn_Tipo = 'DSV' and OriSol_ID = " & OriSol_ID
		end if
		
		Set objRS = db.Execute(StrSQL)
		If Not objRS.eof and  not objRS.Bof Then
			objConSGA = objRS("Conn_Desc")
		End if
		
		ConSGA.ActiveConnection = objConSGA
		
		if oriSol_id = 6 then
			ConSGA.CommandText = "sgaplus_adm.pck_sgap_interface_cla.pc_retorno_solicitacao_cla"
		end if
		if oriSol_id = 7 then
			ConSGA.CommandText = "sgav_vips.sp_sgav_interface_cla2"
			'ConSGA.CommandText = "sgav_atualiza.sp_sgav_interface_cla2"
		end if
		ConSGA.CommandType = adCmdStoredProc
		
		'*** Carregando parâmetros de entrada
		Set objParam = ConSGA.CreateParameter("p1", adNumeric, adParamInput, 10, id_tarefa)
		ConSGA.Parameters.Append objParam
		
		Set objParam = ConSGA.CreateParameter("p2", adLongVarWChar, adParamInput, 1073741823, Strxml1)
		ConSGA.Parameters.Append objParam
		
		'*** Configurando variável que receberá o retorno
		Set objParam = ConSGA.CreateParameter("Ret1", adNumeric, adParamOutput, 10)
		ConSGA.Parameters.Append objParam
		
		Set objParam = ConSGA.CreateParameter("Ret2", adVarChar, adParamOutput, 100 )
		ConSGA.Parameters.Append objParam
		
		'*** Executando a stored procedure
		ConSGA.Execute
		
		cod_retorno  = ConSGA.Parameters("Ret1").value
		desc_retorno = ConSGA.Parameters("Ret2").value
		
		strxmlResp = cod_retorno & " - " & desc_retorno
		
		Response.write strxmlResp

	End if
else
	response.write "0800"
end if
%>