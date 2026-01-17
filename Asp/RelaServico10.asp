<!--#include file="../inc/Data.asp"-->
<!--#include file="../inc/EnviarEndereco1123_CSL.asp"-->

<%
if strLoginRede = "EDAR" then
'Option Explicit

Function GetRefreshDate()
	Dim d
	d = DateAdd("d", 0, now())
	GetRefreshDate = CDate(Year(d) & "/" & FixZero(Month(d)) & "/" & FixZero(Day(d)) & " 00:00:10")
	
	'GetRefreshDate = now() + "00:01:00"
End Function

Function FixZero(n)
	If n < 10 Then
		FixZero = "0" & n
	Else
		FixZero = "" & n
	End If
End Function

Dim intRefreshSeconds
intRefreshSeconds = DateDiff("s", Now(), GetRefreshDate())
%>
<html>
<head>
<title>Checagem de Serviço 1123. </title>
<meta http-equiv="refresh" content="5">
</head>
<body>
Last refresh: <%= Now() %>
<br>
Next refresh: <%= GetRefreshDate() %>
intRefreshSeconds: <%= intRefreshSeconds %>
</body>
</html>

	
	<%
	'strIDLogico = Request("txt_IDLogico")
	Aprovisi_ID = Request("txt_Aprovisi_ID")
	rdo_acao = Request("rdo_acao")
	%>
	<form name="Form_1" method="post" action="RelaServico8.asp">
	   <input type="text" name="txt_IDLogico" value="<%=strIDLogico%>">
	   <br><br>
	   <input type="submit" name="btnok" value="Enviar endereco por cep">	   
	    <input type="button" name="btnlimp" value="Limpar" onclick="Form_1.txt_IDLogico.value=''">
	</form>
	<%
	'if strIDLogico <> "" then
	'	EnviarEntregarAprovASMS(strIDLogico)
	'else
	'	response.write "<font color=red>Favor informar o ID lógico.</font>"
	'End if
	%>
	<br>
	<br>
	<form name="Form_2" method="post" action="RelaServico8.asp">
	   <input type="text" name="txt_Aprovisi_ID" value="<%=Aprovisi_ID%>">
	   <br>
	   <input type="radio" name="rdo_acao" value="End_COMPLETO" <%if rdo_acao = "End_COMPLETO" then%>checked<%end if%>> Validar Endereco completo
	  
	   
	   <br><br>
	   <input type="submit" name="btnok" value="Enviar">
	   <input type="button" name="btnlimp" value="Limpar" onclick="Form_2.txt_Aprovisi_ID.value=''">
	</form>	
	<%
	'if Aprovisi_ID <> "" then
		'if rdo_acao = "" then
		'	response.write "<font color=red>Favor informar a ação.</font>"
		'	response.end
		'end if
				
		
		
		'if rdo_acao = "End_COMPLETO" then
			EnviarEndereco1123_CSL
		'elseif rdo_acao = "RetSolicitar" then
			'Interface_Solicitar_Return ID_Tarefa,OrigemSol_ID,estacao,propAcesso,dblIdLogico,dblsol_id,Aprovisi_ID
		'else
			'EnviarAprovASMS(Aprovisi_ID) 'EnviarEntregarAprovASMS(strIDLogico)
			'Interface_Status_Return ID_Tarefa,OrigemSol_ID,"Solicitação iniciada no CLA.",dblIdLogico,dblsol_id,Aprovisi_ID
		'end if
	'else
	'	response.write "<font color=red>Favor informar o Aprovisi_ID.</font>"
	'End if
	
else
	response.write "1123"
end if
%>