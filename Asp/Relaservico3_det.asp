<%
'•EMBRATEL
'	- Sistema			: CLA
'	- Arquivo			: Detalhar_job.asp
'	- Responsável		: PRSSILV
'	- Descrição			: Detalha os eventos de cada job.
'	- Criação			: 04/09/2008
%>
<!--#include file="../inc/data.asp"-->
<%
strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

if strLoginRede = "MSCAPRI" or strLoginRede  = "EDAR" then
%>

<html>
<head>
<meta http-equiv="Content-Language" content="pt-br">
<title>Monitor de jobs - Detalhamento da execução do JOB <%=date()%> - <%=Time()%>.</title>

<link rel="stylesheet" type="text/css" href="../css/stylesheets.css">
</head>
<table style="width: 100%;" class=tabela_registros>
    <tr class=titulos_registros>
        <td align="center">&nbsp;<b>Nº</b></td>
        <td align="center">&nbsp;<b>Sistema</b></td>
		<td align="center">&nbsp;<b>Job</b></td>
		<td align="center">&nbsp;<b>Tarefa</b></td>
		<td align="center">&nbsp;<b>Data Execução</b></td>
        <td align="center">&nbsp;<b>Status</b></td>
    </tr>
<%
job_id = Request.ServerVariables("QUERY_STRING")

Set objRS = db.execute("select Tarefa_ID,Tarefa,Data,Status,Job_Sistema,Job_Desc from CLA_ValidarTransacao inner join cla_relaservico3 on CLA_ValidarTransacao.job_id = cla_relaservico3.job_id where CLA_ValidarTransacao.Job_id = " & job_id & " and CLA_ValidarTransacao.data >= Job_Dt_Ini_Ult_Exec")

Do Until objRS.EOF = True 
%>
    <tr>
        <td align="center" nowrap>&nbsp;<%=objRS("Tarefa_ID")%></td>
        <td align="center" nowrap>&nbsp;<%=objRS("Job_Sistema")%></td>
		<td align="left" nowrap>&nbsp;<%=objRS("Job_Desc")%></td>
		<td align="left" nowrap>&nbsp;<%=objRS("Tarefa")%></td>
		<td align="center" nowrap>&nbsp;<%=objRS("Data")%></td>
        <td align="center" nowrap>&nbsp;
		    <%if objRS("Status") = 0 then%>
              <font color="#009900"><b>EXECUTADA</b></font>
			<%else%>
              <font color="#CA0000"><b>FALHOU</b></font>
			<%end if%>
        </td>
    </tr>
<%
  objRS.MoveNext 
Loop  
%>
    <tr class=titulos_registros>
        <td align="center">&nbsp;</td>
        <td align="center">&nbsp;</td>
        <td align="center">&nbsp;</td>
		<td align="center">&nbsp;</td>
        <td align="center">&nbsp;</td>
        <td align="center">&nbsp;</td>
    </tr>
</table>

</html>

<%else%>
  <center>Página em construção.</center>
<%end if%>
</html>