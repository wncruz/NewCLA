<%
'•EMBRATEL
'	- Sistema			: CLA
'	- Arquivo			: Monitor_jobs.asp
'	- Responsável		: PRSSILV
'	- Descrição			: Monitora os Jobs Ativos.
'	- Criação Autonoma	: 13/09/2008
%>
<!--#include file="../inc/data.asp"-->
<%
strLoginRede = ucase(mid(Request.ServerVariables("Logon_User"), Instr(Request.ServerVariables("Logon_User"),"\")+1))

if strLoginRede = "MSCAPRI" or strLoginRede  = "EDAR" then

db.execute("CLA_sp_relaservico3")
%>
<html>
<head>
<meta http-equiv="Content-Language" content="pt-br">
<title>Monitoração de JOBS - <%=date()%> - <%=Time()%></title>

<link rel="stylesheet" type="text/css" href="../css/stylesheets.css">
</head>

<br>
<center><h3>Monitoração de JOBS - <%=date()%> - <%=Time()%></h3></center>

<table align="center" class=tabela_registros cellspacing="2" cellpadding="2" border="0" width="642" bordercolor="#CCE7FF">
<tr class=titulos_registros>
	<td align="center"><font color="#FFFFFF"><b>Sistema</b></font></td>
	<td align="center"><font color="#FFFFFF"><b>Period</b></font></td>
	<td align="center"><font color="#FFFFFF"><b>JOB</b></font></td>
	<td align="center"><font color="#FFFFFF"><b>Hora Exec.</b></font></td>
	<td align="center"><font color="#FFFFFF"><b>Data Últ. Exec.</b></font></td>
	<td align="center"><font color="#FFFFFF"><b>Status</b></font></td>
</tr>
<%
Set objRS = db.execute("select Job_ID,Job_Sistema,Job_desc,job_hora_exec,job_dt_ult_exec,Job_Status,Job_Period from cla_relaservico3 order by job_sistema,job_desc")

Do Until objRS.EOF = True 
%>
<tr onMouseOver="this.style.backgroundColor='#CCE7FF';" onMouseOut="this.style.backgroundColor='';">
	<td align="center" nowrap><%=objRS("job_sistema")%></td>
	<td align="center" nowrap><%=objRS("Job_Period")%></td>
	<td align="left" nowrap><a href="Relaservico3_det.asp?<%=objRS("job_id")%>" target="_blank" title="Clique para detalhar..."><%=objRS("job_desc")%></a></td>
	<td align="center" nowrap><%if not isnull(objRS("job_hora_exec")) then response.write timevalue(objRS("job_hora_exec")) else response.write "-" end if %></td>
	<td align="center" nowrap><%=objRS("job_dt_ult_exec")%></td>
	<td align="center" nowrap><b>
	<%select case objRS("job_status")
	    case 1
	      response.write "<font color ='#009900'>OK"
	    case 2
	      response.write "<font color ='#C6A800'>Em Execução"
	    case 3
	      response.write "<font color ='#CA0000'>Erro"
		case 4
	      response.write "<font color ='#606060'>Não executada"
	  end select
	%></font></b></td>
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
<br>
<%else%>
  <center>Página em construção.</center>
<%end if%>
</html>