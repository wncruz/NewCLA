<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: header.ASP
'	- Descrição			: Arquivo com o menu do CLA
%>
<Html>
<Head>
<Title>CLA - Controle Local de Acesso</Title>
<link rel=stylesheet type="text/css" href="../css/cla.css">
</head>
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>
<Script language='javascript'>
if (top.location != self.location)
{
	top.location = self.location;
}
</Script>

<body leftmargin="0" topmargin="0" onLoad="resposta(<%if DBAction <> "" then response.write DBAction else response.write "0" end if%>,'main.asp');">
<table width="760" border="0" cellspacing="0" cellpadding="0">
<tr bgcolor="#f1f1f1"> 
	<td valign=top>
		<a href="smbd_main.asp"><img name="embratel" src="../imagens/topo_embratel2.jpg" width=760px height=80px border="0"></a>
		<div style="position:absolute;left:550;top:0;">
			<table border=0 cellspacing=0 cellpadding="0">
				<tr>
					<td width=60><font color=white size=1>Servidor</font></td>
					<td><font color=white size=1><%=Request.ServerVariables("SERVER_NAME") %></font></td>
				</tr>	
				<tr>
					<td><font color=white size=1>Banco</font></td>
					<td><font color=white size=1><%=strBanco%></font></td>
				</tr>	
				<tr>
					<td><font color=white size=1>Usuário</font></td>
					<td><font color=white size=1><%=strUserName%></font></td>
				</tr>	
			</table>
		</div>
	</td>	
</tr>
</table>
<table valign="top" width="780" cellspacing="2" cellpadding="0">
<input type=hidden name=hdnUserHerder value="<%=strUserName%>">