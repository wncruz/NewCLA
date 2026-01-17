<%
'	- Sistema			: CLA
'	- Arquivo			: headerMig.ASP
'	- Descrição			: Arquivo base para aplicativo de correção do cla
%>
<Html>
<Head>
<Title>CLA - Controle Local de Acesso</Title>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='JavaScript' src='../javascript/formatamenu.js'></script>
<script language='JavaScript' src='../javascript/montamenu.js'></script>
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>
<Script language='javascript'>
if (top.location != self.location)
{
	top.location = self.location;
}
try{
<%=strGeral%>
}
catch(e){
	alert(e.description)
}
</Script>
<SCRIPT language='javascript'>
  javascript:window.history.forward(1);
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" onLoad="resposta(<%if DBAction <> "" then response.write DBAction else response.write "0" end if%>,'main.asp');">
<table width="760" border="0" cellspacing="0" cellpadding="0">
<tr bgcolor="#f1f1f1"> 
	<td valign=top>	<!--a href="main.asp"-->
		<img name="embratel" src="../imagens/topo_embratel.jpg" width=760px height=80px border="0"><!--/a-->
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
<!--tr>
	<td height=12px bgcolor="#f1f1f1" valign=top>
		<div id=divMenu style="position:absolute;top:81;left:5">
			<table width="<%=190*(intTotalCelula)%>" border="0" cellspacing="0" cellpadding="0" >
				<tr>
					<%if blnAcessoLog then%>
					<td nowrap width=175px >
						<span id=spnAcessoLog onMouseOut="popDown('eMenu1')" onClick="showInput(false);popUp('eMenu1',event)" style='cursor:hand' ><font class=clsMenu >Acesso Lógico</font></span>
					</td>
					<%End if%>
					<%if blnAcessoFis and blnAcessoLog then%>
					<td nowrap width=175px >
						<span id=spnAcessoFis  onMouseOut="popDown('eMenu2')" onClick="showInput(false);popUp('eMenu2',event)" style='cursor:hand' ><font class=clsMenu >|&nbsp;Acesso Físico</font></span>
					</td>
					<%Else%>
						<%if blnAcessoFis and not blnAcessoLog then%>
						<td nowrap width=175px >
							<span id=spnAcessoFis onMouseOut="popDown('eMenu1')" onClick="showInput(false);popUp('eMenu1',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Acesso Físico</font></span>
						</td>
						<%End if%>
					<%End if%>
					<td nowrap width=175px >
						<span id=spnConsultas onMouseOut="popDown('eMenu<%=intMenuOrd%>')" onClick="showInput(false);popUp('eMenu<%=intMenuOrd%>',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Consultas</font></span>
					</td>
					<%if blnTabelas then %>
					<td nowrap width=175px >
						<span id=spnTabelas onMouseOut="popDown('eMenu<%=intMenuOrd+1%>')" onClick="showInput(false);popUp('eMenu<%=intMenuOrd+1%>',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Tabelas</font></span>
					</td>
					<%End if%>
					<td nowrap >
						<span id=spnCrms style='cursor:hand' onClick="window.open('http://<%=Request.ServerVariables("SERVER_NAME")%>/crmsf/inicio.htm','CRMSF','')" onmouseover="showtip(this,event,'Controle de Radio,Mux,Satélite e Fibra');" onmouseout="hidetip();"><font class=clsMenu >|&nbsp;CRMSF</font></span>
					</td>
				</tr>
			</table>
		</div>
	</td>
</tr-->
</table>
<table valign="top" width="780" cellspacing="2" cellpadding="0">
<input type=hidden name=hdnUserHerder value="<%=strUserName%>">
</table>