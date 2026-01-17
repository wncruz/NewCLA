<!--#include file="../inc/data.asp"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
</HEAD>
<BODY topmargin=0 leftmargin=0 class=TA>
<form name=Form1 method=Post>
<table width=100% cellspacing="1" cellpadding="0">
	<tr >
		<th width=3%>&nbsp;Selecione&nbsp;</th>
		<th width=15%>&nbsp;Designação</th>
	</tr>
	
	<%
	'dim idAcessoFisico
	set idAcessoFisico =  Request.QueryString("hdnIdAcessoFisico") 'Request.Form("hdnIdAcessoFisico")
	
	strSQL = "select desig_id , desig_designacao  from cla_designacao inner join cla_acessofisico on cla_designacao.acf_id = cla_acessofisico.acf_id where acf_idAcessofisico =  '" & idAcessoFisico & "'"
	
	set objRSdesig = server.createobject("ADODB.RECORDSET")
	
	set objRSdesig = db.execute(strSQL)
	
	if Not objRSdesig.Eof and Not objRSdesig.Bof then
		while not objRSdesig.eof
		if strClass = "clsSilver" then strClass = "clsSilver2" else strClass = "clsSilver" End if
			%>
				<tr class="<%=strClass%>" width=100% >
					<td width=3%><input type=checkbox name="checkCompTronco2m" value="<%=objRSdesig("desig_id")%>"></td>
					<td width=15%><%=objRSdesig("desig_designacao")%></td>
				</tr>
				
			<% 
			objRSdesig.movenext
		wend
		Response.Write "<script language=javascript>parent.divTronco2M.style.display = '';</script>"
	Else
		Response.Write "<script language=javascript>alert('Não existe Tronco 2M para o ID Físico!')</script>"
		Response.Write "<script language=javascript>parent.divTronco2M.style.display = 'none'; </script>"
	End if
		
	%>
</table>	
</Form>
</BODY>
</HTML>
