<!--#include file="../inc/data.asp"-->

<%
Dim nup
Dim strNroAcesso
Dim idFisico
Dim idSolicitacao
Dim dblSolId
Dim dblPedId
Dim dblAcfId
Dim strAcaoPedido
Dim strDM
Dim gravarDireto

	If len(request("ID")) > 3 then
		nup = request("Id")
		if Trim(nup) <> "" then
			Set rs = db.execute("CLA_sp_sel_ProvidenciaSnoa '"& mid(nup, 1, 2) &"',"& mid(nup, 4, 5) &","& mid(nup, 10, 4) & ",null,null,null,null")
			if Not rs.Eof and Not rs.bof then
				dblSolId = rs("Sol_Id")
				dblPedId = rs("Ped_Id")
				dblAcfId      = rs("Acf_id")
				gravarDireto  = "1"
				strAcaoPedido = rs("Tprc_Des")
				strDM = ucase(rs("Ped_Prefixo")) & "-" & right("00000" & rs("Ped_Numero"),5) & "/" & rs("Ped_Ano")
			Else
				Response.Write "<script language=javascript>parent.resposta(730,'')</script>"			
				Response.End 
			End if	
			
		End if
	
	Else
	 	
	 	If request("numero") <> "" then
			strNroAcesso = request("numero")
			if strNroAcesso = "" then 
				strNroAcesso = "null"
			Else
				strNroAcesso = "'" & strNroAcesso & "'"
			End if	

			Set rs = db.execute("CLA_sp_sel_ProvidenciaSnoa null,null,null," & strNroAcesso &",null,null,null")
			if Not rs.Eof and Not rs.bof then
				dblSolId = rs("Sol_Id")
				strAcaoPedido = rs("Tprc_Des")
				strDM = ucase(rs("Ped_Prefixo")) & "-" & right("00000" & rs("Ped_Numero"),5) & "/" & rs("Ped_Ano")
				dblAcfId = rs("Acf_id")
				gravarDireto  = "1"
			Else
				Response.Write "<script language=javascript>parent.resposta(730,'')</script>"			
				Response.End 
			End if	
		Else
 		
 			If request("idFisico") <> "" then
 				idFisico = request("idFisico")
				if idFisico = "" then 
					idFisico = "null"
				Else
					idFisico = "'" & idFisico & "'"
				End if	

				Set rs = db.execute("CLA_sp_sel_ProvidenciaSnoa null,null,null,null," & idFisico &",null,null")
				if Not rs.Eof and Not rs.bof then
					dblSolId = rs("Sol_Id")
					strAcaoPedido = rs("Tprc_Des")
					strDM = ucase(rs("Ped_Prefixo")) & "-" & right("00000" & rs("Ped_Numero"),5) & "/" & rs("Ped_Ano")
					dblAcfId = rs("Acf_id")
					gravarDireto  = "1"
				Else
					Response.Write "<script language=javascript>parent.resposta(730,'')</script>"			
					Response.End 
				End if	
			
			Else
 			
 				If request("idSolicitacao") <> "" then
					idSolicitacao = request("idSolicitacao")
					if idSolicitacao = "" then 
						idSolicitacao = "null"
					Else
						idSolicitacao = "'" & idSolicitacao & "'"
					End if	

					Set rs = db.execute("CLA_sp_sel_ProvidenciaSnoa null,null,null,null,null," & idSolicitacao & ",null")
					if Not rs.Eof and Not rs.bof then
						dblSolId = rs("Sol_Id")
						strAcaoPedido = rs("Tprc_Des")
						dblAcfId = rs("Acf_id")
						gravarDireto  = "0"
						strDM = "------"
					Else
						Response.Write "<script language=javascript>parent.resposta(730,'')</script>"			
						Response.End 
					End if	
				
				Else

	 				If request("numeroSNOA") <> "" then
						numeroSNOA = request("numeroSNOA")
						if numeroSNOA = "" then 
							numeroSNOA = "null"
						Else
							numeroSNOA = "'" & numeroSNOA & "'"
						End if	

						Set rs = db.execute("CLA_sp_sel_ProvidenciaSnoa null,null,null,null,null,null, " & numeroSNOA )
						if Not rs.Eof and Not rs.bof then
							dblSolId = rs("Sol_Id")
							strAcaoPedido = rs("Tprc_Des")
							dblAcfId = rs("Acf_id")
							gravarDireto  = "0"
							strDM = "------"
						Else
							Response.Write "<script language=javascript>parent.resposta(730,'')</script>"			
							Response.End 
						End if	
					End if
				End if
			End if
		End if
	End if
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel=stylesheet type="text/css" href="../css/cla.css">
</HEAD>

<BODY topmargin=0 leftmargin=0>

<SCRIPT LANGUAGE="JavaScript">

var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")

function DetalharItem()
{
	with (document.forms[0])
	{
		
		var strNome = "Facilidade" + hdnSolId.value + hdnPedId.value
		var objJanela = window.open()
		objJanela.name = strNome
		target = strNome
		//target = window.top.name
		action = "facilidade_new_cns.asp"
		submit()
	}
}
</script>



<%
if dblSolId <> "" then
%>

<Form name="Form1" method="Post">

<input type=hidden name=hdnSolId value="<%=dblSolId %>">
<input type=hidden name=hdnPedId value="<%=dblPedId %>">

<table border=0 cellspacing="1" cellpadding="0" width="760"> 
	<th>&nbsp;•&nbsp;Dados da Solicitação</th>
	<!--
	<th align=rigth><a href="javascript:DetalharItem()"><font color=white>Mais...</font></a></th>
	-->
	<th align=rigth>
		<a href="javascript:DetalharItem()"> 
		<font color=white>Mais...</font></a>
	</th>

</table>

<table border=0 cellspacing="1" cellpadding="0" width="760"> 
	<tr class=clsSilver>
		<td nowrap width=150px height=20px>Solicitação de Acesso</td>
		<td nowrap ><%=dblSolId%></td>

	</tr>
	<tr class=clsSilver>
		<td nowrap width=150px height=20px>Pedido de Acesso</td>
		<td nowrap ><%=strDM%></td>
	</tr>
	<tr class=clsSilver>
		<td nowrap width=150px height=20px>Ação do Pedido</td>
		<td nowrap ><%=strAcaoPedido%></td>
	</tr>
</table>

<!--
<table border=0 cellspacing="1" cellpadding="0" width="760">  
	<tr>
		<td>
			<iframe	id			= "IFrmMotivoPend"
				    name        = "IFrmMotivoPend" 
				    width       = "100%" 
				    height      = "150px"
				    src			= "../inc/MotivoPendencia.asp?dblSolId=<%=dblSolId%>&dblPedId=<%=dblPedId%>&dblAcfId=<%=dblAcfId%>&gravarDireto=<%=gravarDireto%>"
				    frameborder = "0"
				    scrolling   = "auto" 
				    align       = "left">
			</iFrame>
		</td>
	</tr>
</table>
-->
<table border=0 cellspacing="1" cellpadding="0" width="760">
	<tr>
		<td>
			<iframe	id			= "IFrmListaStatusSNOA"
				    name        = "IFrmListaStatusSNOA"
				    width       = "100%"
				    height      = "300px"
				    src			= "../inc/ListaStatusSNOA.asp?dblSolId=<%=dblSolId%>&dblPedId=<%=dblPedId%>&telaaceitar=3"
				    frameborder = "0"
				    scrolling   = "auto"
				    align       = "left">
			</iFrame>
		</td>
	</tr>
</table>




<%End if%>

</form>
</BODY>
</HTML>
