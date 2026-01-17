<%@ CodePage=65001 %>
<%
	Response.ContentType = "text/html; charset=utf-8"
	Response.Charset = "UTF-8"
%>
<!--#include file="../inc/data.asp"-->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel=stylesheet type="text/css" href="../css/cla.css">
</HEAD>
<BODY topmargin=0 leftmargin=0>
<!--#include file="ConsultaProcesso.asp"-->
<SCRIPT LANGUAGE=javascript>
<!--
self.focus() 
//-->

			function DetalharItem()
			{
				with (document.forms[0])
				{
		
					//alert(hdnPedId2.value)
					hdnPedId.value = hdnPedId2.value
					var strNome = "Facilidade" + hdnSolId.value + hdnPedId2.value
					var objJanela = window.open()
					objJanela.name = strNome
					target = strNome
					//target = window.top.name
					action = "facilidade_new_cns.asp"
					submit()
				}
			}

</SCRIPT>

<Form name=Form1 method=Post action="facilidade.asp">
			
	<input type=hidden name=hdnSolId value="<%=dblSolId%>">
	<input type=hidden name=hdnPedId value="<%=dblPedId%>">
	<input type=hidden name=dblRecId value="<%=Request.Form("dblRecId")%>">
	<input type=hidden name=id value="<%=Request("id")%>">
	<input type=hidden name=gla value="<%=Request.Form("gla")%>">
	<input type=hidden name=provedor value="<%=Request.Form("provedor")%>">
	<input type=hidden name=status value="<%=Request.Form("status")%>">
	<input type=hidden name=cboProvedor value="<%=Request.Form("cboProvedor")%>">
	<input type=hidden name=cboEstacao value="<%=Request.Form("cboEstacao")%>">
	<input type=hidden name=cboDominioNO value="<%=Request.Form("cboDominioNO")%>">

			<%
				Set objRSPed =	db.execute("CLA_sp_view_pedido " & dblSolId & ",null,null,null,null,null,null,null,null,'T',null")

				if not objRSPed.Eof and not objRSPed.Bof then
					While not objRSPed.Eof
						dblPedId = objRSPed("Ped_ID")
						objRSPed.MoveNext
					Wend
					%>
					<input type=hidden name=hdnPedId2 value="<%=dblPedId%>">
					<%
				end if


				set ObjSnoa = db.execute("select top 1 pedido_compra_snoa  from cla_Assoclogicosnoa where sol_id =  " & dblSolId & " and pedido_compra_snoa is not null ")
				if not ObjSnoa.Eof and not ObjSnoa.Bof then
					strSnoa = UCASE(ObjSnoa("pedido_compra_snoa"))
				end if

			%>
	
	<table cellspacing=1 cellpadding=1  width=760 border=0>
		<tr>
			<td align=center>
					<% 	if trim(strSnoa) <> ""  then%> 
						<input type="button" class="button" name="btnFechar" value="Detalhes SNOA" onClick="javascript:DetalharItem()" >
					<%End if %>
			<input type="button" class="button" name="btnFechar" value="Fechar" onClick="javascript:window.close()" >
			</td>
		</tr>
	</table>		
</form>								 
<P>&nbsp;</P>
</BODY>
</HTML>
