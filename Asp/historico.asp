<!--#include file="../inc/data.asp"-->

<!--#include file="../inc/header.asp"-->
<%
if request("action") = "Gravar" then
	Vetor_Campos(1)="adInteger,50,adParamInput,"& ucase(request("Ped_ID"))
	Vetor_Campos(2)="adInteger,50,adParamInput,"& ucase(dblUsuId)
	Vetor_Campos(3)="adInteger,50,adParamInput,"& ucase(request("status"))
	Vetor_Campos(4)="adLongvarchar,800,adParamInput,"& ucase(request("historico"))
	Vetor_Campos(5)="adInteger,2,adParamOutput,0" 
	Call APENDA_PARAM("CLA_sp_ins_historico",5,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
	If DBAction = 1 then
		Set sts = db.execute("CLA_CLA_sp_sel_Statuspedido " & request("status"))

		If sts("Stp_Notifica") = true then

			Set ped = db.execute("CLA_sp_sel_pedido " & request("ped_id"))
			Set cli = db.execute("CLA_sp_sel_cliente " & ped("cli_id"))

			Stp_CS = sts("Stp_CS")
			Stp_GIC = sts("Stp_GIC")
			Stp_GLA = sts("Stp_GLA")

			'ENVIO DE EMAIL PARA CS, GIC, GLA
			Set pedido = db.execute("CLA_sp_sel_pedido " & ped_id)
			Set numero = db.execute("CLA_sp_view_pedidonumped " & ped_id)
			numero_pedido = ucase(numero("Acp_Prefixo") & "-" & right("00000" & numero("Acp_Numero"),5) & "/" & numero("Acp_Ano"))

			Set cliente = db.execute("CLA_sp_sel_cliente " & pedido("cli_id"))
			from = "acessosp@embratel.com.br"

			'Resgata informações do pedido para o subject
			Set objRSPed = db.execute("CLA_sp_sel_emailprovedor " & ped_id)
			if not objRSPed.Eof and Not objRSPed.Bof then
				subject = trim(ucase(objRSPed("aca_desc_a"))) & "  -  " & trim(objRSPed("cli_nome")) & "  -  " & ucase(objRSPed("prefixo")) & "-" & right("00000" & objRSPed("numero"),5) & "/" & objRSPed("ano")
			Else
				subject	= numero_pedido				
			End if	

			message = "<table rules=groups bgcolor=#eeeeee cellspacing=0 cellpadding=5 bordercolorlight=#003388 bordercolordark=#ffffff width=680>"
			message = message & "<tr><td><font face='verdana' color='#003388'>"
			message = message & "O Status do pedido <b>" & numero_pedido & "</b>, cliente " & cliente("cli_nome") & ",<br>"
			message = message & "contrato " & pedido("ped_ncontratoebt") & ", foi alterado em "
			message = message & right("00" & day(date),2) & "/" & right("00" & month(date),2) & "/" & year(date) & " para:"
			message = message & "</font></td></tr>"
			message = message & "<tr><td><font face='verdana' color='#003388'>"
			message = message & "Status: <b>" & sts("Stp_Desc") & "</b>"
			message = message & "</font></td></tr>"
			message = message & "<tr><td><font face='verdana' color='#003388'>"
			message = message & "Histórico:"
			message = message & "</font></td></tr>"
			message = message & "<tr><td><font face='verdana' color='#003388'>"
			message = message & "<b>" & request("historico") & "</b>"
			message = message & "</font></td></tr></table>"
							
			'Envio de e-mails para Agentes do pedido CS,GIC e GLA
			Set objRSAgente = db.execute("CLA_sp_view_agentepedido " & ped_id)
					
			While not objRSAgente.Eof
				blnEnviar = false
				Select Case Ucase(Trim(objRSAgente("AGE_DESC")))
					Case "GLA"
						if Stp_GLA then	blnEnviar = true End if
					Case "CS"
						if Stp_CS then blnEnviar = true	End if
					Case "GIC"
						if Stp_GIC then	blnEnviar = true End if
					Case Else
						blnEnviar = true
				End Select
				'Dispara e-mail
				if blnEnviar and Trim(objRSAgente("Usu_Email")) <> "" and not isnull(objRSAgente("Usu_Email")) then
					toEmail = Trim(objRSAgente("Usu_Email"))
					email from, toEmail, subject, message
				End if	
				objRSAgente.MoveNext
			Wend

		end if
	end if
end if

'set rb = db.execute("CLA_sp_sel_numeropedido "& request("Nup_ID")) 
set rb = db.execute("CLA_sp_view_pedidoacaopedido "& request("Ped_ID")) 
%>

<table width="760"><tr><td align="center">
<tr>
<td align="center" class="titulo">Histórico<br><br></td>
</tr>

<tr><td>
<table rules="groups" bgcolor="#eeeeee" cellspacing="0" cellpadding="5" bordercolorlight="#003388" bordercolordark="#ffffff" width="760"> 
 <tr>
 <td>
 <center>
 <form method=post name="f" action="historico.asp?Ped_ID=<%=request("Ped_ID")%>&Nup_ID=<%=request("Nup_ID") %>" onSubmit="return checa(this)">
 <table border=0>
 <tr>
 <td>Pedido de acesso:</td>
 <td><%=ucase(rb("Acp_prefixo") & "-" & right("00000"& rb("Acp_numero"),5) & "/" & rb("Acp_ano"))%></td>
 <input type="hidden" name="nropedido" value="<%=ucase(rb("Acp_prefixo") & "-" & right( "00000"& rb("Acp_numero"),5) & "/" & rb("Acp_ano"))%>">
 </tr>
 <tr>
 <td><font class="clsObrig">:: </font>Status do Pedido</td>
 <td>
 <select name="status">
 	<option value=""></option>
 <%set st = db.execute("CLA_sp_sel_Status null,1")
   do	 while not st.eof
 %>
 <option value="<%=st("sts_id")%>"><%=ucase(st("Sts_Desc"))%>
<%
st.movenext
loop
%>
 </select>
</td>
</tr>

<tr>
<td>Histórico</td>
</tr>

<tr>
<td colspan="2"><textarea name="historico" cols="70" rows="4"></textarea></td>
</tr>

<tr>
<td colspan=2 align="center">
<SCRIPT LANGUAGE="JavaScript">
function checa(f) {
	if (f.status.value == "") {
		alert("O status é um campo obrigatório !");
	    f.status.focus();
    	return false;
    }	
	if (f.data.value == "") {
		alert("A data é um campo obrigatório !");
	    f.data.focus();
    	return false;
    }
	if (f.data.value.substr(2,1) != "/" || f.data.value.substr(5,1) != "/") {
			alert("Data em Formato inválido!");
			f.data.focus();
	    	return false;
	}
	if (f.historico.value == "") {
		alert("O historico é um campo obrigatório !");
	    f.historico.focus();
    	return false;
    }
	return true;
}
</script>
<input type="submit" name="action" value="Gravar" class="button">
</td>
</tr>
</table>
</form>
<%
set sp = db.execute("CLA_sp_sel_historico "& request("Ped_ID"))
if not sp.eof then
 %> 
 <table cellpadding="1" cellspacing="0" width="760">
 <tr>
 <th width="120">&nbsp;Data</th>
 <th>&nbsp;Usuario</th>
 <th>&nbsp;Status</th>
 <th>&nbsp;Historico</th>
 </tr>
	<%
	do while not sp.eof
	set st = db.execute("CLA_CLA_sp_sel_Statuspedido "&sp("Stp_ID")) 	
	set us = db.execute("CLA_sp_sel_usuario "&sp("Usu_ID"))
	%>
	<tr>
	<% Hist = trim(sp("Hip_Historico"))%>
	<td>&nbsp;<%=right("00" & day(sp("Hip_Data")),2) & "/" & right("00" & month(sp("Hip_Data")),2) & "/" & year(sp("Hip_Data")) & " - " & right("00" & hour(sp("Hip_Data")),2) & ":" & right("00" & minute(sp("Hip_Data")),2) & ":" & right("00" & second(sp("Hip_Data")),2)%></td>
	<td>&nbsp;<%=us("Usu_nome")%></td>
	<td>&nbsp;<%=ucase(st("Stp_Desc"))%></td>
	<td>&nbsp;<%=Hist%></td>
	</tr>
	<%
	st.close
	us.close
	sp.movenext
	loop
	%>
 </table>
	<%end if %>
</td>
</tr>
</table>
<table width="760">
<tr>
<td>
<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
</td>
</tr>
</table>
</td>
</tr>
</table>
</body>
</html>
<%
sp.close
rb.close
%>
