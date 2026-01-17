<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<head>
    
</head>

<SCRIPT LANGUAGE="JavaScript">

function DetalharItem(dblSolId)
{
	with (document.forms[0])
	{
		hdnSolId.value = dblSolId
		hdnAcao.value = "DetalheSolicitacao"
		target = "DetalheSolic"
		action = "ConsultaGeralDet.asp"
		submit()
	}	
}

function Consultar()
{
	with (document.forms[0])
	{
		if (Cbo_OrisolAprov.value == "" )
		{
			alert("Selecione o Aprovisionador.");
			Cbo_OrisolAprov.focus();
			return;
		}
		
		/*if (cboAcao.value == "" )
		{
			alert("Selecione uma Ação.");
			cboAcao.focus();
			return;
		}*/
		
		if (Cbo_OrisolAprov.value == "6" || Cbo_OrisolAprov.value == "7")
		{
		if (txt_oe_numero.value == "")
		{
			alert("Informe o Número da Order Entry.");
			txt_oe_numero.focus();
			return;
		}
	
		if (txt_oe_ano.value == "")
		{
			alert("Informe o Ano da Order Entry.");
			txt_oe_ano.focus();
			return;
		}
		
		if (txt_oe_item.value == "")
		{
			alert("Informe o Item da Order Entry.");
			txt_oe_item.focus();
			return;
		}
		}
		
		if (Cbo_OrisolAprov.value == "10" )
		{
			if (txt_variavel.value == "" || txt_ss.value == "" || txt_num_sol.value == "" || txt_ano_sol.value == "")
			{
				alert("Informe a Solicitação CFD.");
				txt_variavel.focus();
				return;
			}
		
			
		}


		if (Cbo_OrisolAprov.value == "6")
		{
			if (txt_designacao.value == "")
			{
				alert("Informe a Designação.");
				txt_designacao.focus();
				return;
			}
		}
							
		action = "ConsultaClaAprovisionador.asp?Consulta=1"
		submit()
	}
}

function CompletarCampoIA(obj)
{
	//alert(obj.value)
	if (obj.value != "" && obj.value != 0 )
	{
		var intLen = parseInt(obj.size) - parseInt(obj.value.length)
	
		switch (obj.TIPO.toUpperCase())
		{
			case "N":
				for (var intIndex=0;intIndex<intLen;intIndex++)
				{
					obj.value = "0" + obj.value
				}
				break
			default :
				for (var intIndex=0;intIndex<intLen;intIndex++)
				{
					obj.value = obj.value + " "
					//alert(obj.value)
				}
		}
	}	
}
</script>
<form action="ConsultaClaAprovisionador.asp" name="Form1" method="post">
<input type=hidden name=hdnSolId>
<input type=hidden name=hdnAcao>
<%
'response.write Date()
'if Date() < "01/11/2013" and (strLoginRede<>"JCARTUS" and strLoginRede<>"SCESAR" and strLoginRede<>"MSCAPRI") then
'	response.write "<br><p align=center><b><font color=red><I>FUNCIONALIDADE AINDA NÃO DISPONÍVEL PARA USO.</I><br> <br>(Previsão 01/Nov/13)</font></b></p>"
'	response.end
'end if
%>
<table border=0 cellspacing="1" cellpadding="0" width="760" >
<tr >
	<th colspan=2 ><table width="760"><tr><th><center>Consulta de Solicitação de Acesso - Aprovisionador</th>
			<th width="26">
			
			</th></tr></table>
		</th>
</tr>
 
<tr class=clsSilver>
	<td>Aprovisionador</td>
	<td>
	    <select name="Cbo_OrisolAprov">
		  <option value=""></option> 
		  <%
		  set rsCboOrisolAprov = db.execute("select Orisol_ID,OriSol_Alias from CLA_OrigemSolicitacao where Orisol_InterfAprov = 1 order by orisol_alias ")
		  do while not rsCboOrisolAprov.eof
		 ' response.write "<script>alert('rsCboOrisolAprov=" & rsCboOrisolAprov("Orisol_ID") & "')</script>" 
		  %>
			<option value="<%=rsCboOrisolAprov("Orisol_ID")%>"><%=ucase(rsCboOrisolAprov("OriSol_Alias"))%></option>
		  <%
			rsCboOrisolAprov.movenext
			loop
		  %>
		</select>
	</td>
</tr>
<tr class=clsSilver>
	<td>Acao</td>
	<td>
		<select name="cboAcao">
			  <option value="" > Selecione uma Ação </option>
			  <option value="ATV" <%if request("cboAcao") = "ATV" then%>selected<%end if%>>ATIVAÇÃO</option>
			  <option value="ALT" <%if request("cboAcao") = "ALT" then%>selected<%end if%>>ALTERACAO</option>
			  <option value="DES" <%if request("cboAcao") = "DES" then%>selected<%end if%>>DESATIVAÇÃO</option>
		</select>
	</td>
</tr>

<tr class=clsSilver>
	<td>OE</td>
	<td>
        <input id="txt_oe_numero" type="text" title="Número" maxlength="10" size="11" 
            class=text onKeyUp="ValidarTipo(this,0)" 
            name="txt_oe_numero" value='<%=request("txt_oe_numero")%>'>&nbsp;/
        <input id="txt_oe_ano" type="text" title="Ano" maxlength="4" size="5" 
            class=text onKeyUp="ValidarTipo(this,0)" name="txt_oe_ano" 
            value='<%=request("txt_oe_ano")%>'>&nbsp;item&nbsp;
        <input id="txt_oe_item" type="text" title="Item" maxlength="3" size="4" 
            class=text onKeyUp="ValidarTipo(this,0)" name="txt_oe_item" 
            value='<%=request("txt_oe_item")%>'>
	</td>
</tr>
<tr class=clsSilver>
	<td>SOLICITACAO - CFD</td>
	<td>
        <input id="txt_variavel" type="text" title="XXXXXXXX" maxlength="8" size="8" 
            class=text name="txt_variavel" value='<%=request("txt_variavel")%>' onblur="CompletarCampoIA(this)" TIPO="A">&nbsp;
        <input id="txt_ss" type="text" title="IA" maxlength="2" size="2" 
            class=text name="txt_ss"         value='<%=request("txt_ss")%>' onblur="CompletarCampoIA(this)" TIPO="A">&nbsp;
		<input id="txt_num_sol" type="text" title="Numero" maxlength="4" size="4" 
            class=text name="txt_num_sol"    onKeyUp="ValidarTipo(this,0)"     value='<%=request("txt_num_sol")%>' onblur="CompletarCampoIA(this)" TIPO="N">&nbsp;/
		<input id="txt_ano_sol" type="text" title="Ano" maxlength="4" size="4" 
            class=text name="txt_ano_sol"      onKeyUp="ValidarTipo(this,0)"   value='<%=request("txt_ano_sol")%>' onblur="CompletarCampoIA(this)">&nbsp;
        
	</td>
</tr>
 
<tr class=clsSilver>
	<td>Designação</td>
	<td>
	    <input id="txt_designacao" type="text" class="text" maxlength="60" 
            name="txt_designacao" size="40"  value='<%=request("txt_designacao")%>'></td>
	 
</tr>

<tr>
	<td colspan=2 align=center height=35px>
		<input type="button" name="btconsulta" value="Consultar" class="button" onClick="Consultar()">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px">
        </td>
</tr>
 
</table>
 
 

<%
if Request.QueryString ("Consulta") = "1"  then
	 
	if request("Cbo_OrisolAprov") <>  "10" then
		strSql = " select top 1 acl_designacaoservico,id_tarefa_can,Aprov_Enviado,orisol_descricao,data_recebimento,CLA_Aprovisionador.Sol_ID,Sts_Desc,OriSol_Descricao,OE_NUMERO,OE_ANO,OE_ITEM,Acao, id_endereco , " &_
			" Cli_Nome,Cli_CC,Cli_SubCC,Ser_Sigla,Ser_Desc,CLA_Aprovisionador.Acl_IDAcessoLogico,Aprov_Cancelado,aprovisi_dtCancAuto, sol_idcan , CLA_Aprovisionador.IA , CLA_Aprovisionador.Orisol_id from " &_
		"CLA_Aprovisionador (nolock) left join cla_solicitacao  on CLA_Aprovisionador.Sol_ID = cla_solicitacao.Sol_ID left join cla_statussolicitacao " &_
		"on cla_solicitacao.stssol_id = cla_statussolicitacao.stssol_id left join cla_status on cla_statussolicitacao.sts_id = cla_status.sts_id " &_
			" Where Acao <> 'CAN' "
	else
		strSql = "select top 2 acl_designacaoservico,id_tarefa_can,Aprov_Enviado,orisol_descricao,data_recebimento,CLA_Aprovisionador.Sol_ID,Sts_Desc,OriSol_Descricao,OE_NUMERO,OE_ANO,OE_ITEM,Acao,id_endereco, " &_
			" Cli_Nome,Cli_CC,Cli_SubCC,Ser_Sigla,Ser_Desc,CLA_Aprovisionador.Acl_IDAcessoLogico,Aprov_Cancelado,aprovisi_dtCancAuto, sol_idcan , CLA_Aprovisionador.IA , CLA_Aprovisionador.Orisol_id from " &_
			" CLA_Aprovisionador (nolock) left join cla_solicitacao  on CLA_Aprovisionador.Sol_ID = cla_solicitacao.Sol_ID left join cla_statussolicitacao " &_
			" on cla_solicitacao.stssol_id = cla_statussolicitacao.stssol_id left join cla_status on cla_statussolicitacao.sts_id = cla_status.sts_id " &_
			" Where Acao <> 'CAN' "
	
	end if 
		
	if request("Cbo_OrisolAprov") <> "" then
		strSql = strSql & "and CLA_Aprovisionador.orisol_id = " & request("Cbo_OrisolAprov")
	end if 
	
	if request("txt_oe_numero") <> "" then
		strSql = strSql & " and OE_NUMERO =" & request("txt_oe_numero") 
	end if 
	
	if request("txt_oe_ano") <> "" then
		strSql = strSql & "  and OE_ANO=" & request("txt_oe_ano") 
	end if
	
	if request("txt_oe_item") <> "" then
		strSql = strSql & " and oe_item=" & request("txt_oe_item") 
	end if 
	
	if request("txt_variavel") <> "" then
		txt_ia = request("txt_variavel") + request("txt_ss") + request("txt_num_sol") + request("txt_ano_sol")
	
		strSql = strSql & " and ia='" & txt_ia & "'"
	end if
	
	if request("cboAcao") <> "" then
		strSql = strSql & " and acao='" & Request("cboAcao") & "'"
	end if
	
	if request("txt_designacao") <> "" then
		strSql = strSql & " and acl_designacaoservico='" & Request("txt_designacao") & "'"
	end if
		
	strSql = strSql & " Order by Aprovisi_Id desc"
	'response.write strSql
	Set ObjRs = db.execute(strSql)
	'response.write "***Cli_Nome=" 
'	 response.write rs_lista("Cli_Nome")
		
	'''		Set objRSCli = db.execute("CLA_sp_sel_solucao_ssa " & strNroSev & ",0")
	'''		if Not objRSCli.eof and Not objRSCli.bof then
	'''			strSolSel = "<table border=0 cellspacing=1 cellpadding=1><tr class=clsSilver2><td>Provedor</td><td>Facilidade</td><td>Prazo</td></tr>"
				'Soluções indicadas pelo SSA
	'''			While Not objRSCli.eof
	'''				if Trim(objRSCli("Sol_Selecionada")) = 1 then
	'''					strSolSel = strSolSel & "<tr><td>" & Trim(objRSCli("For_Des")) & "</td><td bgcolor=#f2f2f2>" & Trim(objRSCli("Fac_Des")) & "</td><td bgcolor=#f2f2f2>" & Trim(objRSCli("Sol_PrazoCompleto")) & "</td></tr>"
	'''				End if
	'''				objRSCli.MoveNext
	'''			Wend
	'''			strSolSel = strSolSel & "</table>"
	'''			strRespostaSSA = strSolSel
	'''		End if
	'''	End if
	'''End if

	if not ObjRs.Eof and not ObjRs.Bof then
		While not ObjRs.Eof
	
		
		If trim(ObjRs("Acao")) = "ATV" Then StrAcao = "ATIVAÇÃO " 	End if 
		If trim(ObjRs("Acao")) = "ALT" Then StrAcao = "ALTERAÇÃO " 	End if 			
		'If trim(rs_lista("Acao")) = "CAN" Then StrAcao = "CANCELAMENTO " End if 
		If trim(ObjRs("Acao")) = "DES" Then StrAcao = "DESATIVAÇÃO " 	End if
		
		
		'if trim(rs_lista("Acao")) = "CAN" then
		'	strSql = "select top 1 * from CLA_Aprovisionador where id_tarefa=" & rs_lista("id_tarefa_can") & " and aprovisi_id < 244028 order by aprovisi_id desc"
		'end if
		
		
%>			
		
		<table border=0 cellspacing="1" cellpadding="0" width="760">
			<% IF ObjRs("Orisol_id")  <> "10" then %>
				<tr><td height=20><th>&nbsp;N. OE/Ano: </th></td>                     <td class=clsSilver>&nbsp;<%=ObjRs("OE_NUMERO")%>/<%=ObjRs("OE_ANO")%></td></TR>			
				<TR><td height=20><th>&nbsp;Item da OE: </th></td>                    <td class=clsSilver>&nbsp;<%=ObjRs("OE_ITEM")%></TD></tr>
			<% ELSE %>
				<TR><td height=20><th>&nbsp;SOLICITACAO - CFD: </th></td>             <td class=clsSilver>&nbsp;<%=ObjRs("IA")%></TD></tr>
					
			<% END IF %>
			<TR><td height=20><th>&nbsp;Ponta : </th></td>                  <td class=clsSilver>&nbsp;<%=ObjRs("id_endereco")%></td></TR>						 			 
			<TR><td height=20><th>&nbsp;Nome Cliente: </th></td>                  <td class=clsSilver>&nbsp;<%=ObjRs("Cli_Nome")%></td></TR>						 			 
			<tr><td height=20><th>&nbsp;Numero da Solicitação:&nbsp; </th></td>   <td class=clsSilver>&nbsp;<a href="javascript:DetalharItem(<%=ObjRs("sol_id")%>)"><%=ObjRs("sol_id")%></a></td></TR>			
			<TR><td height=20><th>&nbsp;Ação: </th></td>                          <td class=clsSilver>&nbsp;<%=StrAcao%></td></tr>
 			<TR><td height=20><th>&nbsp;Status Macro: </th></td>                  <td class=clsSilver>&nbsp;<%=ObjRs("sts_desc")%></td></TR>						 						 
			<tr><td height=20><th>&nbsp;ID-Logico: </th></td>                     <td class=clsSilver>&nbsp;<%=ObjRs("Acl_IDAcessoLogico")%></td></TR>			
			<TR><td height=20><th>&nbsp;Data Receb. da OE no CLA:&nbsp; </th></td><td class=clsSilver>&nbsp;<%=ObjRs("data_recebimento")%></td></tr>
			<%If trim(ObjRs("Acao")) <> "DES" Then	%>		
				<tr><td height=20><th>&nbsp;Cancelado: </th></td>                     <td class=clsSilver>&nbsp;<%=ObjRs("Aprov_Cancelado")%></td></TR>			
				<TR><td height=20><th>&nbsp;Solicitação de Cancelamento: </th></td>   <td class=clsSilver>&nbsp;<%=ObjRs("sol_idcan")%></td></tr>
				<tr><td height=20><th>&nbsp;Data de Cancelamento:&nbsp; </th></td>    <td class=clsSilver>&nbsp;<%=ObjRs("aprovisi_dtCancAuto")%>&nbsp;</td></TR>			
<%End If%>			
			<TR><td height=20><th>&nbsp;Aprovisionador: </th></td>                <td class=clsSilver>&nbsp;<%=ObjRs("orisol_descricao")%></td></tr>
			<TR><td height=20><th>&nbsp;Designação: </th></td>                <td class=clsSilver>&nbsp;<%=ObjRs("acl_designacaoservico")%></td></tr>
		</table>
		
<%

'response.write "Aprov_Cancelado=" & rs_lista("Aprov_Cancelado")
'response.write "Aprov_Enviado=" & rs_lista("Aprov_Enviado") 
'if (rs_lista("Aprov_Cancelado") is null or rs_lista("Aprov_Cancelado")<>"S") and (rs_lista("Aprov_Enviado") is null or trim(rs_lista("Aprov_Enviado"))="") then 
If isNull(ObjRs("Aprov_Cancelado")) and isNull(ObjRs("Aprov_Enviado")) Then 
	response.write "<table width= 760 border= 0 cellspacing= 0 cellpadding= 0 valign=top><tr><td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;<I><B>OE AGUARDANDO CRIAÇÃO DE SOLICITAÇÃO DE ACESSO NO CLA_APROVISIONADOR</B></I>.</font></td></tr></table>	"
End If
		
	
	
	'set rs_lista = Nothing
	'DesconectarCla()

	ObjRs.MoveNext
	Wend
	
		DesconectarCla()
	
	Else
	%>
		<table width= 760 border= 0 cellspacing= 0 cellpadding= 0 valign=top><tr><td align=center valign=center width=100% height=20 ><font color=red>&nbsp;•&nbsp;Registro(s) não encontrado(s).</font></td></tr></table>	
<%	
	end if
	
end if
%>

</body>
</html>
 