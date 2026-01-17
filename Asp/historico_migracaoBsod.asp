<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
Dim sAcao
Dim sAcf_IdAcessoFisico  
sAcao = Request("acao")
sAcf_IdAcessoFisico = Request("txtAcf_IdAcessoFisico") '"BRE 00000002688"'
 
%>
<script language="Javascript">

function trim(str) {
	return str.replace(/^\s+|\s+$/g,"");
}

function Migrar()
{
		if (trim(document.getElementById("txtFac_Vlan").value)==""){
			alert("Informar a CVLAN.")
			document.getElementById("txtFac_Vlan").focus()
			return
		}
		if (trim(document.getElementById("txtFac_Porta").value)==""){
			alert("Informar a PORTA.")
			document.getElementById("txtFac_Porta").focus()
			return
		}
		if (trim(document.getElementById("txtFac_PE").value)==""){
			alert("Informar o PE.")
			document.getElementById("txtFac_PE").focus()
			return
		}		
		if (trim(document.getElementById("txtFac_TipoPorta").value)==""){
			alert("Informar o Tipo Porta.")
			document.getElementById("txtFac_TipoPorta").focus()
			return
		}		
				
		if (confirm("Deseja realmente migrar o ID Acesso Físico '" + document.getElementById("txtAcf_IdAcessoFisico").value + "' para HFC BSoD ?")){      
				with(document.forms[0])
				{
					hdnAcao.value = "MigrarBSoD"			 
					target = "IFrmProcesso"					 
					action = "ProcessoCla.asp"
					submit() 
				}
		} 
}

function Procurar()
{
		if (trim(document.getElementById("txtAcf_IdAcessoFisico").value)==""){
			alert("Informar o ID Acesso Físico.")
			document.getElementById("txtAcf_IdAcessoFisico").focus()
			return
		}
	with (document.forms[0])
{ 
		action = "historico_migracaoBsod.asp?acao=consulta"
		submit()
	}
}
</script>

<html>
<body>
<form name=Form1 method="post" >
<input type=hidden id=hdnAcao name=hdnAcao>
<table border=0 cellspacing="1" cellpadding = 0 width="760" >
<tr class=clsSilver>
	<th colspan=2><p align=center>Histórico - Migração de acesso físico HFC BSoD VIA e HFC BSoD NET para HFC BSoD</p></th>
</tr>
<tr class=clsSilver>
	<td width=25% height=30>&nbsp;ID Acesso Físico</td>
	<td><input type="text" id="txtAcf_IdAcessoFisico" name="txtAcf_IdAcessoFisico" value="<%=sAcf_IdAcessoFisico%>" size=18 class=text  maxlength=15></td>
</tr>
<tr>
	<td colspan=2 align="center" height=30px >
		<input type="button" class="button" name="btnProcurar" value="Procurar" style="width:100px" onclick="Procurar()">&nbsp;
		<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
	</td>
</tr>
</table>
<br>
<%
If sAcao = "consulta" Then
	If IsNULL(sAcf_IdAcessoFisico) or trim(sAcf_IdAcessoFisico) = "" Then		 
		response.write "<script>alert('Favor informar o ID Acesso Físico.')</script>" 		 
	Else	
		Set objRSMigr = db.execute("select top 1 * from CLA_LogMigracaoBSoD where acf_idacessofisico='" & sAcf_IdAcessoFisico & "'")
		
		If objRSMigr.eof Then
			response.write "<script>alert('ID Acesso Físico (" & sAcf_IdAcessoFisico & ") não encontrado.');document.getElementById('txtAcf_IdAcessoFisico').focus()</script>" 
		Else	
%>
<table border=0 cellspacing="1" cellpadding = 0 width="760" >
<tr class=clsSilver>
	<th colspan=2><p align=center>Histórico - Migração de acesso físico HFC BSoD VIA e HFC BSoD NET para HFC BSoD</p></th>
</tr>
		<tr class="clsSilver">
		<td width="200" height=20>&nbsp;&nbsp;&nbsp;&nbsp;Id Acesso Lógico</td><td width="570">&nbsp;<%=objRSMigr("acl_idacessologico")%></td>
 		</tr>
		<tr class="clsSilver">
		<td width="200" height=20>&nbsp;&nbsp;&nbsp;&nbsp;Nº Acesso (Contrato NET/VIA)</td><td>&nbsp;<%=objRSMigr("Acf_NroAcessoPtaEbt")%></td>
 		</tr> 		
		<tr class="clsSilver">
		<td width="200" height=20>&nbsp;&nbsp;&nbsp;&nbsp;ID Físico</td><td>&nbsp;<%=sAcf_IdAcessoFisico%></td>
 		</tr> 	 		
 		<tr class="clsSilver">
		<td width="200" height=20>&nbsp;&nbsp;&nbsp;&nbsp;Data da Migração</td><td>&nbsp;<%=objRSMigr("Dt_Migracao")%></td>
 		</tr> 	
 		<tr class="clsSilver">
		<td width="200" height=20>&nbsp;&nbsp;&nbsp;&nbsp;Responsável p/ Migração</td><td>&nbsp;<%=objRSMigr("Username_migrador")%></td>
 		</tr> 	 		
 		<tr class="clsSilver">
		<th colspan=2>&nbsp;&nbsp;&nbsp;&nbsp;Informações <u>Anteriores</u> à Migração</th>
 		</tr>  	
 		<tr class="clsSilver">
		<td width="200" height=20>&nbsp;&nbsp;&nbsp;&nbsp;Provedor</td><td>&nbsp;<%=objRSMigr("Pro_Nome_Old")%></td>
 		</tr> 	 			
 		
 		<tr class="clsSilver">
		<td width="200" height=20>&nbsp;&nbsp;&nbsp;&nbsp;CVLAN</td><td>&nbsp;<%=objRSMigr("Fac_Vlan_Old")%></td>
 		</tr> 
 		<tr class="clsSilver">
		<td width="200" height=20>&nbsp;&nbsp;&nbsp;&nbsp;PORTA</td><td>&nbsp;<%=objRSMigr("Fac_Porta_Old")%></td>
 		</tr> 
 		<tr class="clsSilver">
		<td width="200" height=20>&nbsp;&nbsp;&nbsp;&nbsp;PE</td><td>&nbsp;<%=objRSMigr("Fac_PE_Old")%></td>
 		</tr> 
 		<tr class="clsSilver">
		<td width="200" height=20>&nbsp;&nbsp;&nbsp;&nbsp;SVLAN</td><td>&nbsp;<%=objRSMigr("Fac_SVlan_Old")%></td>
 		</tr> 
 		 		 		 		 		
 		 <tr class="clsSilver">
		<th colspan=2>&nbsp;&nbsp;&nbsp;&nbsp;Informações <u>Posteriores</u> à Migração</th>
 		</tr>  
 		<tr class="clsSilver">
		<td width="200" height=20>&nbsp;&nbsp;&nbsp;&nbsp;Provedor</td><td>&nbsp;<%=objRSMigr("Pro_Nome_New")%></td>
 		</tr> 	 	 		
 		<tr class="clsSilver">
		<td width="200" height=20>&nbsp;&nbsp;&nbsp;&nbsp;CVLAN</td><td>&nbsp;<%=objRSMigr("Fac_Vlan_New")%></td>
 		</tr> 
 		<tr class="clsSilver">
		<td width="200" height=20>&nbsp;&nbsp;&nbsp;&nbsp;PORTA</td><td>&nbsp;<%=objRSMigr("Fac_Porta_New")%></td>
 		</tr> 
 		<tr class="clsSilver">
		<td width="200" height=20>&nbsp;&nbsp;&nbsp;&nbsp;PE</td><td>&nbsp;<%=objRSMigr("Fac_PE_New")%></td>
 		</tr> 
 		<tr class="clsSilver">
		<td width="200" height=20>&nbsp;&nbsp;&nbsp;&nbsp;SVLAN</td><td>&nbsp;<%=objRSMigr("Fac_SVlan_New")%></td>
 		</tr> 
 		<tr class="clsSilver">
		<td width="200" height=20>&nbsp;&nbsp;&nbsp;&nbsp;Tipo Porta</td>
		<td>
<%		
				Set objRSTP = db.Execute("select * from CLA_TipoPorta_HFCBSoD where Cod_tipoPorta='" & objRSMigr("Fac_TipoPortaNew") & "'")
				If Not objRSTP.eof Then
		   		Response.Write objRSTP("Desc_tipoPorta") 
		   	End If		
		   	
%>		   	
		</td>
 		</tr>  		
</table>	
<%
		End If
	End If
	 
End If
%>	
<iframe	id			= "IFrmProcesso"
		name        = "IFrmProcesso"
		width       = "0"
		height      = "0"
		frameborder = "0"
		scrolling   = "no"
		align       = "left">
</iFrame>
</form> 
</body>
</html>
