<%
'	- Sistema			: CLA
'	- Arquivo			: estruturaCombo.asp
'	- Respons?vel		: EDAR
'	- Descri??o			: Associa??o de Combo
%>
<!--#include file="../inc/data.asp"-->
<%
Dim dblID
Dim objRSVel
Dim strSel
Dim dblIDAtual

dblID = Request.QueryString("ID")
estrutura_combo_id = Request.QueryString("assoc")

'response.write "<script>alert('"&dblID&"')</script>"
'response.write "<script>alert('"& estrutura_combo_id&"')</script>"
'response.write "<script>alert('"&assoc_tecfac_id &"')</script>"

'if Trim(dblID) = "" then
	'dblID = Request.Form("hdnId")
'End if	

If Request.Form("hdnAcao")="Gravar" then

	'response.write "<script>alert('"& Request.Form("hdnId") &"')</script>"
	'response.write "<script>alert('"& Request.Form("hdnestrutura") &"')</script>"
	If Request.Form("hdnestrutura") <> "" then
		'response.write "<script>alert('"& Request.Form("hdnestrutura") &"')</script>"
		Vetor_Campos(1)="adInteger,2,adParamInput,"& Request.Form("hdnestrutura") 'Request.Form("hdnassoc_tecfac_id") 
		Vetor_Campos(2)="adInteger,2,adParamInput," & Request.Form("hdnId")	
		 		 
	Else
		'response.write "<script>alert('"& Request.Form("hdnestrutura")&"')</script>"
		Vetor_Campos(1)="adInteger,2,adParamInput,"
		Vetor_Campos(2)="adInteger,2,adParamInput,"& Request.Form("hdnId")
	End if

	'response.write "<script>alert('"&Request.Form("hdnId")&"')</script>"
	'response.write "<script>alert('"&Request.Form("hdnassoc_tecfac_id")&"')</script>"

	
	
	
	Vetor_Campos(3)="adWChar,100,adParamInput," & Request.Form("txtLabel") 
	Vetor_Campos(4)="adWChar,5,adParamInput," & Request.Form("rdoStatus") 
	Vetor_Campos(5)="adWChar,10,adParamInput," & strloginrede 
	Vetor_Campos(6)="adInteger,2,adParamOutput,0"  
	
	Call APENDA_PARAM("CLA_sp_ins_estruturaCombo",6,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
	'response.write "<script>alert('"&DBAction &"')</script>"

	dblID  = Request.Form("hdnId")

	estrutura_combo_id  = Request.Form("hdnestrutura")
End if

If dblID <> "" then
	'Set objRS = db.execute("CLA_sp_sel_assocTecnologiaFacilidade null , " & assoc_tecfac_id )

	Set objRS = db.execute("CLA_sp_sel_EstruturaCombo null, " & dblID)


	newCombo_nome 	= Trim(objRS("newCombo_nome"))  
	

	'label 		= Trim(objRS("label"))
	'formato_des	= Trim(objRS("formato"))
	'Tamanho 	= Trim(objRS("Tamanho"))

	'estrutura_combo_id = Trim(objRS("estrutura_combo_id"))

	
	'response.write "<script>alert('"&formato_des 	&"')</script>"
	'response.write "<script>alert('"&Tamanho&"')</script>"

	Set objRS  = Nothing
end if

If estrutura_combo_id <> "" then 

	
	
	
	Set objRS = db.execute("CLA_sp_sel_EstruturaCombo null, null , " & estrutura_combo_id )


	newcombo_nome 	= Trim(objRS("newcombo_nome"))  
	
	'response.write newcombo_nome
	'response.end
	

	label 		= Trim(objRS("label"))
	
	status		= Trim(objRS("status"))
	

	
	


	estrutura_combo_id = Trim(objRS("estrutura_combo_id"))
	dblID = Trim(objRS("newcombo_id"))

	'estrutura 	= Trim(objRS("estrutura_tec_fac_id"))

	'response.write "<script>alert('"&formato	 	 	&"')</script>"
	'response.write "<script>alert('"&Tamanho 	&"')</script>"

	Set objRS  = Nothing
	
End if
%>
<!--#include file="../inc/header.asp"-->
<form action="AssocTecnologiaFacilidade.asp" method="post" >
<SCRIPT LANGUAGE="JavaScript">
function checa(f) 
{
	if (!ValidarCampos(f.txtLabel,"Label")) return false;
	
	if (getCheckedRadioValue(f.rdoStatus)=="")
	{
		alert (" Status ? um campo obrigat?rio.");
		return false;
	}
	

	return true;
}

function GravarEstruturaCombo()
{
	if (!checa(document.forms[0])) return
	with (document.forms[0])
	{
		action = "estruturaCombo.asp"
		hdnAcao.value = "Gravar"
		submit()
	}
}
</script>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnId value="<%=dblID%>" >
<input type=hidden name=hdncombo_id value="<%=dblID%>" >
<input type=hidden name=hdnestrutura  value="<%=estrutura_combo_id%>" >

<tr>
	<td>
		<table border="0" cellspacing="1" cellpadding=0 width="760">
			<tr>
				<th colspan=2><p align="center">Estrutura de Combo</p></th>
			</tr>
			<tr class=clsSilver>
				<td>
					<font class="clsObrig">:: </font>Combo </font>
				</td>
				
				<td>
					
					<%= newCombo_nome   %>
				
					
				</td>
			</tr>
			
			<tr class=clsSilver>
				<td ><font class="clsObrig">:: </font>Label&nbsp;&nbsp;</td>
				<td ><input type=text name=txtLabel maxlength=100 size = 100 class="text" value="<%=Label%>">&nbsp;
				
			</tr>
			
			
			<tr class=clsSilver>
				<td width=150px ><font class="clsObrig">:: </font>STATUS</td>
				<td>
					<input type=radio value=S name=rdoStatus <%if status		= "S" then response.write "checked"%> >&nbsp;ATIVO&nbsp;&nbsp;
					<input type=radio value=N name=rdoStatus <%if status		= "N" then response.write "checked"%> >&nbsp;INATIVO
				
					
				</td>
			</tr>
			
		</table>
		<table width="760" border=0>
		<tr>
			<td colspan=2 align="center"><br>
				<input type="button" class="button" name="btnGravar" value="Gravar" onClick="GravarEstruturaCombo()" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
				<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnId.value = '';LimparForm();setarFocus('txtLabel');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
				<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('cst_estruturaCombo.asp?ID=<%=dblID%>')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');" >
				<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
			</td>
		</tr>
		</table>
		<table width="760" border=0>
		<tr>
			<td>
				<font class="clsObrig">:: </font> Campos de preenchimento obrigat?rio.
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</body>
<SCRIPT LANGUAGE=javascript>

</SCRIPT>

</html>
<%
Set objRSAssocTecFac = Nothing
DesconectarCla()
%>
