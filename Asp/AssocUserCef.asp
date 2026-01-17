<!--#include file="../inc/data.asp"-->

<%
Dim dblID
Dim objRSCef

dblID = Request.QueryString("ID")
if Trim(dblID) = "" then
	 dblID  = Request.Form("hdnID")
End if	 

If Server.HTMLEncode(request("action"))="Gravar" then

	If dblID="" then
		Vetor_Campos(1)="adInteger,2,adParamInput,"
	Else
		Vetor_Campos(1)="adInteger,2,adParamInput,"& dblID
	End if
	Vetor_Campos(2)="adWChar,30,adParamInput,"& ucase(Server.HTMLEncode(request("txtUserName")))
	Vetor_Campos(3)="adInteger,2,adParamInput," & Server.HTMLEncode(Request.Form("cboCef")) 
	Vetor_Campos(4)="adWChar,10,adParamInput," & strloginrede '-->PSOUTO 20/04/06
	Vetor_Campos(5)="adInteger,2,adParamOutput,0"  
	
	Call APENDA_PARAM("CLA_sp_ins_usuarioctfc",5,Vetor_Campos)
	' /PSOUTO
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value

End if


if dblID <> "" then
	set objRSCef = db.execute("CLA_sp_sel_usuarioctfc " & dblID)
end if
%>
<!--#include file="../inc/header.asp"-->
<form action="AssocUserCef.asp" method="post" onSubmit="return checa(this)">
<SCRIPT LANGUAGE="JavaScript">
function checa(f) {

	if (!ValidarCampos(f.txtUserName,"Usuário(UserName)")) return false;
	if (!ValidarCampos(f.cboCef,"O Centro funcional")) return false;

	return true;
}

</script>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnID value="<%=dblID%>" >

<tr>
	<td >
		<table border="0" cellspacing="1" cellpadding=0 width="760">
			<tr>
				<th colspan=2><p align=center>Associação de Usuário com Centro Funcional</p></td>
			</tr>

			<tr class=clsSilver>
				<td width=25% >
					<font class="clsObrig">:: </font>Usuário (Username)
				</td>
				<td><input type=text class=text name=txtUserName size=30 maxlength=30 value="<%If Trim(dblID)<> "" then Response.Write TratarAspasHtml(objRSCef("Usu_UserName")) else Response.Write TratarAspasHtml(Request.Form("txtUserName"))%>" >
				</td>
			</tr>

			<tr class=clsSilver>
				<td>
					<font class="clsObrig">:: </font>Centro Funcional
				</td>
				<td>
					<select name="cboCef">
						<option value=""></option>
						<% 
							Dim dblIDAtual
							Dim strSel
							
							set objRS = db.execute("CLA_sp_sel_centrofuncionalFull ")
							If Trim(dblID)<> "" then
								dblIDAtual = objRSCef("Ctfc_id")
							Else
								dblIDAtual = Request.Form("cboCef") 
							End if

							While Not objRS.Eof
								strSel = ""
								if Cdbl("0" & objRS("Ctfc_id")) = Cdbl("0" & dblIDAtual) then strSel = " selected "
								Response.Write "<Option value="& objRS("Ctfc_id") & strSel & ">" & objRS("Ctf_AreaFuncional") & " - " & objRS("Cid_Sigla") & " "  & objRS("Esc_Sigla") & " - " & objRS("Age_Sigla") & " - " & objRS("Age_Desc") & "</Option>"
								objRS.MoveNext
							Wend
							Set objRS = Nothing
						%>
					</select>
				</td>
			</tr>

		</table>
		<table width="760" border=0>
		<tr>
			<td colspan=2 align="center"><br>
				<input type="submit" class="button" name="action" value="Gravar" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
				<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnID.value='';LimparForm();setarFocus('txtUserName');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
				<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('AssocUserCef_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
				<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
			</td>
		</tr>
		</table>
		<table width="760" border=0>
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
<SCRIPT LANGUAGE=javascript>
<!--
setarFocus('txtUserName')
//-->
</SCRIPT>

</html>
<%
Set objRSCef = Nothing
DesconectarCla()
%>
