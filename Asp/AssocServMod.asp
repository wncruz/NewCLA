<%

'	- Sistema			: CLA
'	- Arquivo			: AssocServMod.asp
'	- Responsável		: Vital
'	- Descrição			: Associaçção de serviço com modalidade
%>
<!--#include file="../inc/data.asp"-->
<%
Dim dblID
Dim objRSVel
Dim strSel
Dim dblIDAtual

dblID = Server.HTMLEncode(Request.QueryString("ID"))
if Trim(dblID) = "" then
	dblID = Request.Form("hdnId")
End if	

If Request.Form("hdnAcao")="Gravar" then

	If dblID="" then
		Vetor_Campos(1)="adInteger,2,adParamInput,"
	Else
		Vetor_Campos(1)="adInteger,2,adParamInput,"& dblID
	End if

	'response.write "<script>alert('"&Request.Form("cboServico")&"')</script>"
	'response.write "<script>alert('"&Request.Form("cboMod")&"')</script>"
	
	
	Vetor_Campos(2)="adInteger,2,adParamInput,"& Server.HTMLEncode(Request.Form("cboServico")) 
	Vetor_Campos(3)="adInteger,2,adParamInput," & Server.HTMLEncode(Request.Form("cboMod")) 

	Vetor_Campos(4)="adInteger,2,adParamOutput,0"  
	
	Call APENDA_PARAM("CLA_sp_ins_AssocServMod",4,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value

End if

If dblID <> "" then
	Set objRSVel = db.execute("CLA_sp_sel_AssocServMod " & dblID)
End if
%>
<!--#include file="../inc/header.asp"-->
<form action="AssocServMod.asp" method="post" >
<SCRIPT LANGUAGE="JavaScript">
function checa(f) 
{
	if (!ValidarCampos(f.cboServico,"Serviço")) return false;
	if (!ValidarCampos(f.cboMod,"A Modalidade")) return false;
	return true;
}

function GravarServMod()
{
	if (!checa(document.forms[0])) return
	with (document.forms[0])
	{
		action = "AssocServMod.asp"
		hdnAcao.value = "Gravar"
		submit()
	}
}
</script>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnUFAtual>
<input type=hidden name=hdnId value="<%=dblID%>" >
<tr>
	<td>
		<table border="0" cellspacing="1" cellpadding=0 width="760">
			<tr>
				<th colspan=2><p align="center">Associação de Serviço com modalidade</p></th>
			</tr>
			<tr class=clsSilver>
				<td>
					<font class="clsObrig">:: </font>Serviço
				</td>
				<td>
					<select name="cboServico">
						<option value=""></option>
						<% set objRS = db.execute("CLA_sp_sel_servico null,null,null,1")
							If Trim(dblID)<> "" then
								dblIDAtual = objRSVel("Ser_id")
							Else
								dblIDAtual = Server.HTMLEncode(Request.Form("cboServico")) 
							End if

							While Not objRS.Eof
								strSel = ""
								if Cdbl("0" & objRS("Ser_id")) = Cdbl("0" & dblIDAtual) then strSel = " selected "
								Response.Write "<Option value="& Trim(objRS("Ser_id")) & strSel & ">" & Trim(objRS("Ser_Desc")) & "</Option>"
								objRS.MoveNext
							Wend
							Set objRS = Nothing
						%>
					</select>
				</td>
			</tr>

			<tr class=clsSilver>
				<td>
					<font class="clsObrig">:: </font>Modalidade
				</td>
				<td>
					<select name="cboMod">
						<option value=""></option>
						<% set objRS = db.execute("CLA_sp_sel_modalidade") 
							If Trim(dblID)<> "" then
								dblIDAtual = objRSVel("Mod_id")
							Else
								dblIDAtual = Server.HTMLEncode(Request.Form("cboMod")) 
							End if

							While Not objRS.Eof
								strSel = ""
								if Cdbl("0" & objRS("Mod_id")) = Cdbl("0" & dblIDAtual) then strSel = " selected "
								Response.Write "<Option value="& objRS("Mod_id") & strSel & ">" & objRS("Mod_Desc") & "</Option>"
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
				<input type="button" class="button" name="btnGravar" value="Gravar" onClick="GravarServMod()" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
				<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnId.value = '';LimparForm();setarFocus('cboServico');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
				<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('AssocServMod_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');" >
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
setarFocus('cboServico');
//-->
</SCRIPT>

</html>
<%
Set objRSVel = Nothing
DesconectarCla()
%>
