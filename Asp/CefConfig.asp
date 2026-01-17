<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: CefConfig.asp
'	- Responsável		: Vital
'	- Descrição			: Cadastra/Altera Holding
%>
<!--#include file="../inc/data.asp"-->
<%
Dim strDesc
Dim objRSCtf
Dim dblID

dblID = Request.QueryString("ID")
if Trim(dblID) = "" then
	 dblID  = Request.Form("hdnID")
End if	 

If Trim(Request.Form("btnGravar")) = "Gravar" then

	Vetor_Campos(1)="adInteger,2,adParamInput,"& ucase(request("cboCef"))
	Vetor_Campos(2)="adInteger,2,adParamInput,"& Cint("0" & request("chkRedirCarteira"))
	Vetor_Campos(3)="adInteger,2,adParamOutput,0"
	'@@JKNUP: Alterado. GRADE. Correção para view rádio.
	if Cint("0" & request("chkGerenAvaliador")) = 0 then
		Vetor_Campos(4)="adInteger,2,adParamInput,"
	Else
		Vetor_Campos(4)="adInteger,2,adParamInput,"& Cint("0" & request("chkGerenAvaliador"))	
	End if
	Vetor_Campos(5)="adWChar,10,adParamInput," & strloginrede '-->PSOUTO 12/04/06
	'</@@JKNUP>
	Call APENDA_PARAM("CLA_sp_ins_ConfigCtf ",5,Vetor_Campos)
	ObjCmd.Execute
	DBAction = ObjCmd.Parameters("RET").value

End if

If Trim(dblID) <> "" then
	set objRSCtf = db.execute("CLA_sp_sel_ConfigCtf " & dblID)
	If Not objRSCtf.Eof and Not objRSCtf.Bof then
		dblCefId = objRSCtf("Ctf_ID")
		strRedirCarteira = objRSCtf("Cfg_RedirecionamentoCarteira")
		strGerenAvaliador = objRSCtf("Cfg_Avaliador")
	End if
Else	
	dblCefId = Request.Form("cboCef") 
	strRedirCarteira = Request.Form("chkRedirCarteira") 
	strGerenAvaliador = Request.Form("chkGerenAvaliador") 
End if

%>
<!--#include file="../inc/header.asp"-->
<form action="CefConfig.asp" method="post" onSubmit="return checa(this)">
<input type=hidden name=hdnID value="<%=dblID%>" >
<SCRIPT LANGUAGE="JavaScript">
function checa(f) 
{
	if (!ValidarCampos(f.cboCef,"Centro Funcional")) return false;
	return true;
}
</script>
<tr>
	<td >
		<table border=0 cellspacing="1" cellpadding="0" width="760">
		<tr>
			<th colspan=2><p align="center">Parâmetros do Centro Funcional</p></th>
		</tr>
		<tr class=clsSilver>
			<td width=170>
				<font class="clsObrig">:: </font>Centro Funcional
			</td>
			<td>
				<select name="cboCef">
					<option value=""></option>
					<% 
						Dim dblIDAtual
						Dim strSel
							
						set objRS = db.execute("CLA_sp_sel_centrofuncional ")

						While Not objRS.Eof
							strSel = ""
							if Cdbl("0" & objRS("Ctf_id")) = Cdbl("0" & dblCefId) then strSel = " selected "
							Response.Write "<Option value="& objRS("Ctf_id") & strSel & ">" & objRS("Ctf_AreaFuncional") & " - " & objRS("Cid_Sigla") & "</Option>"
							objRS.MoveNext
						Wend
						Set objRS = Nothing
					%>
				</select>
			</td>
		</tr>
		<tr class=clsSilver>
			<td width=170>
				&nbsp;&nbsp;&nbsp;Redirecionar Carteira
			</td>
			<td>&nbsp;
				<input type=checkbox name=chkRedirCarteira value=1 <%if strRedirCarteira = "1" then Response.Write " checked" End if%>>
			</td>	
		</tr>
		
		<tr class=clsSilver>
			<td width=170>
				&nbsp;&nbsp;&nbsp;Possui Avaliador de Acesso
			</td>
			<td>&nbsp;
				<input type=checkbox name=chkGerenAvaliador value=1 <%if  strGerenAvaliador = "1" then Response.Write " checked" End if%>>
			</td>	
		</tr>
		
		</table>
		<table width="760">
			<tr>
				<td colspan=2 align="center">
					<br>
					<input type="submit" class="button" name="btnGravar" value="Gravar" accesskey="I" onmouseover="showtip(this,event,'Incluir (Alt+I)');">&nbsp;
					<input type="button" class="button" name="btnLimpar" value="Limpar" onclick="document.forms[0].hdnID.value='';LimparForm();setarFocus('cboCef');" accesskey="L" onmouseover="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
					<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('CefConfig_main.asp')" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');">
					<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">
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
<SCRIPT LANGUAGE=javascript>
<!--
setarFocus('cboCef')
//-->
</SCRIPT>

</html>
<%
Set objRSCtf = Nothing
DesconectarCla()
%>
