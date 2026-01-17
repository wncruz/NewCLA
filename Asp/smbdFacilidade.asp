<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: smbdFacilidade.asp
'	- Descrição			: Alocação de facilidade
%>
<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/smbdheader.asp"-->
<!--#include file="../inc/AlocacaoFac.asp"-->
<%
if Trim(Request.Form("hdnSolId")) <> "" then

	dblSolId = Request.Form("hdnSolId")
	dblPedId = Request.Form("hdnPedId")
	if dblPedId = "" then 
		Response.Write "<script language=javascript>alert('Pedido indisponível.');window.location.replace('facilidade_main.asp');</script>"
		Response.End 
	End if	
		
	Set objRS = db.execute("[dbo].[CLA_sp_view_smbd_pedido] null,null,null,null,null,null," & dblPedId)
	if objRS.Eof and objRS.Bof then 'Não existe registros
		Response.Write "<script language=javascript>alert('Pedido indisponível.');window.location.replace('facilidade_main.asp');</script>"
		Response.End 
	End if
	strProId = Request.Form("cboProvedor") 
	if strProId = "" then  strProId	= Trim(objRS("Pro_id"))
	strNroSolic		= Trim(objRS("Sol_id"))
	strPropAcesso	= Trim(objRS("Acf_Proprietario"))
	strIdLogico		= Trim(objRS("Acl_IDAcessoLogico"))
	strDataSolic	= Formatar_Data(Trim(objRS("Sol_Data")))
	strDM			= UCASE(objRS("Ped_Prefixo") & "-" & right("00000" & objRS("Ped_Numero"),5) & "/" & objRS("Ped_Ano"))
	strDataPedido	= Formatar_Data(objRS("Ped_Data"))
	strSatus		= Trim(objRS("Sts_Desc"))
	strCliente		= Trim(objRS("Cli_Nome"))
	intTipoProcesso = objRS("Tprc_idLog")
	strObs			= Trim(objRS("Ped_Obs")) 
	strAcao			= AcaoPedido(objRS("Tprc_id"))
	intTipoAcao		= objRS("Tprc_id")
				
	Set objRSEndPto = db.execute("CLA_sp_view_Ponto null," & dblPedId)
	if not objRSEndPto.Eof and not objRSEndPto.bof then
		strEndereco		= objRSEndPto("Tpl_Sigla") & " " & objRSEndPto("End_NomeLogr") & ", " & objRSEndPto("End_NroLogr") & " " & objRSEndPto("Aec_Complemento") & " • " & objRSEndPto("End_Bairro") & " • " & objRSEndPto("End_Cep") & " • " & objRSEndPto("Cid_Desc") & " • " & objRSEndPto("Est_Sigla")
	End if	
	Set objRSEndPto = Nothing
		
	strCidSigla	= Trim(objRS("Cid_Sigla"))
	strUfSigla	= Trim(objRS("Est_Sigla"))
	strTplSigla = Trim(objRS("Tpl_Sigla")) 
	strNomeLogr	= Trim(objRS("End_NomeLogr")) 
	strNroEnd	= Trim(objRS("End_NroLogr"))
	strCep		= Trim(objRS("End_Cep"))

	strNroServico	= Trim(objRS("Acl_NContratoServico"))
	strDesigServico = Trim(objRS("Acl_DesignacaoServico"))
	strServico		= Trim(objRS("Ser_Desc"))
	strVelServico	= Trim(objRS("DescVelAcessoLog"))
	strVelAcessoFis	= Trim(objRS("DescVelAcessoFis"))

	strPrmId		= Trim(objRS("Prm_id"))
	strRegId		= Trim(objRS("Reg_id"))
	
	strLocalInstala = Request.Form("estacao") 
	if strLocalInstala = "" then  strLocalInstala	= Trim(objRS("Esc_IdEntrega"))
	strLocalConfig	= Trim(objRS("Esc_IdConfiguracao"))
	strRecurso		= Trim(objRS("Rec_IDEntrega"))

	strDistrib = Request.Form("distribuidor") 
	if strDistrib = "" then  strDistrib	= Trim(objRS("Dst_Id"))

	if Request.Form("cboProvedor") <> "" then 
		strRede = 1 
	Else	
		strRede			= objRS("Sis_ID")
	End if	
	
	strExec = " CLA_sp_view_acessofisico null," & dblPedId & ", null, null, null, null, null, null, 'T' "
	
	set objRSFis = db.execute(strExec)
	''set objRSFis = db.execute("CLA_sp_view_acessofisico null," & dblPedId & ", null, null, null, null, null, null, 'T' )"
	
	
	if Not objRSFis.Eof and not objRSFis.Bof then
		strTecnologia	= objRSFis("Tec_Sigla")
		strVelFis	=	objRSFis("Vel_Desc")
		intTipoVel	=	objRSFis("Acf_TipoVel")
	End if
	
	set objRSPla = db.execute("CLA_sp_sel_facilidade null,null,null, " & objRSFis("Acf_ID")  )
	''set objRSPla = db.execute("CLA_sp_sel_facilidade null,null,null, " & 18695 )

	if Not objRSPla.Eof and not objRSPla.Bof then
		strPla	= objRSPla("Pla_ID")
	else
		strPla	= "0"
	End if
	

	Set objRSFis = Nothing

Else
	Response.Write "<script language=javascript>window.location.replace('facilidade_main.asp')</script>"
	Response.End
End if

strUserGicL		= strUserName

'Deterministico
blnGravado = false

'*** Para atender a alocação a partir do controle de rede determinística
if Trim(strRede)		= "" or isNull(strRede)			then strRede	= Request.Form("cboRede") 
if Trim(strProId)		= "" or isNull(strProId)		then strProId	= Request.Form("cboProvedor") 
if Trim(strLocalInstala)= "" or isNull(strLocalInstala) then strLocalInstala = Request.Form("cboLocalInstala") 
intCount = 0

strCamposDet = "<table cellspacing=1 cellpadding=0 width=760 border=0><th  width=30>&nbsp;</th>"
strCamposDet = strCamposDet & "<th width=100>&nbsp;Provedor</th>"
strCamposDet = strCamposDet & "<th width=100 nowrap>&nbsp;Nº Acesso Pta Ebt</th>"
strCamposDet = strCamposDet & "<th width=80>&nbsp;Fila</th>"
strCamposDet = strCamposDet & "<th width=80>&nbsp;Bastidor</th>"
strCamposDet = strCamposDet & "<th width=80>&nbsp;Régua</th>"
strCamposDet = strCamposDet & "<th width=80>&nbsp;Posição</th>"
strCamposDet = strCamposDet & "<th width=80>&nbsp;Timeslot</th>"
strCamposDet = strCamposDet & "<th width=130>&nbsp;Nº Acesso Pta CLI</th></table>"

strDet = strDet & "<table rules=groups border=0 cellspacing=1 cellpadding=1 bordercolorlight=#003388 bordercolordark=#ffffff width=760>"
strDet = strDet & "<tr>"
strDet = strDet & "<th colspan=7>&nbsp;•&nbsp;Informações da Facilidade</th>"
strDet = strDet & "</tr>"
strDet = strDet & "<tr class=clsSilver2>"
strDet = strDet & "<td><font class=clsObrig>::</font>&nbsp;Cod.Prov</td>"
strDet = strDet & "<td>&nbsp;Nº Acesso Pta Ebt</td>"
strDet = strDet & "<td><font class=clsObrig>::</font>&nbsp;Fila</td>"
strDet = strDet & "<td><font class=clsObrig>::</font>&nbsp;Bastidor</td>"
strDet = strDet & "<td><font class=clsObrig>::</font>&nbsp;Régua</td>"
strDet = strDet & "<td><font class=clsObrig>::</font>&nbsp;Posição</td>"
strDet = strDet & "<td><font class=clsObrig>::</font>&nbsp;Timeslot</td>"
strDet = strDet & "	</tr>"

strDet = strDet & "<tr class=clsSilver>"
strDet = strDet & "<td><span id=spnCodProv><select name=cboCodProv>"
strDet = strDet & "	<option value=""""></option>"
if Trim(Request.Form("cboProvedor")) <> "" then 
	Set objRS = db.execute("CLA_sp_sel_Provedor null," & Request.Form("cboProvedor"))
	While Not objRS.Eof
		strDet = strDet & "<Option value=" & Trim(objRS("Pro_ID")) & ">" & TratarAspasJS(objRS("Pro_Cod")) & "</Option>"
		objRS.MoveNext
	Wend
End if	
strDet = strDet & "</select></span></td>"
strDet = strDet & "	<td ><input type=text class=text name=txtNroAcessoEbt maxlength=25 size=15 onKeyPress=""if (ValidaProvedor(this)){parent.IFrmProcesso1.TamanhoMax(this)}""></td>"
strDet = strDet & "	<td ><input type=text class=text name=txtFila maxlength=25 size=10 value="""&Request.Form("fila")&""" >(A5)</td>"
strDet = strDet & "	<td ><input type=text class=text name=txtBastidor maxlength=3 size=4 value="""&Request.Form("bastidor")&""" onKeyUp=""ValidarTipo(this,0)"">(N3)</td>"
strDet = strDet & "	<td ><input type=text class=text name=txtRegua maxlength=2 size=3 value="""&Request.Form("regua")&""" >(A2)</td>"
strDet = strDet & "	<td ><input type=text class=text name=txtPosicao maxlength=3 size=5 value="""&Request.Form("posicao")&""" >(A3)</td>"
strDet = strDet & "	<td ><input type=text class=text name=txtTimeslot maxlength=9 size=10 onKeyUp=""ValidarNTipo(this,0,4,4,1,0,4)"" TIPO=N  value="""&Request.Form("timeslot")&""">(N4-N4)</td>"
strDet = strDet & "</tr>"

strDet = strDet & "<tr class=clsSilver2>"
strDet = strDet & " <td>Nº Acesso Pta Cli</td>"
strDet = strDet & "	<td nowrap>CCTO Provedor</td>"
strDet = strDet & "	<td nowrap>IU / URB</td>"
strDet = strDet & "	<td nowrap>CNL Ponta A</td>"
strDet = strDet & "	<td nowrap colspan=3>CNL Ponta B</td>"

''@davif 
if (strExigePortadora) = "S" Then
	strDet = strDet & "	<td nowrap colspan=2>Portadora</td>"
else
	strDet = strDet & "	<td nowrap colspan=2></td>"
end if
''/@davif 

strDet = strDet & "</tr>"


strDet = strDet & "<tr class=clsSilver>"
strDet = strDet & "	<td ><input type=text class=text name=txtNroAcessoCli size=15 maxlength=25 >(A25)</td>"
strDet = strDet & "<td><input type=text size=8 maxlength=8 class=text name=txtCCTOPro >(A8)</td>"
strDet = strDet & "<td >"
strDet = strDet & "	<input type=radio name=rdoUrbano value=I Index=0  onClick=""AdicionarCNL(this)"">IU&nbsp;"
strDet = strDet & "	<input type=radio name=rdoUrbano value=U Index=1 onClick=""AdicionarCNL(this)"">URB"
strDet = strDet & "</td>"
strDet = strDet & "<td><input type=text size=4 maxlength=4 class=text name=txtCNLPtaA >(A4)</td>"
strDet = strDet & "<td colspan=4><input type=text size=4 maxlength=4 class=text name=txtCNLPtaB >(A4)</td>"

''@davif 
if strExigePortadora = "S" then
	strDet = strDet & "<td valign=top colspan=2> "
	strDet = strDet & "		<select name=cboPortadora>"
	strDet = strDet & "			<Option value=""""></Option>"
	strDet = strDet & "			<Option value=Constante>Constante</Option>"
	strDet = strDet & "			<Option value=PseudoControlada>PseudoControlada</Option>"
	strDet = strDet & "		</select> </td>"
Else
	strDet = strDet & "<td valign=top colspan=2> </td>"
End if
''/@davif 

strDet = strDet & "</tr>"

strDet = strDet & "<tr class=clsSilver2>"
strDet = strDet & "	<td><font class=clsObrig>:: </font>Prop Modem</td>"
strDet = strDet & "	<td><font class=clsObrig>:: </font>Qtde Modem</td>"
if strPropAcesso <> "EBT" then
	strDet = strDet & "	<td><div id=""divFaturaLabel""><font class=clsObrig>:: </font>Fatura</div></td>"
	strDet = strDet & "	<td colspan=4>Obs</td>"
Else	
	strDet = strDet & "	<td colspan=5>Obs</td>"
End if	
strDet = strDet & "</tr>"

strDet = strDet & "<tr class=clsSilver>"
strDet = strDet & "	<td valign=top>"
strDet = strDet & "		<select name=cboPropModem>"
strDet = strDet & "			<Option value=""""></Option>"
strDet = strDet & "			<Option value=CLI >CLI</Option>"
if strPropAcesso = "EBT" then
	strDet = strDet & "			<Option value=EBT selected>EBT</Option>"
Else	
	strDet = strDet & "			<Option value=EBT >EBT</Option>"
End if	
strDet = strDet & "			<Option value=TER >TER</Option>"
strDet = strDet & "		</select>"
strDet = strDet & "	</td>"
strDet = strDet & "	<td valign=top><input type=text size=4 maxlength=2 class=text name=txtQtdeModem onKeyup=""ValidarTipo(this,0)"">(N2)</td>"

if strPropAcesso <> "EBT" then
	strDet = strDet & "	<td valign=top><div id=""divFatura"">"
	strDet = strDet & "		<input type=radio name=rdoFatura Index=0 value=S >SIM&nbsp;"
	strDet = strDet & "		<input type=radio name=rdoFatura Index=1 value=N >Não</div>"
	strDet = strDet & "	</td>"
	strDet = strDet & "	<td valign=top colspan=4>"
	strDet = strDet & "		<textarea name=txtObsFac cols=40 rows=2 onkeydown=""MaxLength(this,300);"" ></textarea>A(300)"
	strDet = strDet & "	</td>"
	strDet = strDet & "</tr>"
Else	
	strDet = strDet & "	<td valign=top colspan=5>"
	strDet = strDet & "		<textarea name=txtObsFac cols=40 rows=2 onkeydown=""MaxLength(this,300);"" ></textarea>A(300)"
	strDet = strDet & "	</td>"
	strDet = strDet & "</tr>"
End if
strDet = strDet & "<tr class=clsSilver2>"
strDet = strDet & "<td colspan=7 align=right>"
strDet = strDet & "<input type=button class=button name=btnAdd value=""Adicionar / Alterar"" onClick=""AddFac()"" accesskey=""A"" onmouseover=""showtip(this,event,\'Adicionar Facilidade(Alt+A)\');"">&nbsp;"
strDet = strDet & "<input type=button class=button name=btnRem value=""Remover"" onClick=""RemoverFacLista()"" accesskey=""R"" onmouseover=""showtip(this,event,\'Remover Facilidade(Alt+R)\');"">&nbsp;"
strDet = strDet & "<input type=button class=button name=btnLimpar value=""Limpar"" onClick=""LimparCamposFac()"" accesskey=""Q"" onmouseover=""showtip(this,event,\'Limpar Facilidade(Alt+Q)\');"">&nbsp;"
strDet = strDet & "</td>"
strDet = strDet & "</tr>"
strDet = strDet & "</table>"
	
'NÃO DETERMINISTICO
strCamposNdet = "<table cellspacing=1 cellpadding=0 width=760 border=0><th  width=30>&nbsp;</th>"
strCamposNdet = strCamposNdet & "<th  width=100>&nbsp;Provedor</th>"
strCamposNdet = strCamposNdet & "<th  width=100 nowrap>&nbsp;Nº Acesso Pta Ebt</th>"
strCamposNdet = strCamposNdet & "<th  width=100>&nbsp;Tronco</th>"
strCamposNdet = strCamposNdet & "<th  width=100>&nbsp;Par</th>"
strCamposNdet = strCamposNdet & "<th  width=310>&nbsp;Nº Acesso Pta CLI</th></table>"

strNDet = ""
strNDet = strNDet & "<table border=0 cellspacing=1 cellpadding=1 width=760>"
strNDet = strNDet & "	<tr >"
strNDet = strNDet & "		<th colspan=5>&nbsp;•&nbsp;Informações da Facilidade</th>"
strNDet = strNDet & "	</tr>"
strNDet = strNDet & "	<tr class=clsSilver2>"
strNDet = strNDet & "		<td nowrap><font class=clsObrig>::</font>&nbsp;Cód.Prov</td>"
strNDet = strNDet & "		<td nowrap>&nbsp;Nº Acesso Pta Ebt</td>"
strNDet = strNDet & "		<td nowrap><font class=clsObrig>::</font>&nbsp;Tronco</td>"
strNDet = strNDet & "		<td nowrap><font class=clsObrig>::</font>&nbsp;Par</td>"
strNDet = strNDet & "		<td nowrap>Nº Acesso Pta CLI</td>"
strNDet = strNDet & "	</tr>"

strNDet = strNDet & "<tr class=clsSilver>"
strNDet = strNDet & "<td><span id=spnCodProv><select name=cboCodProv>"
strNDet = strNDet & "<option value=""""></option>"
strNDet = strNDet & "</select></span></td>"
strNDet = strNDet & "<td><input type=text class=text name=txtNroAcessoEbt maxlength=25 size=15 onKeyPress=""if (ValidaProvedor(this)){parent.IFrmProcesso1.TamanhoMax(this)}"" ></td>"
strNDet = strNDet & "<td><input type=text class=text name=txtTronco maxlength=20 size=10></td>"
strNDet = strNDet & "<td><input type=text class=text name=txtPar maxlength=9 size=10 onKeyUp=""ValidarNTipo(this,0,4,4,1,0,4)"" TIPO=N  ></td>"
strNDet = strNDet & "<td><input type=text class=text name=txtNroAcessoCli maxlength=25 size=15></td>"
strNDet = strNDet & "</tr>"

strNDet = strNDet & "<tr class=clsSilver2>"
strNDet = strNDet & "	<td nowrap>CCTO Provedor</td>"
strNDet = strNDet & "	<td nowrap>IU / URB</td>"
strNDet = strNDet & "	<td nowrap>CNL Ponta A</td>"
strNDet = strNDet & "	<td nowrap colspan=2>CNL Ponta B</td>"

''@davi 
if strExigePortadora = "S" then
	strNDet = strNDet & "	<td nowrap colspan=1>Portadora </td>"
Else
	strNDet = strNDet & "	<td nowrap colspan=1> </td>"
End if


strNDet = strNDet & "</tr>"

strNDet = strNDet & "<tr class=clsSilver>"
strNDet = strNDet & "<td><input type=text size=8 maxlength=8 class=text name=txtCCTOPro></td>"
strNDet = strNDet & "<td >"
strNDet = strNDet & "	<input type=radio name=rdoUrbano value=I Index=0  onClick=""AdicionarCNL(this)"">IU&nbsp;"
strNDet = strNDet & "	<input type=radio name=rdoUrbano value=U Index=1 onClick=""AdicionarCNL(this)"">URB"
strNDet = strNDet & "</td>"
strNDet = strNDet & "<td><input type=text size=4 maxlength=4 class=text name=txtCNLPtaA></td>"
strNDet = strNDet & "<td colspan=2><input type=text size=4 maxlength=4 class=text name=txtCNLPtaB></td>"

''@davif 
if strExigePortadora = "S" then
	strNDet = strNDet & "<td nowrap colspan=2> "
	strNDet = strNDet & "		<select name=cboPortadora>"
	strNDet = strNDet & "			<Option value=""""></Option>"
	strNDet = strNDet & "			<Option value=Constante >Constante</Option>"
	strNDet = strNDet & "			<Option value=PseudoControlada >PseudoControlada</Option>"
	strNDet = strNDet & "		</select> </td>"
Else
	strNDet = strNDet & "<td nowrap colspan=2>  </td>"
End if
''/@davif 

strNDet = strNDet & "</tr>"

strNDet = strNDet & "<tr class=clsSilver2>"
strNDet = strNDet & "	<td><font class=clsObrig>:: </font>Prop Modem</td>"
strNDet = strNDet & "	<td><font class=clsObrig>:: </font>Qtde Modem</td>"
if strPropAcesso <> "EBT" then
	strNDet = strNDet & "	<td><div id=""divFaturaLabel""><font class=clsObrig>:: </font>Fatura</div></td>"
	strNDet = strNDet & "	<td colspan=2>Obs</td>"
Else	
	strNDet = strNDet & "	<td colspan=3>Obs</td>"
End if	
strNDet = strNDet & "</tr>"
strNDet = strNDet & "<tr class=clsSilver>"
strNDet = strNDet & "	<td valign=top>"
strNDet = strNDet & "		<select name=cboPropModem>"
strNDet = strNDet & "			<Option value=""""></Option>"
strNDet = strNDet & "			<Option value=CLI >CLI</Option>"
if strPropAcesso = "EBT" then
	strNDet = strNDet & "			<Option value=EBT selected>EBT</Option>"
Else	
	strNDet = strNDet & "			<Option value=EBT >EBT</Option>"
End if	
strNDet = strNDet & "			<Option value=TER >TER</Option>"
strNDet = strNDet & "		</select>"
strNDet = strNDet & "	</td>"
strNDet = strNDet & "	<td valign=top><input type=text size=4 maxlength=4 class=text name=txtQtdeModem onKeyup=""ValidarTipo(this,0)""></td>"

if strPropAcesso <> "EBT" then
	strNDet = strNDet & "	<td valign=top><div id=""divFatura"">"
	strNDet = strNDet & "		<input type=radio name=rdoFatura Index=0 value=S >SIM&nbsp;"
	strNDet = strNDet & "		<input type=radio name=rdoFatura Index=1 value=N >Não</div>"
	strNDet = strNDet & "	</td>"
	strNDet = strNDet & "	<td valign=top colspan=4>"
	strNDet = strNDet & "		<textarea name=txtObsFac cols=40 rows=2 onkeydown=""MaxLength(this,300);"" ></textarea>A(300)"
	strNDet = strNDet & "	</td>"
	strNDet = strNDet & "</tr>"
Else	
	strNDet = strNDet & "	<td valign=top colspan=5>"
	strNDet = strNDet & "		<textarea name=txtObsFac cols=40 rows=2 onkeydown=""MaxLength(this,300);"" ></textarea>A(300)"
	strNDet = strNDet & "	</td>"
	strNDet = strNDet & "</tr>"
End if

strNDet = strNDet & "<tr class=clsSilver2>"
strNDet = strNDet & "<td colspan=5 align=right>"
strNDet = strNDet & "<input type=button class=button name=btnAdd value=""Adicionar / Alterar""  onClick=""AddFac()"" accesskey=""A"" onmouseover=""showtip(this,event,\'Adicionar Facilidade(Alt+A)\');"">&nbsp;"
strNDet = strNDet & "<input type=button class=button name=btnRem value=""Remover"" onClick=""RemoverFacLista()"" accesskey=""R"" onmouseover=""showtip(this,event,\'Remover Facilidade(Alt+R)\');"">&nbsp;"
strNDet = strNDet & "<input type=button class=button name=btnLimpar value=""Limpar"" onClick=""LimparCamposFac()"" accesskey=""Q"" onmouseover=""showtip(this,event,\'Limpar Facilidade(Alt+Q)\');"">&nbsp;"
strNDet = strNDet & "</td>"
strNDet = strNDet & "</tr>"
strNDet = strNDet & "</table>" 
	
'ADE
strCamposAde = "<table cellspacing=1 cellpadding=0 width=760 border=0><th  width=30>&nbsp;</th>"
strCamposAde = strCamposAde & "<th width=100>&nbsp;Nº Acesso</th>"
strCamposAde = strCamposAde & "<th width=100>&nbsp;Cabo</th>"
strCamposAde = strCamposAde & "<th width=100>&nbsp;Par</th>"
strCamposAde = strCamposAde & "<th nowrap width=100>&nbsp;PADE</th>"
strCamposAde = strCamposAde & "<th width=100>&nbsp;Derivação</th>"
strCamposAde = strCamposAde & "<th nowrap width=100>&nbsp;T. Cabo</th></table>"

strAde = strAde & "<table rules=groups border=0 cellspacing=1 cellpadding=1 bordercolorlight=#003388 bordercolordark=#ffffff width=760>"
strAde = strAde & "	<tr>"
strAde = strAde & "		<th colspan=6>&nbsp;•&nbsp;Informações da Facilidade</td>"
strAde = strAde & "	</tr>"
strAde = strAde & "	<tr class=clsSilver2>"
strAde = strAde & "		<td width=150>&nbsp;Nº Acesso</td>"
strAde = strAde & "		<td width=100><font class=clsObrig>::</font>&nbsp;Cabo</td>"
strAde = strAde & "		<td width=120><font class=clsObrig>::</font>&nbsp;Par</td>"
strAde = strAde & "		<td nowrap><font class=clsObrig></font>&nbsp;PADE</td>"
strAde = strAde & "		<td width=100><font class=clsObrig></font>&nbsp;Derivação</td>"
strAde = strAde & "		<td nowrap width=100><font class=clsObrig></font>&nbsp;T. Cabo</td>"
strAde = strAde & "	</tr>"
strAde = strAde & "<tr class=clsSilver>"
strAde = strAde & "<td width=150><input type=text class=text name=txtNroAcessoEbt	maxlength=25 size=15 onKeyPress=""if (ValidaProvedor(this)){parent.IFrmProcesso1.TamanhoMax(this)}""></td>"
strAde = strAde & "<td width=100><input type=text class=text name=txtCabo		maxlength=2	 size=5	 onKeyUp=""ValidarTipo(this,0)"" ></td>"
strAde = strAde & "<td width=120><input type=text class=text name=txtPar		maxlength=9  size=10 onKeyUp=""ValidarNTipo(this,0,4,4,1,0,4)"" TIPO=N ></td>"
strAde = strAde & "<td><input type=text class=text name=txtCaixaEmenda  maxlength=3 size=5 ></td>"
strAde = strAde & "<td width=100><input type=text class=text name=txtLateral	maxlength=4  size=5></td>"
strAde = strAde & "<td width=100><select name=cboTipoCabo >"
strAde = strAde & "			<option value=""""></option>"
strAde = strAde & "			<option value=""H"" >H</option>"
strAde = strAde & "			<option value=""X"" >X</option>"
'strAde = strAde & "			<option value=XDSL >XDSL</option>"
strAde = strAde & "			<option value=""C"" >C</option>"
strAde = strAde & "		</select></td>"
strAde = strAde & "</tr>"

strAde = strAde & "<tr class=clsSilver2>"
strAde = strAde & "	<td><font class=clsObrig>:: </font>Prop Modem</td>"
strAde = strAde & "	<td><font class=clsObrig>:: </font>Qtde Modem</td>"
if strPropAcesso <> "EBT" then
	strAde = strAde & "	<td><div id=""divFaturaLabel""><font class=clsObrig>:: </font>Fatura</div></td>"
	strAde = strAde & "	<td colspan=3>Obs</td>"
Else
	strAde = strAde & "	<td colspan=4>Obs</td>"
End if	
strAde = strAde & "</tr>"
strAde = strAde & "<tr class=clsSilver>"
strAde = strAde & "	<td valign=top>"
strAde = strAde & "		<select name=cboPropModem>"
strAde = strAde & "			<Option value=""""></Option>"
strAde = strAde & "			<Option value=CLI >CLI</Option>"
if strPropAcesso = "EBT" then
	strAde = strAde & "			<Option value=EBT selected>EBT</Option>"
Else	
	strAde = strAde & "			<Option value=EBT >EBT</Option>"
End if	
strAde = strAde & "			<Option value=TER >TER</Option>"
strAde = strAde & "		</select>"
strAde = strAde & "	</td>"
strAde = strAde & "	<td valign=top><input type=text size=4 maxlength=4 class=text name=txtQtdeModem onKeyup=""ValidarTipo(this,0)""></td>"

if strPropAcesso <> "EBT" then
	strAde = strAde & "	<td valign=top><div id=""divFatura"">"
	strAde = strAde & "		<input type=radio name=rdoFatura Index=0 value=S >SIM&nbsp;"
	strAde = strAde & "		<input type=radio name=rdoFatura Index=1 value=N >Não</div>"
	strAde = strAde & "	</td>"
	strAde = strAde & "	<td valign=top colspan=4>"
	strAde = strAde & "		<textarea name=txtObsFac cols=40 rows=2 onkeydown=""MaxLength(this,300);"" ></textarea>A(300)"
	strAde = strAde & "	</td>"
	strAde = strAde & "</tr>"
Else	
	strAde = strAde & "	<td valign=top colspan=5>"
	strAde = strAde & "		<textarea name=txtObsFac cols=40 rows=2 onkeydown=""MaxLength(this,300);"" ></textarea>A(300)"
	strAde = strAde & "	</td>"
	strAde = strAde & "</tr>"
End if

strAde = strAde & "<tr class=clsSilver2>"
strAde = strAde & "<td colspan=6 align=right>"
strAde = strAde & "<input type=button class=button name=btnAdd value=""Adicionar / Alterar""  onClick=""AddFac()"" accesskey=""A"" onmouseover=""showtip(this,event,\'Adicionar Facilidade(Alt+A)\');"">&nbsp;"
strAde = strAde & "<input type=button class=button name=btnRem value=""Remover"" onClick=""RemoverFacLista()"" accesskey=""R"" onmouseover=""showtip(this,event,\'Remover Facilidade(Alt+R)\');"">&nbsp;"
strAde = strAde & "<input type=button class=button name=btnLimpar value=""Limpar"" onClick=""LimparCamposFac()"" accesskey=""Q"" onmouseover=""showtip(this,event,\'Limpar Facilidade(Alt+Q)\');"">&nbsp;"
strAde = strAde & "</td>"
strAde = strAde & "</tr>"
strAde = strAde & "</table>"

Dim strStatus
strStatus = ""

Dim strAcfObs
Set objXmlDados = Server.CreateObject("Microsoft.XMLDOM")
objXmlDados.loadXml("<xDados/>")

Set objRS = db.Execute("CLA_SP_Sel_Facilidade " & dblPedId)
if intTipoProcesso = 4 then 'Cancelamento
	Set objXmlDados = MontarXmlFacilidade(objXmlDados,objRS,strStatus,intTipoProcesso,"R")
Else
	Set objXmlDados = MontarXmlFacilidade(objXmlDados,objRS,strStatus,intTipoProcesso,"A")
End if	

Set objNode = objXmlDados.SelectNodes("//Facilidade")

intFac = objNode.length
if intFac > 0 then blnGravado = true
strXmlFac = FormatarXml(objXmlDados)

Function MontaXmlFacOrigem(objRSFacOrigem)

	Dim objNodeFac
	Dim objXMLFacOrigem

	Set objXMLFacOrigem = Server.CreateObject("Microsoft.XMLDOM")
	objXMLFacOrigem.loadXml("<xDados/>")

	if not objRSFacOrigem.Eof and not objRSFacOrigem.bof then

		While not objRSFacOrigem.Eof

			Set objNodeFac = objXMLFacOrigem.createNode("element", "Facilidade", "")
			objXMLFacOrigem.documentElement.appendChild (objNodeFac)

			For intIndex=0 to objRSFacOrigem.fields.count-1
				Call AddElemento(objXMLFacOrigem,objNodeFac,objRSFacOrigem.fields(intIndex).name,objRSFacOrigem.fields(intIndex).value)
			Next
			Call AddElemento(objXMLFacOrigem,objNodeFac,"intFacSel",0)
			objRSFacOrigem.MoveNext
		Wend
		objRSFacOrigem.MoveFirst
	End if
	Set MontaXmlFacOrigem = objXMLFacOrigem
End function
%>
<script language='javascript' src="../javascript/xmlFacObjects.js"></script>
<SCRIPT LANGUAGE=javascript>
<!--
var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")
var objXmlReturn = new ActiveXObject("Microsoft.XMLDOM")

var intIndice = 0
var objAryObjs = new Array()
var objAryFac = new Array()
var objAryFacRet

<%if strPropAcesso <> "EBT" then%>
//alert("Proprietario Não EBT")
	<% If strExigePortadora = "S" then %>
		objAryObjs[1] = new Array("txtNroAcessoEbt","txtFila","txtBastidor","txtRegua","txtPosicao","txtTimeslot","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB", "cboPortadora", "rdoUrbano[0]","rdoUrbano[1]","cboRede","Ped_Id","cboPropModem","txtQtdeModem","rdoFatura[0]","rdoFatura[1]","txtObsFac")
		objAryObjs[2] = new Array("txtNroAcessoEbt","txtTronco","txtPar","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB", "cboPortadora", "rdoUrbano[0]","rdoUrbano[1]","cboRede","Ped_Id","cboPropModem","txtQtdeModem","rdoFatura[0]","rdoFatura[1]","txtObsFac")
	<% Else %>
		objAryObjs[1] = new Array("txtNroAcessoEbt","txtFila","txtBastidor","txtRegua","txtPosicao","txtTimeslot","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB",  "rdoUrbano[0]","rdoUrbano[1]","cboRede","Ped_Id","cboPropModem","txtQtdeModem","rdoFatura[0]","rdoFatura[1]","txtObsFac")
		objAryObjs[2] = new Array("txtNroAcessoEbt","txtTronco","txtPar","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB", "rdoUrbano[0]","rdoUrbano[1]","cboRede","Ped_Id","cboPropModem","txtQtdeModem","rdoFatura[0]","rdoFatura[1]","txtObsFac")
	<%End if %>
	objAryObjs[3] = new Array("txtNroAcessoEbt","txtCabo","txtPar","txtLateral","cboTipoCabo","txtCaixaEmenda","cboRede","Ped_Id","cboPropModem","txtQtdeModem","rdoFatura[0]","rdoFatura[1]","txtObsFac")
<%Else%>
	objAryObjs[1] = new Array("txtNroAcessoEbt","txtFila","txtBastidor","txtRegua","txtPosicao","txtTimeslot","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB","rdoUrbano[0]","rdoUrbano[1]","cboRede","Ped_Id","cboPropModem","txtQtdeModem","txtObsFac")
	objAryObjs[2] = new Array("txtNroAcessoEbt","txtTronco","txtPar","txtNroAcessoCli","cboCodProv","txtCCTOPro","txtCNLPtaA","txtCNLPtaB","rdoUrbano[0]","rdoUrbano[1]","cboRede","Ped_Id","cboPropModem","txtQtdeModem","txtObsFac")
	objAryObjs[3] = new Array("txtNroAcessoEbt","txtCabo","txtPar","txtLateral","cboTipoCabo","txtCaixaEmenda","cboRede","Ped_Id","cboPropModem","txtQtdeModem","txtObsFac")
<%End if%>

objAryFac[0] = new Array("","")

function CarregarLista()
{
	objXmlGeral.onreadystatechange = CheckStateXml;
	objXmlGeral.resolveExternals = false;
	if (parseInt(<%=intFac%>) != 0){
		objXmlGeral.loadXML("<%=strXmlFac%>") 
	}else{
		var objXmlRoot = objXmlGeral.createNode("element","xDados","")
		objXmlGeral.appendChild (objXmlRoot)
	}	
}
//Verifica se o Xml já esta carregado
function CheckStateXml()
{
  var state = objXmlGeral.readyState;
  
  if (state == 4)
  {
    var err = objXmlGeral.parseError;
    if (err.errorCode != 0)
    {
      alert(err.reason)
    } 
    else 
    {
		AtualizarListaFac()
	}
  }
}


function CarregarDoc()
{
	document.onreadystatechange = CheckStateDoc;
	document.resolveExternals = false;
}

function CheckStateDoc()
{
  var state = document.readyState;
  
  if (state == "complete")
  {
	ResgatarInfoRede()
	PopularCamposFac()
  }
}


function VoltarOrigem()
{
	var strPagina = new String("<%=Request.Form("hdnPaginaOrig")%>")
	with (document.forms[0])
	{
		target = self.name
		if (strPagina == "" || strPagina.toUpperCase().indexOf("FACILIDADE.ASP") != -1)
		{
			action = "facilidade_main.asp"

		}else
		{
			action = strPagina
		}	
		submit()
	}
}


function RetornaCboPlataforma(sisID, PlaID)
{
	if (sisID != 1) {
		spnPlataforma.innerHTML = ""
		return 
	}
	var xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
	var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
	var strXML
	
	strXML = "<root>"
	strXML = strXML + "<plaid>" +  PlaID + "</plaid>"
	strXML = strXML + "<funcao></funcao>"
	strXML = strXML + "</root>" 
	
	xmlDoc.loadXML(strXML);
	xmlhttp.Open("POST","RetornaPlataforma.asp" , false);
	xmlhttp.Send(xmlDoc.xml);
	
	strXML = xmlhttp.responseText;
	spnPlataforma.innerHTML = strXML
	
}


CarregarDoc()
//-->
</SCRIPT>
<table border=0 cellspacing="1" cellpadding="1" width="760">
<Form name="Form2" method="Post">
<input type="hidden" name="hdnProvedor" value="<%=strProId%>">
<input type="hidden" name="hdnEstacaoAtual" >
<input type="hidden" name="hdnNomeLocal" >
<input type="hidden" name="Ped_Id" value="<%=dblPedId%>">
<input type="hidden" name="hdnAcao">
<input type=hidden	 name="hdnPlataforma" value="<%=strPla%>" >
<input type="hidden" name="hdnRecId">
<input type="hidden" name="hdnSubAcao">
<input type="hidden" name="hdnRede" value="<%=strRede%>">
<input type="hidden" name="hdnPedId"		value="<%=Request.Form("hdnPedId")%>">
<input type="hidden" name="hdnSolId"		value="<%=Request.Form("hdnSolId")%>">
<input type="hidden" name="hdnPaginaOrig"	value="<%=Request.ServerVariables("SCRIPT_NAME")%>">
<input type="hidden" name="txtRazaoSocial"	value="<%=strCliente%>">
<input type="hidden" name="txtEndCid"		value="<%=strCidSigla%>">
<input type="hidden" name="cboUFEnd"		value="<%=strUfSigla%>">
<input type="hidden" name="cboLogrEnd"		value="<%=strTplSigla%>">
<input type="hidden" name="txtEnd"			value="<%=strNomeLogr%>">
<input type="hidden" name="txtNroEnd"		value="<%=strNroEnd%>">
<input type="hidden" name="txtCepEnd"		value="<%=strCep%>">
<input type="hidden" name="hdnTipoProcesso"	value="<%=intTipoProcesso%>">
<input type="hidden" name="hdnTipoAcao"	value="<%=intTipoAcao%>"> <!-- Adicionado PRSS 13/09/2007-->
<input type="hidden" name="hdnDstId">
<input type="hidden" name="hdnFacDetid"> <!-- Adicionado PRSSILV 26/05/2008-->

<input type=hidden name="hdnIdAcessoFisico" value="0">
<input type=hidden name="hdnIdAcessoFisico1" value="0">
<input type=hidden name="hdnCompartilhamento" value="0">
<input type=hidden name="hdnCompartilhamento1" value="0">
<input type=hidden name="hdnAlteracao" value="0">
<input type=hidden name="hdnPodeAlterar" value="S">
<input type=hidden name="hdnPropIdFisico" >
<input type=hidden name="hdnPropIdFisico1" >
<input type=hidden name="hdnXmlReturn" value='<%=FormatarStrXml(Request.Form("hdnXmlReturn"))%>'>

<input type=hidden name="hdnIdLog" value="<%=strIdLogico%>">
<input type=hidden name="hdnPropAcesso" value="<%=strPropAcesso%>">
<input type=hidden name="hdnTecId" value="<%=intTec%>">
<input type=hidden name="hdnTecSigla" value="<%=strTecSigla%>">
<input type=hidden name="hdnEnviarEmail">
<input type=hidden name="hdnControleNovoPed" value="NAO">

<input type=hidden name="hdnCtfcIdGLA" value="<%=dblCtfcIdGAT%>">
<input type=hidden name="hdnCtfcIdGLAE" value="<%=dblCtfcIdGAE%>">

<input type="hidden" name="hdnIntIndice">
<input type="hidden" name="hdnXml" >
<input type="hidden" name="hdnXmlFacLibera">

<tr><th colspan=6><p align=center>Alocação de Facilidade</p></th></tr>

<tr class=clsSilver>
	<td nowrap width=170>Solicitação de Acesso Nº</td>
	<td>&nbsp;<%=strNroSolic%></td>
	<td align=right >Id Lógico</td>
	<td>&nbsp;<%=strIdLogico%></td>
	<td align=right >Data</td>
	<td>&nbsp;<%=strDataSolic%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Pedido de acesso</td>
	<td>&nbsp;<%=strDM%></td>
	<td align=right>Data</td>
	<td colspan=3 nowrap>&nbsp;<%=strDataPedido%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Ação</td>
	<td colspan=5>&nbsp;<%=strAcao%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Status</td>
	<td colspan=5>&nbsp;<%=strSatus%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Cliente</td>
	<td colspan="5">&nbsp;<%=strCliente%></td>
</tr>

<tr class=clsSilver>
	<td width=170>Endereço</td>
	<td colspan="5">&nbsp;<%=strEndereco%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Nº Contrato Serviço</td>
	<td nowrap>&nbsp;<%=strNroServico%></td>
	<td align=right>Designação do Serviço</td>
	<td colspan=3>&nbsp;<%=strDesigServico%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Serviço</td>
	<td >&nbsp;<%=strServico%></td>
	<td align=right>Velocidade do Serviço</td>
	<td colspan=3>&nbsp;<%=strVelServico%></td>
</tr>
<tr class=clsSilver>
	<td width=170>Velocidade do Acesso Físico</td>
	<td colspan=5>&nbsp;<%=strVelFis%>&nbsp;<%=TipoVel(intTipoVel)%></td>
</tr>


<tr class="clsSilver">
	<td width=170>Observações</td>
	<td colspan="6"><textarea name="txtObsProvedor" cols="50" rows="2"><%=strObs%></textarea></td>
</tr>
</table>
<%
if intTipoProcesso = 3 and intTipoAcao = 3 then

	Vetor_Campos(1)="adInteger,2,adParamInput," & dblSolId
	Vetor_Campos(2)="adInteger,2,adParamInput," & dblPedId
	strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_logEvento",2,Vetor_Campos)
	Set objRS = db.Execute(strSqlRet)
'Response.Write "Sol" & dblSolId
'Response.Write "Ped" & dblPedId

	if not objRS.Eof then
		if objRS("Retorno") = "Verdadeiro" then%>
			<table cellpadding=0 cellspacing=1 width=760>
			<tr>
				<tr>
					<th colspan=4>&nbsp;•&nbsp;Informações de Origem</th>
				</tr>
				<tr>
					<th nowrap width=100 >&nbsp;Data</th>
					<th width=200>&nbsp;Campo</th>
					<th width=360>&nbsp;Valor</th>
					<th width=100>&nbsp;Usuário</th>
				</tr>
			<%
				blnCor = true
				While Not objRS.Eof
					if blnCor then
						strHtml = strHtml &  "<tr class=clsSilver >"
						blnCor = false
					Else
						strHtml = strHtml &  "<tr class=clsSilver2>"
						blnCor = true
					End if	
					strHtml = strHtml &  "<td nowrap >&nbsp;"& Formatar_Data(objRS("Log_Data")) &"</td>"		
					strHtml = strHtml &  "<td nowrap>&nbsp;" & objRS("Log_Campo") & "</td>"
					strHtml = strHtml &  "<td nowrap>&nbsp;" & objRS("Log_Valor") & "</td>"
					strHtml = strHtml &  "<td nowrap>&nbsp;"& objRS("Usu_UserName") & "</td>"
					strHtml = strHtml &  "</tr>"
					objRS.MoveNext
				Wend
				Response.Write strHtml
			%>
			</table>
		<%
		End if
	End if
End if%>
<table border=0 cellspacing="1" cellpadding="0" width="760">  
	<tr>
		<td>
			<iframe	id			= "IFrmMotivoPend"
				    name        = "IFrmMotivoPend" 
				    width       = "100%" 
				    height      = "120px"
				    src			= "../inc/MotivoPendencia.asp?dblSolId=<%=dblSolId%>&dblPedId=<%=dblPedId%>"
				    frameborder = "0"
				    scrolling   = "no" 
				    align       = "left">
			</iFrame>
		</td>
	</tr>
</table>
<table rules="groups"  border=0 cellspacing="1" cellpadding="0" bordercolorlight="#003388" bordercolordark="#ffffff" width="760">
	<tr>
		<th colspan=2>&nbsp;•&nbsp;Recurso</th>
	</tr>

	<tr class=clsSilver >
		<td width=170><font class="clsObrig">:: </font>Local de Entrega</td>
		<td >
			<select name="cboLocalInstala" onChange="ResgatarDistribuicao(this,1)"  style="width:100px">
				<option value=""></option>
				<%set objRS = db.execute("CLA_sp_sel_usuarioesc " & dblUsuId)
					While not objRS.Eof 
						strItemSel = ""
						if Trim(strLocalInstala) = Trim(objRS("Esc_ID")) then strItemSel = " Selected " End if
						Response.Write "<Option value=" & objRS("Esc_ID") & strItemSel & ">" & objRS("Cid_Sigla") & "  " & objRS("Esc_Sigla") & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
				%>
			</select>
		</td>
	</tr>

	<tr class=clsSilver>
		<td width=170 nowrap><font class="clsObrig">:: </font>Local de Configuração</td>
		<td  >
			<select name="cboLocalConfig" style="width:100px">
				<option value=""></option>
				<%'set objRS = db.execute("CLA_sp_sel_estacao  null")
					set objRS = db.execute("CLA_sp_sel_usuarioesc " & dblUsuId)
					While not objRS.Eof 
						strItemSel = ""
						if Trim(strLocalConfig) = Trim(objRS("Esc_ID")) then strItemSel = " Selected " End if
						Response.Write "<Option value=" & objRS("Esc_ID") & strItemSel & ">" & objRS("Cid_Sigla") & "  " & objRS("Esc_Sigla") & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
				%>
			</select>
		</td>
	</tr>

	<tr class=clsSilver>
		<td width=170><font class="clsObrig">:: </font>Distribuidor</td>
		<td >
			<span id="spnDistLocalInstala">
			<select name="cboDistLocalInstala" style="width:200px">
				<option value=""></option>
				<%	if Trim(strLocalInstala) <> "" then 
						set objRS = db.execute("CLA_sp_view_recursodistribuicao " & strLocalInstala )
						While not objRS.Eof 
							strItemSel = ""
							if Trim(strDistrib) = Trim(objRS("Dst_ID")) then strItemSel = " Selected " End if
							Response.Write "<Option value=" & objRS("Dst_ID") & strItemSel & ">" & objRS("Dst_Desc") & "</Option>"
							objRS.MoveNext
						Wend
						strItemSel = ""
					End if	
				%>
			</select>
			</span>
		</td>
	</tr>

	<tr class=clsSilver>
		<td width=170><font class="clsObrig">:: </font>Rede</td>
		<td >
			<select name="cboRede" onChange="RetornaCboPlataforma(cboRede.value,<% = strPla  %>)" onBlur="ControleRede(this);" <%if blnGravado then Response.Write " disabled " End if %> >
				<option value=""></option>
				<%
					Set objRS = db.execute("CLA_sp_sel_sistema 0")
					While not objRS.Eof 
						strItemSel = ""
						if Trim(strRede) = Trim(objRS("Sis_ID")) then strItemSel = " Selected " End if
						Response.Write "<Option value='" & objRS("Sis_ID")& "'" & strItemSel & ">" & Ucase(objRS("Sis_Desc")) & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""

				%>
			</select>
			<span ID = spnPlataforma>
			</span>
		</td>
	</tr>


	<tr class=clsSilver>
		<td width=170 ><font class="clsObrig">:: </font>Provedor</td>
		<td >
			<select name="cboProvedor" onBlur="ControleProvedor(this);" onChange="MsgTrocaProvedor();">
				<option value=""></option>
				<%	set objRS = db.execute("CLA_sp_sel_provedor 0,null,1")
					While not objRS.Eof 
						strItemSel = ""
						if Trim(strProId) = Trim(objRS("Pro_ID")) then strItemSel = " Selected " End if
						Response.Write "<Option value='" & Trim(objRS("Pro_ID")) & "'" & strItemSel & ">" & objRS("Pro_Nome") & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
				%>
			</select>
		</td>
	</tr>
	<tr class=clsSilver>
		<td width=170><font class="clsObrig">:: </font>Prazos de Contratação de Acesso</td>
		<td >
			<span id=spnRegimeCntr>
				<select name="cboRegimeCntr">
					<option value=""></option>
					<%
						if Trim(strProID) <> "" then
							set objRS = db.execute("CLA_sp_sel_regimecontrato 0," & strProID)
							While not objRS.Eof 
								strItemSel = ""
								if Trim(strRegId) = Trim(objRS("Reg_ID")) then strItemSel = " Selected " End if
								Response.Write "<Option value='" & Trim(objRS("Reg_ID")) & "'" & strItemSel & ">" & LimparStr(Trim(objRS("Pro_Nome"))) & " - " & LimparStr(Trim(objRS("Tct_Desc"))) & "</Option>"
								objRS.MoveNext
							Wend
							strItemSel = ""
						End if			
					%>
				</select>
			</span>
		</td>
	</tr>
	<tr class=clsSilver>
		<td width=170>&nbsp;&nbsp;&nbsp;Promoção</td>
		<td >
			<span id=spnPromocao>
			<select name="cboPromocao" style="width:200px">
				<option value=""></option>
				<%
					if Trim(strProId) <> "" then
						set objRS = db.execute("CLA_sp_sel_promocaoprovedor 0," & strProId)
						While not objRS.Eof 
							strItemSel = ""
							if Trim(strPrmId) = Trim(objRS("Prm_ID")) then strItemSel = " Selected " End if
							Response.Write "<Option value='" & Trim(objRS("Prm_ID")) & "'" & strItemSel & ">" & objRS("Prm_Desc") & "</Option>"
							objRS.MoveNext
						Wend
						strItemSel = ""
					End if	
				%>
			</select>
			</span>
		</td>
	</tr>
	<tr class=clsSilver>
		<td width=170>&nbsp;&nbsp;&nbsp;Status</td>
		<td>
			<input type=text size=25 readonly style='BACKGROUND-COLOR:#eeeeee' class=text name="txtStatus" tabIndex=-1 VALUE="<%=strStatus%>" onClick="AbrirInterligacoes()" style="cursor:hand" onmouseover="showtip(this,event,'Interligações Alocadas');">
		</td>	
	</tr>	
	<%
	if intTipoProcesso = 3 and intTipoAcao = 3 then

			Set objDic = Server.CreateObject("Scripting.Dictionary") 

			Vetor_Campos(1)="adInteger,2,adParamInput," & dblPedId
			strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_HistoricoFac",1,Vetor_Campos)

			Set objRS = db.Execute(strSqlRet)
			objRS.Close
			objRS.CursorLocation = adUseClient
			objRS.Open

			strXmlFacOrigem = FormatarXml(MontaXmlFacOrigem(objRS))

			if not objRS.Eof and not objRS.Bof then
				'objRS.MoveFirst

				strHtml = "<tr class=clsSilver>"
				strHtml = strHtml &  "<td colspan=2>"
				strHtml = "<table cellpadding=0 cellspacing=1 width=760>"

				dblSisIdFacOrig = objRS("Sis_ID")

				Select Case dblSisIdFacOrig

					Case 1
						strHtml = strHtml &  "<tr>"
						strHtml = strHtml &  "<th colspan=7>&nbsp;•&nbsp;Facilidades de Origem</th>"
						strHtml = strHtml &  "</tr>"

						strHtml = strHtml &  "<tr>"
						strHtml = strHtml &  "<th nowrap >&nbsp;Nº Acesso Pta EBT</th>"
						strHtml = strHtml &  "<th nowrap >&nbsp;Fila</th>"
						strHtml = strHtml &  "<th >&nbsp;Bastidor</th>"
						strHtml = strHtml &  "<th >&nbsp;Régua</th>"
						strHtml = strHtml &  "<th >&nbsp;Posição</th>"
						strHtml = strHtml &  "<th >&nbsp;Timeslot</th>"
						strHtml = strHtml &  "<th></th>"
						strHtml = strHtml &  "</tr>"
					Case 2
						strHtml = strHtml &  "<tr>"
						strHtml = strHtml &  "<th colspan=4>&nbsp;•&nbsp;Facilidades de Origem</th>"
						strHtml = strHtml &  "</tr>"

						strHtml = strHtml &  "<tr>"
						strHtml = strHtml &  "<th nowrap >&nbsp;Nº Acesso Pta EBT</th>"
						strHtml = strHtml &  "<th >&nbsp;Tronco</th>"
						strHtml = strHtml &  "<th >&nbsp;Par</th>"
						strHtml = strHtml &  "<th></th>"
						strHtml = strHtml &  "</tr>"
					Case 3
						strHtml = strHtml &  "<tr>"
						strHtml = strHtml &  "<th colspan=7>&nbsp;•&nbsp;Facilidades de Origem</th>"
						strHtml = strHtml &  "</tr>"

						strHtml = strHtml &  "<tr>"
						strHtml = strHtml &  "<th nowrap >&nbsp;Nº Acesso Pta EBT</th>"
						strHtml = strHtml &  "<th >&nbsp;Cabo</th>"
						strHtml = strHtml &  "<th >&nbsp;Par</th>"
						strHtml = strHtml &  "<th >&nbsp;PADE</th>"
						strHtml = strHtml &  "<th >&nbsp;Derivação</th>"
						strHtml = strHtml &  "<th >&nbsp;Tipo Cabo</th>"
						strHtml = strHtml &  "<th></th>"
						strHtml = strHtml &  "</tr>"

				End Select	

				blnCor = true
				While Not objRS.Eof
					Select Case dblSisIdFacOrig
						Case 1
							strRepresentacao = objRS("Fac_TimeSlot")
						Case Else	
							strRepresentacao = objRS("Fac_Par")

					End Select		

					if not isNull(objRS("Fac_Representacao")) then
						strRepresentacao = objRS("Fac_Representacao")
					End if

					Select Case dblSisIdFacOrig
						Case 1  
							if  Not objDic.Exists(strRepresentacao) then
								Call objDic.Add(strRepresentacao,objRS("Fac_Id")) 
								if blnCor then
									strHtml = strHtml &  "<tr class=clsSilver >"
									blnCor = false
								Else
									strHtml = strHtml &  "<tr class=clsSilver2>"
									blnCor = true
								End if	
								strHtml = strHtml &  "<td nowrap >&nbsp;"	& objRS("Acf_NroAcessoPtaEbt")	&"</td>"
								strHtml = strHtml &  "<td nowrap >&nbsp;"	& objRS("Fac_Fila")	&"</td>"		
								strHtml = strHtml &  "<td nowrap >&nbsp;"	& objRS("Fac_Bastidor")	&"</td>"		
								strHtml = strHtml &  "<td nowrap>&nbsp;"	& objRS("Fac_Regua")		& "</td>"
								strHtml = strHtml &  "<td nowrap>&nbsp;"	& objRS("Fac_Posicao")	& "</td>"
								strHtml = strHtml &  "<td nowrap>&nbsp;"	& strRepresentacao	& "</td>"
								strHtml = strHtml &  "<td >&nbsp;<input type=button class=button name=btnLiberarFac value=Alocar onClick='LiberarFacilidade(" & objRS("Fac_Id") & ")'></td>"
								strHtml = strHtml &  "</tr>"
							End if	
						Case 2
							if  Not objDic.Exists(strRepresentacao) then
								Call objDic.Add(strRepresentacao,objRS("Fac_Id")) 
								if blnCor then
									strHtml = strHtml &  "<tr class=clsSilver >"
									blnCor = false
								Else
									strHtml = strHtml &  "<tr class=clsSilver2>"
									blnCor = true
								End if	
								strHtml = strHtml &  "<td nowrap >&nbsp;"	& objRS("Acf_NroAcessoPtaEbt")	&"</td>"
								strHtml = strHtml &  "<td nowrap >&nbsp;"& objRS("Fac_Tronco")&"</td>"		
								strHtml = strHtml &  "<td nowrap>&nbsp;" & strRepresentacao	& "</td>"
								strHtml = strHtml &  "<td >&nbsp;<input type=button class=button name=btnLiberarFac value=Alocar onClick='LiberarFacilidade(" & objRS("Fac_Id") & ")'></td>"
								strHtml = strHtml &  "</tr>"
							End if	
						Case 3	
							if  Not objDic.Exists(strRepresentacao) then
								Call objDic.Add(strRepresentacao,objRS("Fac_Id")) 
								if blnCor then
									strHtml = strHtml &  "<tr class=clsSilver >"
									blnCor = false
								Else
									strHtml = strHtml &  "<tr class=clsSilver2>"
									blnCor = true
								End if	
								strHtml = strHtml &  "<td nowrap >&nbsp;"& objRS("Acf_NroAcessoPtaEbt")	&"</td>"
								strHtml = strHtml &  "<td nowrap >&nbsp;"& objRS("Fac_Tronco")	& "</td>"		
								strHtml = strHtml &  "<td nowrap>&nbsp;" & objRS("Fac_Par")		& "</td>"
								strHtml = strHtml &  "<td nowrap>&nbsp;" & objRS("Fac_Lateral")	& "</td>"
								strHtml = strHtml &  "<td nowrap>&nbsp;" & objRS("Fac_CxEmenda")	& "</td>"
								strHtml = strHtml &  "<td nowrap>&nbsp;" & strRepresentacao	& "</td>"
								strHtml = strHtml &  "<td >&nbsp;<input type=button class=button name=btnLiberarFac value=Alocar onClick='LiberarFacilidade(" & objRS("Fac_Id") & ")'></td>"
								strHtml = strHtml &  "</tr>"
							End if	
					End Select
					objRS.MoveNext
				Wend
				strHtml = strHtml &  "</table></td></tr>"
				Response.Write strHtml
			End if
	End if%>
	<tr class=clsSilver>
		<td colspan="2">
			<div id=divIDFis1 style="DISPLAY: 'none'">
				<table width=100%>
					<tr>
						<td >	
							<iframe	id			= "IFrmIDFis1"
								    name        = "IFrmIDFis1" 
								    width       = "100%" 
								    height      = "45px"
								    frameborder = "0"
								    scrolling   = "auto"
								    align       = "left">
							</iFrame>
						</td>
					</tr>		
				</table>	
			</div>	
		</td>
	</tr>
</table>
<table cellspacing=0 cellpadding=0 width=760 border=0>
	<tr><th colspan=6>&nbsp;•&nbsp;Facilidades do Pedido</th></tr>
</table>
<span id=spnCampos></span>
<table cellspacing=1 cellpadding=0 width=760 border=0>
	<tr class=clsSilver>
		<td >
			<iframe id=IFrmFacilidade 
					name=IFrmFacilidade 
					align=left 
					src="ListaFacilidades.asp" 
					frameBorder=0 
					width="100%" 
					BORDER=0
					height=45>
			</iframe>
		</td>
	</tr>
</table>
<span id=spnDet></span>
<span id=spnNDet></span>
<span id=spnAde></span>
<table border=0 cellspacing="0" cellpadding="0" width=760 >
	<tr>
		<td >
			<iframe	id			= "IFrmProcesso1"
				    name        = "IFrmProcesso1" 
				    width       = "760" 
				    height      = "18px"
				    frameborder = "0"
				    scrolling   = "no" 
				    align       = "left">
			</iFrame>
		</td>
	</tr>
</table>
<table width="760" border=0>
	<tr><td>
	<table width=50% border=0 align=center cellspacing=1 cellpadding=1>
		<tr class=clsSilver2>
			<td colspan=3 align=center><input type="button" class="button" name="btnOK" style="width:150px;height:22px" value="Alocar Facilidade(s)" onclick="GravarFacilidadeAux()" accesskey="I" onmouseover="showtip(this,event,'Alocar Facilidade(s)(Alt+I)');"></td>
		</tr>
		<tr>
			<td ><input type="button" class="button" style="width:150px" name="consultar" value="Posições Disponíveis" onClick="objAryFac[0][0]='';PosicoesLivre('L')" accesskey="L" onmouseover="showtip(this,event,'Consultar posições disponíveis (Alt+L)');"></td>
			<td ><input type="button" class="button" name="btnSair" value="Voltar" style="width:150px" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');" onClick="javascript:window.history.back()"></td>
			<td ><input type="button" class="button" name="btnInterligacao" value="Interligações" style="width:150px" accesskey="B" onmouseover="showtip(this,event,'Voltar (Alt+B)');" onClick="AbrirInterligacoes()"></td>
		</tr>	
	</table>
	</td>
	</tr>	
</table>
</Form>
<table width="760">
	<tr>
		<td>
			<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
		</td>
	</tr>
	<tr>
		<td>
			&nbsp;&nbsp;&nbsp;&nbsp;Legenda: A - Alfanumérico;  N - Numérico;  L - Letra
		</td>
	</tr>

</table>
<SCRIPT LANGUAGE=javascript>
<!--
var strDet = new String('<%=strDet%>')
var strNDet = new String('<%=strNDet%>')
var strAde = new String('<%=strAde%>')

var strCamposDet = new String('<%=strCamposDet%>')
var strCamposNDet = new String('<%=strCamposNDet%>')
var strCamposAde = new String('<%=strCamposAde%>')

function ResgatarInfoRede(){
<%
Select Case Cint("0" & strRede)
	Case 1 'Det
		Response.write "spnDet.innerHTML = '" & strDet &"';"
		Response.write "spnCampos.innerHTML = '" & strCamposDet &"';"
	Case 2 'NDet
		Response.Write "spnNDet.innerHTML = '" & strNDet &"';"
		Response.write "spnCampos.innerHTML = '" & strCamposNDet &"';"
	Case 3 'ADE
		Response.Write "spnAde.innerHTML = '" & strAde &"';"
		Response.write "spnCampos.innerHTML = '" & strCamposAde &"';"
End Select	
%>
}

function AbrirInterligacoes()
{
	with(document.forms[0])
	{
		if (cboRede.value ==1)
		{
			alert('Não é possível exibir interligações para este tipo de rede')
			return false
		}
	
		alert('Atenção, para ser possível a troca de interligação você deve 1o. clicar no botão (Alocar Facilidades)')
		var objAry = new Array(objXmlGeral,hdnIntIndice.value,hdnRecId.value,Ped_Id.value,0)
		var intRet = window.showModalDialog('ManobraInterligacao.asp?hdnPedId='+Ped_Id.value,objAry,'dialogHeight: 300px; dialogWidth: 700px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
		if (intRet == 1){
				intRet = 0
		}
	}
}

function MsgTrocaProvedor()
{
	alert('Atenção, para troca de provedor \nPrimero remova as facilidades existentes!')
	return false
}

function GravarFacilidadeAux()
{
	return GravarFacilidade()
}

function ResgatarInfoRedeEstoque(intRede)
{
	switch (parseInt(intRede))
	{
		case 1: //Det
			spnDet.innerHTML = strDet
			spnCampos.innerHTML = strCamposDet
			spnNDet.innerHTML = ''
			spnAde.innerHTML = ''
			break
		case 2: //NDet
			spnNDet.innerHTML = strNDet
			spnCampos.innerHTML = strCamposNDet
			spnDet.innerHTML = ''
			spnAde.innerHTML = ''
			break
		case 3: //ADE
			spnAde.innerHTML = strAde
			spnCampos.innerHTML = strCamposAde
			spnNDet.innerHTML = ''
			spnDet.innerHTML = ''
			break
	}
}

function PopularCamposFac(){
	<%if Trim(strProId) <> "" then %>
		if ('<%=strRede%>'=='3') ResgatarPadraoProvedor(document.forms[0].cboProvedor,1)
		try{
			document.forms[0].hdnProvedor.value = '<%=strProId%>'
			ResgatarProvedoresAssociados(document.forms[0].hdnProvedor)
		}catch(e){}	
	<%End if%>
}	

function LimparCompartilhamento()
{
	with (document.forms[0])
	{
		var objNodeE = objXmlGeral.selectNodes("//Facilidade[strTipoConsulta='C' || strTipoConsulta='E']")
		if (objNodeE.length > 0)
		{
			hdnIdAcessoFisico.value = "0"
			hdnIdAcessoFisico1.value = "0"
			hdnPropIdFisico.value = "0"
			hdnPropIdFisico1.value = "0"
			hdnCompartilhamento.value = "0"
			hdnCompartilhamento1.value = "0"
			hdnAlteracao.value = "0"
			hdnPodeAlterar.value = "S"
			divIDFis1.style.display = 'none'
			objXmlGeral.loadXML("<xDados/>")
			cboRede.disabled = false;
			AtualizarListaFac()
		}else{
			alert("Item de estoque/compartilhamento não selecionado.")
			return
		}	
	}
}

function LiberarFacilidade()
{
	with (document.forms[0])
	{
		var objXmlFacOrigemAux = new ActiveXObject("Microsoft.XMLDOM")
		objXmlFacOrigemAux.loadXML(objXmlFacOrigem.xml)

		var objNode = objXmlFacOrigemAux.selectNodes("//xDados/Facilidade[Fac_ID="+parseInt(arguments[0])+"]")

		if 	(objNode.length > 0)
		{
			var objNodeRequest = objNode[0].getElementsByTagName("Fac_Representacao")
			if (objNodeRequest.length != 0) //Tem representação
			{
				var stRepresentacao = objNodeRequest[0].childNodes[0].text
				
				for (var intIndex=0;intIndex<objNode.length;intIndex++)
				{
					var objNodeRepresentacao = objNode[intIndex].getElementsByTagName("Fac_Representacao")
					if (objNodeRepresentacao.length > 0)
					{
						if (stRepresentacao == objNodeRepresentacao[0].childNodes[0].text)
						{
							var objNodeListAtual = objNode[intIndex].getElementsByTagName("intFacSel")
							if (objNodeListAtual.length == 0)
							{
								objNodeFilho = objXmlFacOrigemAux.createNode("element", "intFacSel", "")
								objNodeFilho.text = 1
								objNode[intIndex].appendChild (objNodeFilho)
							}	
							else
							{
								objNodeListAtual.item(0).text = 1
							}
						}	
					}
				}
			}
			else
			{
				for (var intIndex=0;intIndex<objNode.length;intIndex++)
				{
					var objNodeListAtual = objNode[intIndex].getElementsByTagName("intFacSel")
					if (objNodeListAtual.length == 0)
					{
						objNodeFilho = objXmlFacOrigemAux.createNode("element", "intFacSel", "")
						objNodeFilho.text = 1
						objNode[intIndex].appendChild (objNodeFilho)
					}	
					else
					{
						objNodeListAtual.item(0).text = 1
					}
				}
			}	
		}else
		{
			alert("Facilidade não encontrada.")
			return
		}
		hdnAcao.value = "LiberarFacilidade"
		hdnXmlFacLibera.value = objXmlFacOrigemAux.xml
		target = "IFrmProcesso"
		action = "smbdProcessoFac.asp"
		submit()
	}
}
<%
if Trim(Request.Form("cboProvedor")) <> "" and  strRede = "" then
		Response.Write	"spnDet.innerHTML = strDet;"
		Response.Write	"spnCampos.innerHTML = strCamposDet;"
		Response.Write	"spnNDet.innerHTML = '';"
		Response.Write	"spnAde.innerHTML = '';"
End if
%>
//-->
</SCRIPT>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso"
	    width       = "100%"
	    height      = "100%"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
<iframe	id			= "IFrmProcesso2"
	    name        = "IFrmProcesso2" 
	    width       = "100%" 
	    height      = "100%"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
<iframe	id			= "IFrmProcesso3"
	    name        = "IFrmProcesso3" 
	    width       = "100%" 
	    height      = "100%"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>
</body>
</html>
<SCRIPT LANGUAGE=javascript>
<!--
var objXmlFacOrigem = new ActiveXObject("Microsoft.XMLDOM")
objXmlFacOrigem.loadXML("<%=strXmlFacOrigem%>") 
RetornaCboPlataforma(document.forms[0].cboRede.value,<% = strPla  %>)
if (document.forms[0].cboRede.disabled) document.forms[0].cboPlataforma.disabled = true 

//-->
</SCRIPT>
