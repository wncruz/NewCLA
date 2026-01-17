<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ProcessoMonitoracao.asp
'	- Descrição			: Lista as Solicitações na monitoração
%>
<!--#include file="../inc/data.asp"-->
<Html>
<Body topmargin=0 leftmargin=0 class=TA>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<form name=Form1 Method=post action="ProcessoEntregaAtivacao.asp">
<input type=hidden name=cboUsuario value="<%=Request.Form("cboUsuario")%>">
<input type=hidden name=cboProvedor value="<%=Request.Form("cboProvedor")%>">
<input type=hidden name=cboStatus value="<%=Request.Form("cboStatus")%>">
<input type=hidden name=txtPedNum value="<%=Request.Form("txtPedNum")%>">
<input type=hidden name=txtSolId value="<%=Request.Form("txtSolId")%>">
<input type=hidden name=txtFacID value="<%=Request.Form("txtFacID")%>">
<input type=hidden name=txtCliente value="<%=Request.Form("txtCliente")%>">
<input type=hidden name=txtEndereco value="<%=Request.Form("txtEndereco")%>">
<input type=hidden name=txtNroEnd value="<%=Request.Form("txtNroEnd")%>">
<input type=hidden name=txtComplemento value="<%=Request.Form("txtComplemento")%>">
<input type=hidden name=hdnGICL value="<%=Request.Form("hdnGICL")%>">
<input type=hidden name=hdnLocation value="<%=Request.Form("hdnLocation")%>">
<input type=hidden name=hdnStatus value="<%=Request.Form("hdnStatus")%>">
<%
Dim legfases(7,7,7)
legfases(0,0,0)="ATIVAR"
legfases(0,0,1)="Cálculo"
legfases(0,0,2)="Vistoria"
legfases(0,0,3)="Infraestrutura"
legfases(0,0,4)="Instalação"
legfases(0,0,6)="Finalização"
legfases(0,0,5)="Finalizado"
legfases(0,1,0)="DESATIVAR"
legfases(0,1,2)="Vistoria"
legfases(0,1,3)="Infraestrutura"
legfases(0,1,4)="Desinstalação"
legfases(0,1,5)="Finalização"
legfases(0,1,6)="Finalizado"
legfases(0,2,0)="ALTERAR"
legfases(0,2,1)="Cálculo"
legfases(0,2,2)="Vistoria"
legfases(0,2,3)="Infraestrutura"
legfases(0,2,4)="Alteração"
legfases(0,2,5)="Finalização"
legfases(0,2,6)="Finalizado"
legfases(1,0,0)="ATIVAR"
legfases(1,0,1)="Survey"
legfases(1,0,2)="Projeto"
legfases(1,0,3)="Aprovação Projeto"
legfases(1,0,4)="Construção"
legfases(1,0,5)="Instalação"
legfases(1,0,6)="Finalização"
legfases(1,0,7)="Finalizado"
legfases(1,1,0)="DESATIVAR"
legfases(1,1,1)="Survey"
legfases(1,1,2)="Projeto"
legfases(1,1,3)="Aprovação Projeto"
legfases(1,1,4)="Construção"
legfases(1,1,5)="Desinstalação"
legfases(1,1,6)="Finalização"
legfases(1,1,7)="Finalizado"
legfases(1,2,0)="ALTERAR"
legfases(1,2,1)="Survey"
legfases(1,2,2)="Projeto"
legfases(1,2,3)="Aprovação Projeto"
legfases(1,2,4)="Construção"
legfases(1,2,5)="Alteração"
legfases(1,2,6)="Finalização"
legfases(1,2,7)="Finalizado"
legfases(2,0,0)="ATIVAR"
legfases(2,0,1)="Site Survey"
legfases(2,0,2)="Infraestrutura"
legfases(2,0,3)="Instalação"
legfases(2,0,4)="Testes"
legfases(2,0,5)="Ativação"
legfases(2,0,6)="Finalização"
legfases(2,0,7)="Finalizado"
legfases(2,1,0)="DESATIVAR"
legfases(2,1,1)="Vistoria Técnica"
legfases(2,1,2)="Infraestrutura"
legfases(2,1,5)="Desinstalação"
legfases(2,1,6)="Finalização"
legfases(2,1,7)="Finalizado"
legfases(2,2,0)="ALTERAR"
legfases(2,2,1)="Vistoria Técnica"
legfases(2,2,2)="Infraestrutura"
legfases(2,2,5)="Alteração"
legfases(2,2,6)="Finalização"
legfases(2,2,7)="Finalizado"

'' LPEREZ 13/12/2005
Dim dblUsuIdMonit		
Dim dblProId 
Dim dblStsId	
Dim strStatus
Dim intIndex
Dim strSql
Dim intCount
Dim strClass
Dim objRSCRM
Dim objRSCLA
Dim blnCRM
Dim blnCla
Dim strCliNome
Dim strSerDesc
Dim strDescAcessoLog
Dim strDescAcao
Dim strPropAcesso
Dim intTecno
Dim intAcao
Dim strFase
Dim intFase
Dim objDicLocalidades
Dim dblSolId
Dim dblCNL
Dim dblPPrefixo
Dim dblPNumero
Dim dblPAno
Dim dblFacID
Dim dblCliente
Dim dblEndereco
Dim dblNroEnd
Dim dblCompl

dblUsuIdMonit	= Trim(Request.Form("cboUsuario"))
dblProId		= Trim(Request.Form("cboProvedor"))
dblStsId		= Trim(Request.Form("hdnStatus"))

dblSolId	= Request.form("txtSolId")
'dblCNL 	= Request.Form("cboCNL")
dblPPrefixo = mid(Request.form("txtPedNum"), 1, 2)
dblPNumero 	= mid(Request.form("txtPedNum"), 4, 5)
dblPAno		= mid(Request.form("txtPedNum"), 10, 4)
dblFacID 	= Request.form("cboFacID")
dblCliente	= Request.form("txtCliente")
dblEndereco = Request.form("txtEndereco")
dblNroEnd	= Request.form("txtNroEnd")
dblCompl	= Request.form("txtComplemento")
txtFacID = Request.form("txtFacID") '--> psouto 12/05/2006

'Dim dblUsuIdMonit		
'Dim dblProId 
'Dim dblStsId	
'Dim strStatus
'Dim intIndex
'Dim strSql
'Dim intCount
'Dim strClass
'Dim objRSCRM
'Dim objRSCLA
'Dim blnCRM
'Dim blnCla
'Dim strCliNome
'Dim strSerDesc
'Dim strDescAcessoLog
'Dim strDescAcao
'Dim strPropAcesso
'Dim intTecno
'Dim intAcao
'Dim strFase
'Dim intFase
'Dim objDicLocalidades
'
'dblUsuIdMonit	= Trim(Request.Form("cboUsuario"))
'dblProId		= Trim(Request.Form("cboProvedor"))
'dblStsId		= Trim(Request.Form("hdnStatus"))

''LP
if dblUsuIdMonit = "" then dblUsuIdMonit	=	"null" End if
if dblProId = ""	  then dblProId=	"null" End if
if dblStsId = ""	  then dblStsId	=	"null" End if

'Carrega objeto dictionary com todas as localidades que o usuário possui.
CALL CarregaLocalidades()
'' LPEREZ 13/12/2005

	Vetor_Campos(1)="adInteger,4,adParamInput,"		&  dblFacID			'@Acl_IDAcessoLogico
	Vetor_Campos(2)="adInteger,4,adParamInput," 	&  dblUsuIdMonit	'@Usu_ID
	Vetor_Campos(3)="adInteger,4,adParamInput,"		&  dblProId			'@Pro_ID
	Vetor_Campos(4)="adInteger,4,adParamInput,"		&  dblStsId 		'@Sts_ID
	Vetor_Campos(5)="adWChar,1,adParamInput,T"	  						'@Situacao
	Vetor_Campos(6)="adWChar,1,adParamInput,P"							'@Agp_Origem
	Vetor_Campos(7)="adInteger,4,adParamInput," 	&  dblUsuId			'@UsuID_Logado
	Vetor_Campos(8)="adInteger,4,adParamInput," 		'@CNL
	Vetor_Campos(9)="adWChar,2,adParamInput,"		&  dblPPrefixo		'@Ped_Prefixo
	Vetor_Campos(10)="adInteger,4,adParamInput,"	&  dblPNumero		'@Ped_Numero
	Vetor_Campos(11)="adInteger,4,adParamInput,"	&  dblPAno			'@Ped_Ano
	Vetor_Campos(12)="adInteger,4,adParamInput,"	&  dblSolID			'@sol_id
	Vetor_Campos(13)="adInteger,60,adParamInput,"	&  dblCliente		'@Cli_Nome
	Vetor_Campos(14)="adInteger,60,adParamInput,"	&  dblEndereco 		'@End_NomeLogr
	Vetor_Campos(15)="adInteger,10,adParamInput,"	&  dblEndNro		'@End_NroLogr
	Vetor_Campos(16)="adInteger,30,adParamInput,"	&  dblComplemento	'@Aec_Complemento
	Vetor_Campos(17)="adInteger,10,adParamInput,"	&  txtFacID			'@End_NroLogr --> PSOUTO 12/05/2006
	
	

strSql = APENDA_PARAMSTRSQL("CLA_sp_view_monitoracao2 ",17,Vetor_Campos)

Call PaginarRS(1,strSql)

intCount=1
if not objRSPag.Eof and not objRSPag.Bof then
	For intIndex = 1 to objRSPag.PageSize

	Set objRSCRM = Server.CreateObject("ADODB.Recordset")
	Set objRSCLA = Server.CreateObject("ADODB.Recordset")
	Set objRSFis = Server.CreateObject("ADODB.Recordset")

	if (intCount mod 2) <> 0 then strClass = "clsSilver" else strClass = "clsSilver2" End if
%>
	<table	border="0" cellspacing="1" cellpadding="0" width="760" align="center" >
		<tr class="<%=strClass%>">
			<%
			if len(Trim(objRSPag("Cli_Nome"))) > 25 then
				strCliNome = TratarAspasHtml(Left(objRSPag("Cli_Nome"),25) & "...")
			Else
				strCliNome = TratarAspasHtml(objRSPag("Cli_Nome"))
			End if	
			if Len(Trim(objRSPag("Ser_desc"))) > 10 then
				strSerDesc = TratarAspasHtml(Left(objRSPag("Ser_desc"),10) & "...")
			Else
				strSerDesc = TratarAspasHtml(objRSPag("Ser_desc"))
			End if	

			if Len(Trim(objRSPag("DescVelAcessoLog"))) > 10 then
				strDescAcessoLog = TratarAspasHtml(Left(objRSPag("DescVelAcessoLog"),10) & "...")
			Else
				strDescAcessoLog = TratarAspasHtml(objRSPag("DescVelAcessoLog"))
			End if	

			if Len(Trim(objRSPag("Tprc_Des"))) > 12 then
				strDescAcao = TratarAspasHtml(Left(objRSPag("Tprc_Des"),12) & "...")
			Else
				strDescAcao = TratarAspasHtml(objRSPag("Tprc_Des"))
			End if	

			if Len(Trim(objRSPag("Sts_Desc"))) > 19 then
				strStatus = TratarAspasHtml(Left(objRSPag("Sts_Desc"),19) & "...")
			Else
				strStatus = TratarAspasHtml(objRSPag("Sts_Desc"))
			End if	
			%>
			<td width=34px><a href="javascript:parent.DetalharItem(<%=objRSPag("Sol_id")%>)">&nbsp;<b><%=objRSPag("Sol_id")%></b></a></td>
			<td width=80px >&nbsp;<%=TratarAspasHtml(objRSPag("Acl_IDAcessoLogico"))%></td>
			<td width=200px><span id=spnCliente onmouseover="showtip(this,event,'<%=TratarAspasJS(Trim(objRSPag("Cli_Nome")))%>');" onmouseout="hidetip();">&nbsp;<%=strCliNome%></span></td>
			<td width=90px ><span id=spnServico onmouseover="showtip(this,event,'<%=TratarAspasJS(Trim(objRSPag("Ser_desc")))%>');" onmouseout="hidetip();">&nbsp;<%=strSerDesc%></span></td>
			<td width=90px nowrap><span id=spnVelAcessoLog onmouseover="showtip(this,event,'<%=TratarAspasJS(Trim(objRSPag("DescVelAcessoLog")))%>');" onmouseout="hidetip();">&nbsp;<%=strDescAcessoLog%></span></td>
			<td width=100px nowrap><span id=spnDescAcao onmouseover="showtip(this,event,'<%=TratarAspasJS(Trim(objRSPag("Tprc_Des")))%>');" onmouseout="hidetip();">&nbsp;<%=strDescAcao%></span></td>
			<td width=165px ><span id=spnStatus onmouseover="showtip(this,event,'<%=TratarAspasJS(Trim(objRSPag("Sts_Desc")))%>');" onmouseout="hidetip();">&nbsp;<%=strStatus%></span></td>
		</tr>
	<%

	
	
	'if Trim(Request.Form("hdnGICL")) <> "" and Request.Form("hdnLocation") = "MONITORACAO" and objDicLocalidades.Exists(cstr(objRSPag("Esc_IDEntrega"))) then

		Set objRSAnt = db.Execute("select Acl_AntAcesso from cla_acessologico where Acl_IDAcessologico = " & Trim(objRSPag("Acl_IDAcessoLogico")))
		
		if not objRSAnt.bof and not objRSAnt.eof then
		  var_Antecipacao = objRSAnt("Acl_AntAcesso")
		end if
		
		Vetor_Campos(1)="adInteger,4,adParamInput," & Trim(objRSPag("Ser_ID"))
		
		strSql = APENDA_PARAMSTRSQL("CLA_sp_sel_designacaoServico ",1,Vetor_Campos)
		
		Set objRS = db.Execute(strSql)
				
		If Not objRS.eof and not objRS.Bof Then
			siglaServico = objRS("Ser_Sigla")
		end if		
		
		Select Case Cdbl("0" & objRSPag("Sts_ID"))
			Case 52
				Response.Write "<tr class=" & strClass & "><td colspan=7 align=right><span id='spn" & Trim(objRSPag("Acl_IDAcessoLogico")) & "' >&nbsp;•&nbsp;Data de Entrega do Acesso ao Serviço&nbsp;"
				Response.Write "<input type=button class=button name=btnAddData value=Adicionar onclick=""AtualizarStatus('dataentrega','" & Trim(objRSPag("Sol_ID")) & "', '" & Trim(objRSPag("Sts_ID")) & "','" & Trim(objRSPag("Acl_IDAcessoLogico")) & "','" & Trim(strPropAcesso) & "','" & Trim(objRSPag("Tprc_ID"))& "','" & Trim(objRSPag("orisol_id")) & "')""  ></span></td></tr>"
			Case 53
				Response.Write "<tr class=" & strClass & "><td colspan=7 align=right><span id='spn" & Trim(objRSPag("Acl_IDAcessoLogico")) & "' >&nbsp;•&nbsp;Pendente de Ativação do Serviço&nbsp;"
				Response.Write "<input type=button class=button name=btnAddAtivacao value=Adicionar onclick=""AtualizarStatus('pendenteativacao','" & Trim(objRSPag("Sol_ID")) & "', '" & Trim(objRSPag("Ser_ID")) & "','" & Trim(objRSPag("Ser_Desc")) & "','" & Trim(objRSPag("Acl_IDAcessoLogico")) & "','" & Trim(objRSPag("Sts_ID")) & "','" & Trim(strPropAcesso) & "','" & Trim(objRSPag("Tprc_ID")) & "','" & Trim(objRSPag("orisol_id")) & "','" & Trim(var_Antecipacao) & "','" & Trim(siglaServico) & "')""></span></td></tr>"
		End Select
	'End if	

	blnCRM = false
	blnCla = false

	Set objRSCRM = db.execute("CLA_SP_SEL_CRMSPROCESSO " & objRSPag("Acl_IDAcessoLogico") & "," & objRSPag("Sol_id"))
	if not objRSCRM.EOF and Not objRSCRM.BOF then
		blnCRM = true
		strPedCRMS = "<table border=0 cellspacing=0 cellpadding=0 width=760 align=center>"
		strPedCRMS = strPedCRMS & "<tr class=" & strClass & ">"
		strPedCRMS = strPedCRMS & "<td  width=100px >&nbsp;Nº Pedido</td>"
		strPedCRMS = strPedCRMS & "<td  width=100px >&nbsp;Data</td>"
		strPedCRMS = strPedCRMS & "<td  width=100px >&nbsp;Ação</td>"
		strPedCRMS = strPedCRMS & "<td  width=100px nowrap>&nbsp;Nº do Acesso</td>"
		strPedCRMS = strPedCRMS & "<td  width=100px >&nbsp;Velocidade</td>"
		strPedCRMS = strPedCRMS & "<td  width=100px >&nbsp;Status</td>"
		strPedCRMS = strPedCRMS & "<td  width=100px >&nbsp;Provedor</td>"
		strPedCRMS = strPedCRMS & "<td  width=60px >&nbsp;Acesso</td>"
		strPedCRMS = strPedCRMS & "</tr>"
	End if

	Vetor_Campos(1)="adInteger,4,adParamInput,"
	Vetor_Campos(2)="adInteger,4,adParamInput,"
	Vetor_Campos(3)="adInteger,4,adParamInput," & objRSPag("Sol_ID")
			
	strSqlRet = APENDA_PARAMSTRSQL("CLA_SP_VIEW_ACESSOFISICO",3,Vetor_Campos)

	Set objRSFis = db.Execute(strSqlRet)

	if Not objRSFis.EOF and not objRSFis.BOF then
		blnCla = true
		strPedCla = "<table	 cellspacing=0 cellpadding=0 width=760 align=center >"
		strPedCla = strPedCla & "<tr class=" & strClass & ">"
		strPedCla = strPedCla & "<td  width=100px>&nbsp;Nº Pedido</td>"
		strPedCla = strPedCla & "<td  width=100px>&nbsp;Data</td>"
		strPedCla = strPedCla & "<td  width=100px>&nbsp;Ação</td>"
		strPedCla = strPedCla & "<td  width=100px nowrap>&nbsp;Nº do Acesso</td>"
		strPedCla = strPedCla & "<td  width=100px>&nbsp;Velocidade</td>"
		strPedCla = strPedCla & "<td  width=100px>&nbsp;Status</td>"
		strPedCla = strPedCla & "<td  width=100px>&nbsp;Provedor</td>"
		strPedCla = strPedCla & "<td  width=60px>&nbsp;Acesso</td>"
		strPedCla = strPedCla & "</tr>"

	End if	

	'Botao Adicionar Status 
	if Trim(Request.Form("hdnGICL")) <> "" and Request.Form("hdnLocation") = "CONSULTA" and  objDicLocalidades.Exists(cstr(objRSPag("Esc_IDEntrega"))) then
		Response.Write "<tr class=" & strClass & "><td colspan=8 align=right><span id='spn" & Trim(objRSPag("Acl_IDAcessoLogico")) & "' >&nbsp;"
		Response.Write "<input type=button class=button name=btnAddStatus value=""Adicionar Status"" onClick=""AdicionarStatus(" & Trim(objRSPag("Sol_ID")) & ")""></span></td></tr>"
	End if
	%>
	</table>
	<%
	if blnCRM then
		Response.Write strPedCRMS
		While Not objRSCRM.Eof%>
		<tr class="<%=strClass%>">
			<td width=100px>&nbsp;<%=objRSCRM("CodInstalar")%></td>
			<%
					Select Case Ucase(Trim(objRSCRM("Acesso")))
						Case "RADIO"
							intTecno = 0
						Case "FIBRA"	
							intTecno = 1
						Case "SATELITE"
							intTecno = 2	
					End Select								
							
					Select Case Ucase(Trim(objRSCRM("ACAO")))
						Case "ATIVAR"
							intAcao = 0
						Case "DESATIVAR"	
							intAcao = 1
						Case "ALTERAR"
							intAcao = 2	
						Case "CANCELAR"
							intAcao = 3	
					End Select								
							
					Select Case objRSCRM("fase")
						Case -1
							strFase = "Em análise"
						Case Else
							intFase = objRSCRM("fase")	
							strFase = legfases(intTecno,intAcao,intFase)
					End Select
			%>
			<td  width=100px>&nbsp;<%=split(Formatar_Data(objRSCRM("Entrada"))," ")(0)%></td>
			<td  width=100px>&nbsp;<%=TratarAspasHtml(objRSCRM("Acao"))%></td>
			<td  width=100px>&nbsp;<%=objRSCRM("Acf_NroAcessoPtaEbt")%></td>
			<td  width=100px>&nbsp;<%=TratarAspasHtml(objRSCRM("VelocidadeAcesso"))%></td>
			<td  width=100px>&nbsp;<%=strFase%></td>
			<td  width=100px>&nbsp;EMBRATEL</td>
			<td  width=60px>&nbsp;<%=TratarAspasHtml(objRSCRM("Acesso"))%></td>
		</tr>
		<tr></tr>
		<%
			if Trim(Request.Form("hdnGICL")) <> ""  and Request.Form("hdnLocation") = "MONITORACAO" then
				if Not isNull(objRSCRM("Pedidoaprovainfra")) and isNull(objRSCRM("Dataaprovainfra")) then
						Response.Write "<tr class="& strClass & "><td colspan=8 align=right>&nbsp;•&nbsp;Aprovar Infra&nbsp;<span id='spn" & Trim(objRSCRM("IdLog")) & "' ><input type='button' class='button' name='" & Trim(objRSCRM("IdLog")) & "' value='Aprovar Infra' onclick=""AtualizarStatus('infra','" & Trim(objRSCRM("OrcadoInfra")) & "','" & Trim(objRSCRM("ValorInfra")) & "','" & Trim(objRSCRM("IdLog")) & "','" & Trim(strPropAcesso) & "')"" ></span></td></tr><tr></tr>"
				End if
			End If	
			objRSCRM.MoveNext
		Wend
		%>
	</table>
	<%
	End if

	if blnCla then
		While not objRSFis.EOF 

			if Ucase(Trim(objRSFis("Acf_Proprietario"))) = "TER" or Ucase(Trim(objRSFis("Acf_Proprietario"))) = "CLI" or (Ucase(Trim(objRSFis("Acf_Proprietario"))) = "EBT" and Ucase(Trim(objRSFis("Tec_Sigla"))) = "ADE") then
				if blnCla then 
					Response.Write strPedCla 
					blnCla = false
				End if	

				''@@ Davif - Otimizacao
				''Set objRSSts = db.Execute("CLA_sp_sel_StatusSolicitacao null,null,3," & objRSFis("Ped_id"))
				''if Not objRSSts.Eof and Not objRSSts.Bof then strStatusDet = objRSSts("Sts_Desc") else strStatusDet = "" End if
				''@@

				'strStatusDet = Trim(objRSFis("StatusDetalhado"))
				strVelocidade = Trim(objRSFis("Vel_Desc")) & " " & TipoVel(objRSFis("Acf_TipoVel"))
				%>
				<tr class="<%=strClass%>">
					<td width=100px>&nbsp;<%=ucase(Trim(objRSFis("Ped_Prefixo"))) & "-" & Trim(right("00000" & objRSFis("Ped_Numero"),5)) & "/" & Trim(objRSFis("Ped_Ano"))%></td>
					<td width=100px>&nbsp;<%=split(Formatar_Data(objRSFis("Ped_Data"))," ")(0)%></td>
					<td width=100px>&nbsp;<%=AcaoPedido(objRSFis("Tprc_Id"))%></td>
					<td width=100px>&nbsp;<%=objRSFis("Acf_NroAcessoPtaEBT")%></td>
					<td width=100px>&nbsp;<span id=spnGeral onmouseover="showtip(this,event,'<%=strVelocidade%>')" onmouseout='hidetip();'><%=FormatarCampo(strVelocidade,11)%></td>
					<td width=100px>&nbsp;<span id=spnGeral onmouseover="showtip(this,event,'<%=strStatusDet%>')" onmouseout='hidetip();'><%=FormatarCampo(strStatusDet,11)%></span></td>
					<td width=100px>&nbsp;<span id=spnGeral onmouseover="showtip(this,event,'<%=TratarAspasHtml(objRSFis("Pro_Nome"))%>')" onmouseout='hidetip();'><%=FormatarCampo(objRSFis("Pro_Nome"),12)%></td>
					<td width=60px>&nbsp;<%=Ucase(Trim(objRSFis("Tec_Sigla")))%></td>
				</tr>
				<tr></tr>
				<%
			End if
			objRSFis.MoveNext
		Wend
	%>
	</table>
	<%
	End if
		intCount = intCount+1
		objRSPag.MoveNext
		if objRSPag.EOF then Exit For
	Next
else
  response.write "<br><center><font color=ff0000><li> Registro(s) não encontrado(s).</font></center>"
End if
%>
<iframe	id			= "IFrmProcesso"
	    name        = "IFrmProcesso" 
	    width       = "0" 
	    height      = "0"
	    frameborder = "0"
	    scrolling   = "no" 
	    align       = "left">
</iFrame>

<script language="JavaScript">
function AtualizarStatus()
{
	if (arguments.length > 0)
	{
		var objAryParam = new Array()
		for (var intIndex=0;intIndex<arguments.length;intIndex++)
		{
			objAryParam[intIndex] = arguments[intIndex]
		}	
		switch (arguments[0])
		{
			case "infra":
				strRet = window.showModalDialog("AtualizarStatus.asp",objAryParam,"dialogHeight: 150px; dialogWidth: 350px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
				if (strRet != undefined)
				{
					eval("spn"+strRet+".innerHTML = ''")
					document.forms[0].submit()
				}	
				break			
			case "dataentrega":
				strRet = window.showModalDialog("AtualizarStatus.asp",objAryParam,"dialogHeight: 150px; dialogWidth: 350px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
				if (strRet != undefined)
				{
					eval("spn"+strRet+".innerHTML = ''")
					document.forms[0].submit()
				}	
				break
			
			case "pendenteativacao":
				
				
				if (objAryParam[7] == "3")
				{
					strRet = window.showModalDialog("AtualizarStatus.asp",objAryParam,"dialogHeight: 300px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
				}
				else
				{
					strRet = window.showModalDialog("AtualizarStatus.asp",objAryParam,"dialogHeight: 200px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;");
				}	
				
				if (strRet != undefined)
				{
					eval("spn"+strRet+".innerHTML = ''")
					document.forms[0].submit()
				}	
				
				if (strRet != undefined)
				{
					eval("spn"+strRet+".innerHTML = ''")
					if(arguments[8]== 3){
						var proseg = 0
						
						do{
							strRet = window.showModalDialog("processomonitoracao_alteradatas.asp?sol_id=" + objAryParam[1],objAryParam,"dialogHeight: 180px; dialogWidth: 350px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No; scroll: No");
							if(strRet != undefined){
								proseg = 1	
							}
						}while(proseg == 0 )
						
					}
					
					document.forms[0].submit()
				}
				break
		}
	}
}

function AdicionarStatus()
{
	var intRet = window.showModalDialog('AddStatus.asp?hdnSolId='+arguments[0],"",'dialogHeight: 200px; dialogWidth: 600px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No;');
}
</script>
<input type=hidden name=hdnNomeSpn>
<input type=hidden name=hdnAcao>
<input type=hidden name=hdnIdLog>
<!--#include file="../inc/ControlesPaginacao.asp"-->
</form>
</body>
</html>

<% 

private sub CarregaLocalidades()

Dim  objLocal

Set objDicLocalidades = Server.CreateObject("Scripting.Dictionary")
Set objLocal = db.execute("CLA_sp_sel_usuarioesc " & dblUsuId & ",4")

While Not objLocal.Eof
	if Not objDicLocalidades.Exists(cstr(objLocal("Esc_ID"))) then
		Call objDicLocalidades.Add (cstr(objLocal("Esc_ID")),objLocal("Esc_ID"))
	End if
	objLocal.MoveNext
Wend

end sub 
%>