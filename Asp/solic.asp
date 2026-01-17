<!--#include file="../inc/data.asp"-->
<%
ON ERROR RESUME NEXT
Server.ScriptTimeout = 8000 'Segundos

strSQL = ""
strSQL = strSQL & "select solic_ID,TNS,RAZAO_SOCIAL,NOME_FANTASIA,cli_cc,SUB_CONTA,SERVICO,VELOCIDADE,(select vel_id from cla_velocidade where vel_desc = VELOCIDADE) as VEL_ID, "
strSQL = strSQL & "(case when substring(COD_DESIGNACAO_SERVICO,4,1) = '/' then replace(COD_DESIGNACAO_SERVICO,'/IP',' /IP') else COD_DESIGNACAO_SERVICO end) as COD_DESIGNACAO_SERVICO,"

strSQL = strSQL & "TIPO_CONTRATO,NUMERO_CONTRATO_SERVICO,DATA_DE_ENTREGA_DO_SERVICO,INC_ORIGEM, "

strSQL = strSQL & "(select Pro_ID from cla_provedor where pro_nome = "
strSQL = strSQL & "  ( "
strSQL = strSQL & "  replace( "
strSQL = strSQL & "    replace( "
strSQL = strSQL & "      replace( "
strSQL = strSQL & "        replace( "
strSQL = strSQL & "          PROVEDOR,' -','-' "
strSQL = strSQL & "               ) "
strSQL = strSQL & "             ,'- ','-' "
strSQL = strSQL & "             )"
strSQL = strSQL & "           ,'OI-DIAL','OI DIAL' "
strSQL = strSQL & "           )"
strSQL = strSQL & "         ,'BAKUP','BACKUP' "
strSQL = strSQL & "         )"
strSQL = strSQL & "   )"
strSQL = strSQL & ") as PRO_ID,"

strSQL = strSQL & "( replace("
strSQL = strSQL & "    replace("
strSQL = strSQL & "      replace("
strSQL = strSQL & "        replace("
strSQL = strSQL & "          PROVEDOR,' -','-'"
strSQL = strSQL & "               )"
strSQL = strSQL & "             ,'- ','-'"
strSQL = strSQL & "             )"
strSQL = strSQL & "           ,'OI-DIAL','OI DIAL'"
strSQL = strSQL & "           )"
strSQL = strSQL & "         ,'BAKUP','BACKUP'"
strSQL = strSQL & "         )"
strSQL = strSQL & ") as PROVEDOR, "

strSQL = strSQL & "(SELECT esc_id from cla_estacao where Cid_sigla = substring(ESTACAO_ENTREGA1,1,4) and esc_sigla = substring(ESTACAO_ENTREGA1,5,3)) as ESTACAO_ESC_ID, "
strSQL = strSQL & "(SELECT esc_id from cla_estacao where Cid_sigla = substring(SIGLA_ESTACAO_CONFIG,1,4) and esc_sigla = substring(SIGLA_ESTACAO_CONFIG,5,3)) as SIGLA_ESC_ID, "

strSQL = strSQL & "(SELECT TPL_SIGLA + ' ' + ESC_NOMELOGR + ', ' + ESC_NROLOGR + ' ' + ESC_BAIRRO + ' ' + ESC_COD_CEP  from cla_estacao where Cid_sigla = substring(ESTACAO_ENTREGA,1,4) and esc_sigla = substring(ESTACAO_ENTREGA,5,3)) as ESTACAO_ENDERECO, "


strSQL = strSQL & "TECNOLOGIA, "
strSQL = strSQL & "VELOCIDADE_FISICO,TIPO_ENDERECO,UF, "

strSQL = strSQL & "CIDADE, "
strSQL = strSQL & "(select top 1 cid_sigla from cla_cidade where cid_desc = CIDADE and Est_Sigla = UF) as CNL, "

strSQL = strSQL & "TIPO_DE_LOGRADOURO,NOME_DO_LOGRADOURO,NUMERO_DO_PREDIO, "
strSQL = strSQL & "BAIRRO,CEP,NOME_CONTATO, "

strSQL = strSQL & "substring( "
strSQL = strSQL & "  replace( "
strSQL = strSQL & "    replace( "
strSQL = strSQL & "      replace("
strSQL = strSQL & "             TELEFONE_CONTATO,' ','' "
strSQL = strSQL & "             ) "
strSQL = strSQL & "           ,'-','' "
strSQL = strSQL & "           ) "
strSQL = strSQL & "         ,'/','' "
strSQL = strSQL & "         ) "
strSQL = strSQL & "       ,1,10 "
strSQL = strSQL & "       ) as TELEFONE_CONTATO, "
strSQL = strSQL & "CNPJ, "

strSQL = strSQL & "INTERFACE_USUARIO,ESTACAO_ENTREGA,INTERFACE_EMBRATEL, "
strSQL = strSQL & "ESTACAO_ENTREGA1,SIGLA_ESTACAO_CONFIG,ORGAO_VENDA,USERNAME_GICN,USERNAME_GICL,gravado,erro,solic_ID,sol_id "

strSQL = strSQL & "FROM CLA_carga_solic "

strSQL = strSQL & "where gravado is null "
strSQL = strSQL & "and TELEFONE_CONTATO is not null "
strSQL = strSQL & "and len( "
strSQL = strSQL & "        substring( "
strSQL = strSQL & "          replace( "
strSQL = strSQL & "            replace( "
strSQL = strSQL & "              replace( "
strSQL = strSQL & "                TELEFONE_CONTATO,' ','' "
strSQL = strSQL & "                ) "
strSQL = strSQL & "              ,'-','' "
strSQL = strSQL & "              ) "
strSQL = strSQL & "            ,'/','' "
strSQL = strSQL & "            ) "
strSQL = strSQL & "          ,1,10 "
strSQL = strSQL & "          ) "
strSQL = strSQL & "        ) in (9,10) "
strSQL = strSQL & "and RAZAO_SOCIAL is not null "
strSQL = strSQL & "and NOME_FANTASIA is not null "
strSQL = strSQL & "and CLI_CC is not null "
strSQL = strSQL & "and SUB_CONTA is not null "
strSQL = strSQL & "and SERVICO is not null "
strSQL = strSQL & "and VELOCIDADE is not null "
strSQL = strSQL & "and TIPO_CONTRATO is not null "
'strSQL = strSQL & "and COD_DESIGNACAO_SERVICO is not null "
strSQL = strSQL & "and VELOCIDADE_FISICO is not null "
strSQL = strSQL & "and TIPO_ENDERECO is not null "
strSQL = strSQL & "and UF is not null "
strSQL = strSQL & "and CIDADE is not null "
strSQL = strSQL & "and TIPO_DE_LOGRADOURO is not null "
strSQL = strSQL & "and NOME_DO_LOGRADOURO is not null  "
strSQL = strSQL & "and NUMERO_DO_PREDIO is not null  "
strSQL = strSQL & "and BAIRRO is not null "
strSQL = strSQL & "and CEP is not null "
strSQL = strSQL & "and NOME_CONTATO is not null "
strSQL = strSQL & "and CNPJ is not null  "
strSQL = strSQL & "and INTERFACE_USUARIO is not null "
strSQL = strSQL & "and ESTACAO_ENTREGA is not null  "
strSQL = strSQL & "and INTERFACE_EMBRATEL is not null "
strSQL = strSQL & "and ESTACAO_ENTREGA1 is not null  "
strSQL = strSQL & "and SIGLA_ESTACAO_CONFIG is not null  "
strSQL = strSQL & "and ORGAO_VENDA is not null "
strSQL = strSQL & "and USERNAME_GICN is not null "
strSQL = strSQL & "and USERNAME_GICL is not null "
strSQL = strSQL & "and (select top 1 cid_sigla from cla_cidade where cid_desc = CIDADE and Est_Sigla = UF)  is not null "
strSQL = strSQL & "order by solic_id "

SET rs = Server.CreateObject("ADODB.Recordset")
rs.Open strSQL, DB

Do Until rs.EOF = True 
    msg_erro = ""
	
	strXML = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
	strXML = strXML & "<xDados>"
	strXML = strXML & "<Acesso>"
	strXML = strXML & "<intIndice>0</intIndice>"
	strXML = strXML & "<hdnAcao>ResgatarCidadeCNL</hdnAcao>"
	strXML = strXML & "<hdnProvedor>"& rs("provedor") &"</hdnProvedor>"
	strXML = strXML & "<hdnCNLNome>txtEndCid</hdnCNLNome>"
	strXML = strXML & "<hdnUFAtual>"& rs("uf") &"</hdnUFAtual>"
	strXML = strXML & "<hdnCNLAtual>"& rs("cnl") &"</hdnCNLAtual>"
	strXML = strXML & "<hdnNomeTxtCidDesc>txtEndCidDesc</hdnNomeTxtCidDesc>"
	strXML = strXML & "<hdnNomeCboCid>EndCid</hdnNomeCboCid>"
	strXML = strXML & "<hdnUserGICL>"& rs("username_gicl") &"</hdnUserGICL>"
	strXML = strXML & "<hdnCtfcId>2932</hdnCtfcId>"
	strXML = strXML & "<hdnRazaoSocial>"& rs("razao_social") &"</hdnRazaoSocial>"
	strXML = strXML & "<hdnCompartilhamento>0</hdnCompartilhamento>"
	strXML = strXML & "<hdnNodeCompartilhado>0</hdnNodeCompartilhado>"
	strXML = strXML & "<hdnCompartilhamento1>0</hdnCompartilhamento1>"
	strXML = strXML & "<hdnEstacaoDestino>"& rs("estacao_entrega") &"</hdnEstacaoDestino>"'  #SPO     IG        #
	strXML = strXML & "<hdnObrigaGla>1</hdnObrigaGla>"
	strXML = strXML & "<hdnUsuID>1225</hdnUsuID>"'migracao
	strXML = strXML & "<hdnTipoProcesso>1</hdnTipoProcesso>"
	strXML = strXML & "<rdoPropAcessoFisico>TER</rdoPropAcessoFisico>"
	strXML = strXML & "<rdoPropAcessoFisicoIndex>0</rdoPropAcessoFisicoIndex>"
	strXML = strXML & "<cboTecnologiaText>:: TECNOLOGIA EBT</cboTecnologiaText>"
	strXML = strXML & "<cboVelAcesso>"& rs("vel_id") &"</cboVelAcesso>"
	strXML = strXML & "<cboVelAcessoText>"& rs("velocidade_fisico") &"</cboVelAcessoText>"
	strXML = strXML & "<txtQtdeCircuitos>1</txtQtdeCircuitos>"
	strXML = strXML & "<cboTipoVelText>TIPO DE VELOCIDADE</cboTipoVelText>"
	strXML = strXML & "<cboProvedor>"& rs("pro_id") &"</cboProvedor>"
	strXML = strXML & "<cboProvedorText>"& rs("provedor") &"</cboProvedorText>"
	strXML = strXML & "<cboTipoPonto>I</cboTipoPonto>"
	strXML = strXML & "<cboTipoPontoText>CLIENTE</cboTipoPontoText>"
	strXML = strXML & "<cboUFEnd>"& rs("uf") &"</cboUFEnd>"
	strXML = strXML & "<cboUFEndText>"& rs("uf") &"</cboUFEndText>"
	strXML = strXML & "<txtEndCid>"& rs("cnl") &"</txtEndCid>"
	strXML = strXML & "<txtEndCidDesc>"& rs("cidade") &"</txtEndCidDesc>"
	strXML = strXML & "<cboLogrEnd>"& rs("tipo_de_logradouro") &"</cboLogrEnd>"
	strXML = strXML & "<txtEnd>"& rs("nome_do_logradouro") &"</txtEnd>"
	strXML = strXML & "<txtNroEnd>"& rs("numero_do_predio") &"</txtNroEnd>"
	strXML = strXML & "<txtBairroEnd>"& rs("bairro") &"</txtBairroEnd>"
	strXML = strXML & "<txtCepEnd>"& rs("cep") &"</txtCepEnd>"
	strXML = strXML & "<btnProcurarCepInstala>Procurar CEP</btnProcurarCepInstala>"
	strXML = strXML & "<txtContatoEnd>GERENTE OPERACIONAL</txtContatoEnd>"
	strXML = strXML & "<txtTelEndArea>"& mid(rs("telefone_contato"),1,2) &"</txtTelEndArea>"
	strXML = strXML & "<txtTelEnd>"& mid(rs("telefone_contato"),3,8) &"</txtTelEnd>"
	strXML = strXML & "<txtCNPJ>"& mid(rs("cnpj"),1,14) &"</txtCNPJ>"
	strXML = strXML & "<cboInterFaceEnd>"& rs("interface_usuario") &"</cboInterFaceEnd>"
	strXML = strXML & "<cboInterFaceEndText>"& rs("interface_embratel") &"</cboInterFaceEndText>"
	strXML = strXML & "<txtCNLSiglaCentroCliDest>"& mid(rs("estacao_entrega"),1,4) &"</txtCNLSiglaCentroCliDest>"'#SPO     #
	strXML = strXML & "<txtComplSiglaCentroCliDest>"& mid(rs("estacao_entrega"),5,3) &"</txtComplSiglaCentroCliDest>"'#IG        #
	strXML = strXML & "<txtEndEstacaoEntrega>"& rs("ESTACAO_ENDERECO") &"</txtEndEstacaoEntrega>"
	strXML = strXML & "<cboInterFaceEndFis>"& rs("interface_usuario") &"</cboInterFaceEndFis>"
	strXML = strXML & "<cboInterFaceEndFisText>"& rs("interface_embratel") &"</cboInterFaceEndFisText>"
	strXML = strXML & "<btnAddAcesso>Alterar</btnAddAcesso>"
	strXML = strXML & "<btnRemAcesso>Remover</btnRemAcesso>"
	strXML = strXML & "<btnIDFis1>Procurar Id Físico</btnIDFis1>"
	strXML = strXML & "<btnLimparAcesso>Limpar</btnLimparAcesso>"
	strXML = strXML & "<TipoAcao>N</TipoAcao>"
	strXML = strXML & "<intOrdem>1</intOrdem>"
	strXML = strXML & "<hdnIntIndice>0</hdnIntIndice>"
	strXML = strXML & "<hdnChaveAcessoFis>0</hdnChaveAcessoFis>"
	strXML = strXML & "<hdnVelAcessoFisSel>444</hdnVelAcessoFisSel>"'SEM VELOCIDADE
	strXML = strXML & "<hdnCNLCliente>"& rs("cnl") &"</hdnCNLCliente>"
	strXML = strXML & "</Acesso>"
	strXML = strXML & "<hdnCompartilhamento>0</hdnCompartilhamento>"
	strXML = strXML & "<hdnAcao>ResgatarUserCoordenacao</hdnAcao>"
	strXML = strXML & "<hdnCboServico>120</hdnCboServico>"
	strXML = strXML & "<hdnUserGICL>"& rs("username_gicl") &"</hdnUserGICL>"
	strXML = strXML & "<hdnDesigServ>"& rs("COD_DESIGNACAO_SERVICO") &"</hdnDesigServ>"
	strXML = strXML & "<hdnPossuiAvaliador>0</hdnPossuiAvaliador>"
	strXML = strXML & "<hdnDtSolicitacao>"&date()&"</hdnDtSolicitacao>"
	strXML = strXML & "<hdnPadraoDesignacao>L4*+/+IP+/+N5*</hdnPadraoDesignacao>"
	strXML = strXML & "<hdnCtfcId>2932</hdnCtfcId>"
	strXML = strXML & "<rdoProjEspecial>N</rdoProjEspecial>"
	strXML = strXML & "<cboOrigemSol>5</cboOrigemSol>"
	strXML = strXML & "<txtRazaoSocial>"& rs("razao_social") &"</txtRazaoSocial>"
	strXML = strXML & "<txtNomeFantasia>"& rs("nome_fantasia") &"</txtNomeFantasia>"
	strXML = strXML & "<txtContaSev>"& rs("cli_cc") &"</txtContaSev>"
	strXML = strXML & "<txtSubContaSev>"& rs("sub_conta") &"</txtSubContaSev>"
	strXML = strXML & "<cboServicoPedido>120</cboServicoPedido>"
	strXML = strXML & "<cboVelServico>"& rs("vel_id") &"</cboVelServico>"
	strXML = strXML & "<txtPadrao>00003</txtPadrao>"
	strXML = strXML & "<rdoNroContrato>3</rdoNroContrato>"
	strXML = strXML & "<txtNroContrServico>"& rs("NUMERO_CONTRATO_SERVICO") &"</txtNroContrServico>"
	strXML = strXML & "<txtDtEntrAcesServ>"&date()+10&"</txtDtEntrAcesServ>"
	strXML = strXML & "<txtDtPrevEntrAcesProv>"&date()+10&"</txtDtPrevEntrAcesProv>"
	strXML = strXML & "<txtDesigAcessoPri0>678</txtDesigAcessoPri0>"
	strXML = strXML & "<hdnServico>120</hdnServico>"
	strXML = strXML & "<hdntxtGICL>"& rs("username_gicl") &"</hdntxtGICL>"
	strXML = strXML & "<intTipoContrato>3</intTipoContrato>"
	strXML = strXML & "<hdnEstacaoAtual>"& rs("ESTACAO_ESC_ID") &"</hdnEstacaoAtual>"
	strXML = strXML & "<hdntxtGICN>"& rs("username_gicn") &"</hdntxtGICN>"
	strXML = strXML & "<hdnCoordenacaoAtual>txtGICN</hdnCoordenacaoAtual>"
	strXML = strXML & "<hdnNecessitaRecurso>S</hdnNecessitaRecurso>"
	strXML = strXML & "<hdnReaproveitarFisico>N</hdnReaproveitarFisico>"
	strXML = strXML & "<rdoEmiteOTS>N</rdoEmiteOTS>"
	strXML = strXML & "<cboLocalEntrega>"& rs("ESTACAO_ESC_ID") &"</cboLocalEntrega>"
	strXML = strXML & "<cboLocalConfig>"& rs("SIGLA_ESC_ID") &"</cboLocalConfig>"
	strXML = strXML & "<cboOrgao>1517</cboOrgao>"
	strXML = strXML & "<txtGICN>"& rs("username_gicn") &"</txtGICN>"
	strXML = strXML & "<hdnStatus>38</hdnStatus>"
	strXML = strXML & "<intTipoProc>1</intTipoProc>"
	strXML = strXML & "</xDados>"
	
	Set objXmlDadosXML = Server.CreateObject("Microsoft.XMLDOM")
	objXmlDadosXML.loadXml(strXml)

	strXmlDataBase = "<xDados/>"
	
	Vetor_Campos(1)="adlongvarchar," & len(strXml)& ",adParamInput," & strXml
	Vetor_Campos(2)="adlongvarchar," & len(strXmlDataBase)& ",adParamInput," & strXmlDataBase

	strSqlRet = APENDA_PARAMSTR("CLA_sp_ins_solicitacaoAtivacao",2,Vetor_Campos)
	
	'Retorno
	strXml =  ForXMLAutoQuery(strSqlRet)
	
	Set objXmlDadosForm = Server.CreateObject("Microsoft.XMLDOM")

	objXmlDadosForm.preserveWhiteSpace = True

	objXmlDadosForm.loadXml(strXml)
	Set objNode = objXmlDadosForm.selectNodes("//CLA_RetornoTmp[@Msg_Id=155]")
	if objNode.length > 0 then
		dblSolId = objNode(0).attributes(2).value
	else
	   SET erro = objXmlDadosForm.selectNodes("//CLA_Mensagem")
	   msg_erro = erro(0).attributes(0).value
	End if
	
	if msg_erro <> "" then
	  'gravar erro
	  strSQL = "Update cla_carga_solic set erro = '"& msg_erro &"' where solic_id = " & rs("solic_Id")
	else
	  'gravar solicitacao
	  strSQL = "Update cla_carga_solic set sol_id= '"& dblsolid &"',gravado = 1,erro = null,data_gravacao = getdate() where solic_id = " & rs("solic_Id")
	end if
	
	db.execute(strSQL)
    
	response.write "<h2>"& strSQL &"</h2>"
	
	cont = cont+1
	rs.MoveNext 
Loop

response.write "<center><h1><font color=blue><b>Total executado:</b> "&cont&" registro(s)</font></h1></center>"
%>