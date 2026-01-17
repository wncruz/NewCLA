<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: header.ASP
'	- Descrição			: Arquivo com o menu do CLA

Dim objDicCef
Dim objDicPag
Dim strGeral
Dim blnAcessoLog
Dim blnAcessoFis
Dim intTotalCelula
Dim strAcesspFis
Dim intPos
Dim intMenuOrd
Dim blnTabelas
Dim intTotalMenu
Dim objAryPag
Dim objAryPagAux
Dim strPagina

Set objRS = db.execute("CLA_sp_sel_usuarioctfc null," & dblUsuId)
Set objDicCef = Server.CreateObject("Scripting.Dictionary")
Set objDicPag = Server.CreateObject("Scripting.Dictionary")

While Not objRS.Eof
	if Not objDicCef.Exists(Ucase(Trim(objRS("Age_Sigla")))) then
		Call objDicCef.Add (Ucase(Trim(objRS("Age_Sigla"))),Ucase(Trim(objRS("Ctfc_ID"))))
	End if
	objRS.MoveNext
Wend

strGeral = ""
blnAcessoLog = false
blnAcessoFis = false

intTotalCelula = 0
if objDicCef.Exists("E") or objDicCef.Exists("GE") or objDicCef.Exists("GAT") then
	if objDicCef.Exists("E") or objDicCef.Exists("GE")then
		strGeral = strGeral & "var aMenu1 = new Array('','5','95','','','','','','','Solicitação de Acesso - Ativação','Solicitacao.asp',0,'Solicitação de Acesso - Alteração','AcessoLogico.asp?acao=alt',0,'Solicitação de Acesso - Cancelamento','AcessoLogico.asp?acao=can',0,'Solicitação de Acesso - Desativação','AcessoLogico.asp?acao=des',0,'Modificações de Informações do Acesso','AlteracaoMain.asp',0);" & Chr(13)
		Call objDicPag.Add("ALTERACAOMAIN.ASP","ALTERACAOMAIN.ASP")
		Call objDicPag.Add("SOLICITACAO.ASP","SOLICITACAO.ASP")
		Call objDicPag.Add("ACESSOLOGICO.ASP","ACESSOLOGICO.ASP")
	Else
		strGeral = strGeral & "var aMenu1 = new Array('','5','95','','','','','','','Modificações de Informações do Acesso','AlteracaoMain.asp',0);" & Chr(13)
		Call objDicPag.Add("ALTERACAOMAIN.ASP","ALTERACAOMAIN.ASP")
	End if	
	blnAcessoLog = true
	intTotalCelula = intTotalCelula + 1
End if	

if blnAcessoLog then
	strAcesspFis = strAcesspFis & "var aMenu2 = new Array('','180','95','','','','','','',"
Else
	strAcesspFis = strAcesspFis & "var aMenu1 = new Array('','5','95','','','','','','',"
End if	

'Aceite
if objDicCef.Exists("ATE") or objDicCef.Exists("MD") or objDicCef.Exists("MT") or objDicCef.Exists("GE") then
	strAcesspFis = strAcesspFis & "'Aceite','aceitacao_main.asp',0,"

	Call objDicPag.Add("ACEITACAO_MAIN.ASP","ACEITACAO_MAIN.ASP")
	Call objDicPag.Add("ACEITACAO.ASP","ACEITACAO.ASP")

	blnAcessoFis = true
	intTotalCelula = intTotalCelula + 1
End if

'Agendamento de Alteração
if objDicCef.Exists("E") or objDicCef.Exists("GE") then
	strAcesspFis = strAcesspFis & "'Agendamento de Alteração do Serviço/Acesso','javascript:document.cookie = ""COOK_LOCATION = CONSULTA"";window.location.replace(""Monitoracao.asp"");',0,"
	blnAcessoFis = true
End if

if objDicCef.Exists("MD") or objDicCef.Exists("MT") or objDicCef.Exists("GAT") or objDicCef.Exists("GE") then
	strAcesspFis = strAcesspFis & "'Alteração na Alocação de Facilidade(Manobra)','Manobra.asp',0,"
	Call objDicPag.Add("MANOBRA.ASP","MANOBRA.ASP")
	blnAcessoFis = true
End if


'Execução / Pendência de Manobra e Pendência de Retirada de Interligação
if objDicCef.Exists("MD") or objDicCef.Exists("MT") or objDicCef.Exists("AGE")or objDicCef.Exists("GE") then
	strAcesspFis = strAcesspFis & "'Execução','execucao_main.asp',0,"
End if

if objDicCef.Exists("GAT") or objDicCef.Exists("GE") then
	strAcesspFis = strAcesspFis & "'Liberação de Estoque','DesativacaoLote.asp',0,"
	Call objDicPag.Add("DESATIVACAOLOTE.ASP","DESATIVACAOLOTE.ASP")
End if

if objDicCef.Exists("MD") or objDicCef.Exists("MT") or objDicCef.Exists("AGE")or objDicCef.Exists("GE") then

	strAcesspFis = strAcesspFis & "'Pendência de Manobra','PENDENCIAMANOBRA.ASP',0,"
	strAcesspFis = strAcesspFis & "'Pendência de Retirada de Interligação','PedenciaInterligacaoRet.asp',0,"
	Call objDicPag.Add("PEDENCIAINTERLIGACAORET.ASP","PEDENCIAINTERLIGACAORET.ASP")
	Call objDicPag.Add("EXECUCAO_MAIN.ASP","EXECUCAO_MAIN.ASP")
	Call objDicPag.Add("EXECUCAO.ASP","EXECUCAO.ASP")
	Call objDicPag.Add("PENDENCIAMANOBRA.ASP","PENDENCIAMANOBRA.ASP")
	call objdicpag.add("CONFIRMACAOMANOBRA.ASP","CONFIRMACAOMANOBRA.ASP")

	blnAcessoFis = true
	intTotalCelula = intTotalCelula + 1
End if	

'Pedido
if objDicCef.Exists("GAT") or objDicCef.Exists("GAE") or objDicCef.Exists("GE") then
	strAcesspFis = strAcesspFis & "'Alocação de Facilidades','facilidade_main.asp',0,"
	Call objDicPag.Add("FACILIDADE_MAIN.ASP","FACILIDADE_MAIN.ASP")
	Call objDicPag.Add("FACILIDADE.ASP","FACILIDADE.ASP")
	Call objDicPag.Add("EMAILPROVEDOR.ASP","EMAILPROVEDOR.ASP")
	Call objDicPag.Add("EMAILATT.ASP","EMAILATT.ASP")
	blnAcessoFis = true
	intTotalCelula = intTotalCelula + 1
End if

											
if blnAcessoFis then
	strGeral = strGeral &  Left(strAcesspFis,len(strAcesspFis)-1) & ");" & Chr(13)
End if	

if not blnAcessoLog and not blnAcessoFis then
	intPos = 5
	intMenuOrd = 1
Else
	if (not blnAcessoLog and blnAcessoFis) or (blnAcessoLog and not blnAcessoFis) then
		intPos = 180
		intMenuOrd = 2
	Else
		intPos = 355	
		intMenuOrd = 3
	End if
End if

on error resume next 
if objDicCef.Exists("GAT") then
	'Consulta
	strGeral = strGeral & "var aMenu" & intMenuOrd & " = new Array('','" & intPos & "','95','','','','','','','Acessos Aceitos','ConsAcessosAceitos.asp',0,'Acessos Ativos (Lista)','consAtivosLista.asp',0,'Acessos Ativos por Velocidade','consAtivosResumoVel.asp',0,'Acessos Físicos','ConsultaAcessoFisico.asp',0,'Acessos Instalados (Lista)','consInstaladosLista.asp',0,'Acessos Instalados (Resumo)','consInstaladosResumo.asp',0,'Acessos Instalados (Sem Expurgo)','consInstaladosResumoExp.asp',0,'Acompanhamento','consacompanhamento.asp',0,'Controle de Acessos de Serviços Ativados','ConsControleAcesso.asp',0,'Controle de Rede Determinística','ConsRedeDet.asp',0,'Cartas ao Provedor','CartasProvedor.asp',0,'Desempenho do GLA - A. Aceitos','cons_GLA.asp',0,'Facilidades','ConsultarFacilidadeRec.asp',0,'Geral','consulta_main.asp',0,'Histórico de Pedido','historico_main.asp',0,'Histórico de Facilidades','cons_historicofacilidade.asp',0,'Monitoração','javascript:document.cookie = ""COOK_LOCATION = MONITORACAO"";window.location.replace(""Monitoracao.asp"");',0," & _
		"'Mais...','',17);" & _
		" var aMenu" & intMenuOrd & "_18 = new Array('Movimentação de Acessos (Lista)','consMovimentacao.asp',0,'Movimentação de Acessos Temporários (Lista)','consMovimentacaoTemp.asp',0,'Pendência de Cabo Interno','ConsPendCaboInterno.asp',0,'Pedidos por Ação','cons_Aca.asp',0,'Pedidos por Provedor e Ação','cons_ProAca.asp',0,'Pedidos por Provedor e Status','cons_ProStp.asp',0,'Pendentes de Instalação (Lista)','consPendInstalaLista.asp',0,'Pendentes de Instalação (Resumo)','consPendInstalaResumo.asp',0,'Percentual de Acessos com Serviços Ativados','consPercentAcesso.asp',0,'Preço-Provedor','javascript:window.open(""../Excel/PrecoProvedor/PrecoProvedor.xls"");',0,'Relatório de Backlogs Provedor','relapendconstacesso.asp',0,'Relatório Telco','RelaTelco.asp',0,'Qtde. de Pedidos por Provedor e Ação','cons_ProAcaQtd.asp',0,'Qtde. de Pedidos Aceitos por Provedor','cons_ProQtd.asp',0,'Qtde. de Pedidos Pendentes por Provedor','cons_ProData.asp',0,'Velocidade por Serviço','consVelServico.asp',0);" & Chr(13)
Else
	strGeral = strGeral & "var aMenu" & intMenuOrd & " = new Array('','" & intPos & "','95','','','','','','','Acessos Aceitos','ConsAcessosAceitos.asp',0,'Acessos Ativos (Lista)','consAtivosLista.asp',0,'Acessos Ativos por Velocidade','consAtivosResumoVel.asp',0,'Acessos Físicos','ConsultaAcessoFisico.asp',0,'Acessos Instalados (Lista)','consInstaladosLista.asp',0,'Acessos Instalados (Resumo)','consInstaladosResumo.asp',0,'Acessos Instalados(Sem Expurgo)','consInstaladosResumoExp.asp',0,'Acompanhamento','consacompanhamento.asp',0,'Controle de Acessos de Serviços Ativados','ConsControleAcesso.asp',0,'Controle de Rede Determinística','ConsRedeDet.asp',0,'Desempenho do GLA - A. Aceitos','cons_GLA.asp',0,'Facilidades','ConsultarFacilidadeRec.asp',0,'Geral','consulta_main.asp',0,'Histórico de Pedido','historico_main.asp',0,'Histórico de Facilidades','cons_historicofacilidade.asp',0,'Monitoração','javascript:document.cookie = ""COOK_LOCATION = MONITORACAO"";window.location.replace(""Monitoracao.asp"");',0,'Pendência de Cabo Interno','ConsPendCaboInterno.asp',0," & _
	"'Mais...','',17);" & _
	" var aMenu"& intMenuOrd & "_18 = new Array('Movimentação de Acessos (Lista)','consMovimentacao.asp',0,'Movimentação de Acessos Temporários (Lista)','consMovimentacaoTemp.asp',0,'Pedidos por Ação','cons_Aca.asp',0,'Pedidos por Provedor e Ação','cons_ProAca.asp',0,'Pedidos por Provedor e Status','cons_ProStp.asp',0,'Pendentes de Instalação (Lista)','consPendInstalaLista.asp',0,'Pendentes de Instalação (Resumo)','consPendInstalaResumo.asp',0,'Percentual de Acessos com Serviços Ativados','consPercentAcesso.asp',0,'Preço-Provedor','javascript:window.open(""../Excel/PrecoProvedor/PrecoProvedor.xls"");',0,'Relatório de Backlogs Provedor','relapendconstacesso.asp',0,'Relatório Telco','RelaTelco.asp',0,'Qtde. de Pedidos por Provedor e Ação','cons_ProAcaQtd.asp',0,'Qtde. de Pedidos Aceitos por Provedor','cons_ProQtd.asp',0,'Qtde. de Pedidos Pendentes por Provedor','cons_ProData.asp',0,'Velocidade por Serviço','consVelServico.asp',0);" & Chr(13)
End if
blnTabelas = false
intTotalMenu = 0


if objDicCef.Exists("PST") or objDicCef.Exists("GE") or objDicCef.Exists("GAT") or objDicCef.Exists("MD") or objDicCef.Exists("MT") or objDicCef.Exists("E") or objDicCef.Exists("GAE") or objDicCef.Exists("ML") or objDicCef.Exists("SMB") then
	strGeral = strGeral & "var aMenu" & intMenuOrd+1 & " = new Array('','" & intPos+190 & "','95','','','','','',''," 
	blnTabelas = true
End if	

if objDicCef.Exists("SMB") then
	strGeral = strGeral & _
				"'SMBD','smbd_main.asp',0," 

	Call objdicpag.add("SMBD_MAIN.ASP","SMBD_MAIN.ASP")
	intTotalMenu = intTotalMenu + 1
End if


if objDicCef.Exists("GE") then
	strGeral = strGeral & _
				"'Agente','agente_main.asp',0," & _
				"'Agente do Pedido','agentepedido_main.asp',0,"

	Call objdicpag.add("AGENTE_MAIN.ASP","AGENTE_MAIN.ASP")
	Call objdicpag.add("AGENTE.ASP","AGENTE.ASP")
	Call objDicPag.Add("AGENTEPEDIDO_MAIN.ASP","AGENTEPEDIDO_MAIN.ASP")
	Call objDicPag.Add("AGENTEPEDIDO.ASP","AGENTEPEDIDO.ASP")

	intTotalMenu = intTotalMenu + 2

End if

if objDicCef.Exists("GE") then
	strGeral = strGeral & _
				"'Associação Usuário com CF','AssocUserCef_main.asp',0," & _
				"'Associação de Serviço com Velocidade', 'AssocServVel_main.asp',0," & _
				"'Associação de Tecnologia com Velocidade', 'AssocTecVel_main.asp',0," & _
				"'Atualizar Base do Correio', 'BaseCorreio.asp',0," & _
				"'Atualizar Carta Provedor (Upload)', 'UploadCartaProvedor.asp',0," & _
				"'Centro Funcional','CentroFuncional_main.asp',0," & _
				"'Cidade','cidade_main.asp',0," & _
				"'Distribuidor','distribuicao_main.asp',0," & _
				"'Estação','estacao_main.asp',0,"

	Call objDicPag.Add("ACAO.ASP","ACAO.ASP")
	Call objdicpag.add("ASSOCUSERCEF_MAIN.ASP","ASSOCUSERCEF_MAIN.ASP")
	Call objdicpag.add("ASSOCUSERCEF.ASP","ASSOCUSERCEF.ASP")
	Call objdicpag.add("ASSOCSERVVEL_MAIN.ASP","ASSOCSERVVEL_MAIN.ASP")
	Call objdicpag.add("ASSOCSERVVEL.ASP","ASSOCSERVVEL.ASP")
	Call objdicpag.add("ASSOCTECVEL_MAIN.ASP","ASSOCTECVEL_MAIN.ASP")
	Call objdicpag.add("ASSOCTECVEL.ASP","ASSOCTECVEL.ASP")
	Call objdicpag.add("BASECORREIO.ASP","BASECORREIO.ASP")
	Call objdicpag.add("UPLOADCARTAPROVEDOR.ASP","UPLOADCARTAPROVEDOR.ASP")
	Call objdicpag.add("CENTROFUNCIONAL_MAIN.ASP","CENTROFUNCIONAL_MAIN.ASP")
	Call objdicpag.add("CENTROFUNCIONAL.ASP","CENTROFUNCIONAL.ASP")
	Call objdicpag.add("CIDADE_MAIN.ASP","CIDADE_MAIN.ASP")
	Call objdicpag.add("CIDADE.ASP","CIDADE.ASP")
	Call objdicpag.add("DISTRIBUICAO_MAIN.ASP","DISTRIBUICAO_MAIN.ASP")
	Call objdicpag.add("DISTRIBUICAO.ASP","DISTRIBUICAO.ASP")
	Call objdicpag.add("ESTACAO_MAIN.ASP","ESTACAO_MAIN.ASP")
	Call objdicpag.add("ESTACAO.ASP","ESTACAO.ASP")

	intTotalMenu = intTotalMenu + 9

End if

if objDicCef.Exists("GE") or objDicCef.Exists("GAT") or objDicCef.Exists("GAE") then
	strGeral = strGeral & "'Facilidade','CadastrarFacilidade.asp',0,"
	Call objdicpag.add("CADASTRARFACILIDADE.ASP","CADASTRARFACILIDADE.ASP")
	intTotalMenu = intTotalMenu + 1
End if

if objDicCef.Exists("PST") then
	strGeral = strGeral & "'Holding','Holding_main.asp',0,"
	Call objdicpag.add("HOLDING_MAIN.ASP","HOLDING_MAIN.ASP")
	Call objdicpag.add("HOLDING.ASP","HOLDING.ASP")
	intTotalMenu = intTotalMenu + 1
End if

if objDicCef.Exists("GE") or objDicCef.Exists("GAT") then
	strGeral = strGeral & "'Logradouro','logradouro_main.asp',0," 

	Call objdicpag.add("LOGRADOURO_MAIN.ASP","LOGRADOURO_MAIN.ASP")
	Call objdicpag.add("LOGRADOURO.ASP","LOGRADOURO.ASP")

	intTotalMenu = intTotalMenu + 1
End if

if objDicCef.Exists("GE") then
	strGeral = strGeral & "'Localidade Subordinada ao CF','LocalidadeSubCef_main.asp',0,"

	Call objdicpag.add("LOCALIDADESUBCEF_MAIN.ASP","LOCALIDADESUBCEF_MAIN.ASP")
	Call objdicpag.add("LOCALIDADESUBCEF.ASP","LOCALIDADESUBCEF.ASP")
	intTotalMenu = intTotalMenu + 1
End if

if objDicCef.Exists("GE") or objDicCef.Exists("GAT") or objDicCef.Exists("GAE") or objDicCef.Exists("MD") or objDicCef.Exists("MT") then
	intTotalMenu = intTotalMenu + 1
	strGeral = strGeral & "'Manutenção de Facilidade','ManutencaoFacilidade.asp',0,"
	Call objdicpag.add("MANUTENCAOFACILIDADE.ASP","MANUTENCAOFACILIDADE.ASP")
End if		

if objDicCef.Exists("GE") then

	intTotalMenu = intTotalMenu + 2
	strGeral = strGeral & "'Órgão','orgao_main.asp',0," & _
						"'Mais...',''," & intTotalMenu & ");" & _
						" var aMenu" & intMenuOrd+1 & "_" & intTotalMenu & " = new Array("

	strGeral = strGeral & "'Parâmetros CF','CefConfig_main.asp',0,"
	strGeral = strGeral & "'Plataforma','Plataforma_main.asp',0,"

	Call objdicpag.add("ORGAO_MAIN.ASP","ORGAO_MAIN.ASP")
	Call objdicpag.add("ORGAO.ASP","ORGAO.ASP")
	call objdicpag.add("CEFCONFIG_MAIN.ASP","CEFCONFIG_MAIN.ASP")
	call objdicpag.add("CEFCONFIG.ASP","CEFCONFIG.ASP")
	call objdicpag.add("PLATAFORMA_MAIN.ASP","CEFCONFIG_MAIN.ASP")
	call objdicpag.add("PLATAFORMA.ASP","CEFCONFIG.ASP")
End if


if objDicCef.Exists("PST") or objDicCef.Exists("GE")then

	strGeral = strGeral & "'Promoção','Promocao_main.asp',0,"
	strGeral = strGeral & "'Provedor','provedor_main.asp',0,"
	strGeral = strGeral & "'Provedor Contrato','provedorcontrato_main.asp',0,"

	call objdicpag.add("PROMOCAO_MAIN.ASP","PROMOCAO_MAIN.ASP")
	call objdicpag.add("PROMOCAO.ASP","PROMOCAO.ASP")
	call objdicpag.add("PROVEDOR_MAIN.ASP","PROVEDOR_MAIN.ASP")
	call objdicpag.add("PROVEDOR.ASP","PROVEDOR.ASP")
	call objdicpag.add("PROVEDORCONTRATO_MAIN.ASP" ,"PROVEDORCONTRATO_MAIN.ASP")
	call objdicpag.add("PROVEDORCONTRATO.ASP" ,"PROVEDORCONTRATO.ASP")

End if

if objDicCef.Exists("PST") then
	strGeral = strGeral &	"'Preço-Provedor','postFile.asp',0,"
						  
	Call objDicPag.Add("POSTFILE.ASP","POSTFILE.ASP")
End if

if objDicCef.Exists("GE") or objDicCef.Exists("GAT") then
	strGeral = strGeral &	"'Recurso','recurso_main.asp',0," & _
							"'Redirecionar Solicitação','redirsolicitacao_main.asp',0,"
						  
	Call objdicpag.add("RECURSO_MAIN.ASP","RECURSO_MAIN.ASP")
	Call objdicpag.add("RECURSO.ASP","RECURSO.ASP")
	Call objdicpag.add("REDIRSOLICITACAO_MAIN.ASP","REDIRSOLICITACAO_MAIN.ASP")
	Call objdicpag.add("REDIRSOLICITACAO.ASP","REDIRSOLICITACAO.ASP")
End if

if objDicCef.Exists("PST") then
	strGeral = strGeral & "'Regime de Contrato','regimecontrato_main.asp',0,"
	call objdicpag.add("REGIMECONTRATO_MAIN.ASP","REGIMECONTRATO_MAIN.ASP")
	call objdicpag.add("REGIMECONTRATO.ASP","REGIMECONTRATO.ASP")
End if	

if objDicCef.Exists("GE") or objDicCef.Exists("MD") or objDicCef.Exists("MT") or objDicCef.Exists("ML") or objDicCef.Exists("GAE") then
	strGeral = strGeral & "'Rede Interna','redeInterna.asp',0,"
	Call objDicPag.Add("REDEINTERNA.ASP","REDEINTERNA.ASP")
End if

if objDicCef.Exists("GE") then

	strGeral = strGeral & "'Serviço','servico_main.asp',0," & _
						"'Status','status_main.asp',0,"

	call objdicpag.add("SERVICO_MAIN.ASP","SERVICO_MAIN.ASP")
	call objdicpag.add("SERVICO.ASP","SERVICO.ASP")
	call objdicpag.add("STATUS_MAIN.ASP","STATUS_MAIN.ASP")
	call objdicpag.add("STATUS.ASP","STATUS.ASP")

End if

if objDicCef.Exists("GE") or objDicCef.Exists("PST") then
	strGeral = strGeral & "'Tipo de Vigência de Contrato','tipocontrato_main.asp',0,"
	call objdicpag.add("TIPOCONTRATO_MAIN.ASP","TIPOCONTRATO_MAIN.ASP")
	call objdicpag.add("TIPOCONTRATO.ASP","TIPOCONTRATO.ASP")
End if

if objDicCef.Exists("GE") or objDicCef.Exists("GAE") then
	strGeral = strGeral & "'Tipo de Rádio','Cadtiporadio_main.asp',0,"
	call objdicpag.add("CADTIPORADIO_MAIN.ASP","CADTIPORADIO_MAIN.ASP")
	call objdicpag.add("CADTIPORADIO.ASP","CADTIPORADIO.ASP")
End if


if objDicCef.Exists("GE") or objDicCef.Exists("GAT") or objDicCef.Exists("E") then

	strGeral = strGeral & "'Tranferência de Carteira','TransferenciaCarteira.asp',0,"
	call objdicpag.add("TRANSFERENCIACARTEIRA.ASP","TRANSFERENCIACARTEIRA.ASP")

End if

if objDicCef.Exists("GE") then

	strGeral = strGeral & "'Usuário','usuario_main.asp',0," & _
							"'Velocidade','velocidade_main.asp',0,"

	call objdicpag.add("USUARIO_MAIN.ASP","USUARIO_MAIN.ASP")
	call objdicpag.add("USUARIO.ASP","USUARIO.ASP")
	call objdicpag.add("VELOCIDADE_MAIN.ASP","VELOCIDADE_MAIN.ASP")
	call objdicpag.add("VELOCIDADE.ASP","VELOCIDADE.ASP")

End if	

if blnTabelas then
	strGeral = Left(strGeral,len(strGeral)-1) & ");" & Chr(13)
End if	

Call objDicPag.Add("CONSPENDCABOINTERNO.ASP","CONSPENDCABOINTERNO.ASP")
Call objDicPag.Add("CONSACESSOSACEITOS.ASP","CONSACESSOSACEITOS.ASP")
Call objDicPag.Add("CONSULTA_MAIN.ASP","CONSULTA_MAIN.ASP")
Call objDicPag.Add("CONSULTAACESSOFISICO.ASP","CONSULTAACESSOFISICO.ASP")
if objDicCef.Exists("GAT") then
	Call objDicPag.Add("CARTASPROVEDOR.ASP","CARTASPROVEDOR.AS")
End if
Call objDicPag.Add("CONSACOMPANHAMENTO.ASP","CONSACOMPANHAMENTO.ASP")
Call objDicPag.Add("MONITORACAO.ASP","MONITORACAO.ASP")
Call objDicPag.Add("CONS_ACA.ASP","CONS_ACA.ASP")
Call objDicPag.Add("CONS_PROACA.ASP","CONS_PROACA.ASP")
Call objDicPag.Add("CONS_PROSTP.ASP","CONS_PROSTP.ASP")
Call objDicPag.Add("CONS_GLA.ASP","CONS_GLA.ASP")
Call objDicPag.Add("CONS_PROACAQTD.ASP","CONS_PROACAQTD.ASP")
Call objDicPag.Add("CONS_PROQTD.ASP","CONS_PROQTD.ASP")
Call objDicPag.Add("CONS_PRODATA.ASP","CONS_PRODATA.ASP")
Call objDicPag.Add("CONSVELSERVICO.ASP","CONSVELSERVICO.ASP")
Call objDicPag.Add("HISTORICO_MAIN.ASP","HISTORICO_MAIN.ASP")
Call objDicPag.Add("CONS_HISTORICOFACILIDADE.ASP","CONS_HISTORICOFACILIDADE.ASP")
Call objDicPag.Add("HISTORICO.ASP","HISTORICO.ASP")
Call objDicPag.Add("CONSULTARFACILIDADEREC.ASP","CONSULTARFACILIDADEREC.ASP")

call objdicpag.add("FACILIDADEDET.ASP","FACILIDADEDET.ASP")
call objdicpag.add("ALTERACAOCAD.ASP","ALTERACAOCAD.ASP")
call objdicpag.add("CONSREDEDET.ASP","CONSREDEDET.ASP")
call objdicpag.add("CONSCONTROLEACESSO.ASP","CONSCONTROLEACESSO.ASP")
call objdicpag.add("MAIN.ASP","MAIN.ASP")
Call objDicPag.Add("ACOMPANHAMENTO.ASP","ACOMPANHAMENTO.ASP")
Call objDicPag.Add("CONSULTAGERALIMP.ASP","CONSULTAGERALIMP.ASP")

Call objDicPag.Add("CONSPENDINSTALALISTA.ASP","CONSPENDINSTALALISTA.ASP")
Call objDicPag.Add("CONSPENDINSTALARESUMO.ASP","CONSPENDINSTALARESUMO.ASP")
Call objDicPag.Add("CONSPERCENTACESSO.ASP","CONSPERCENTACESSO.ASP")
Call objDicPag.Add("CONSINSTALADOSLISTA.ASP","CONSINSTALADOSLISTA.ASP")
Call objDicPag.Add("CONSINSTALADOSRESUMO.ASP","CONSINSTALADOSRESUMO.ASP")
Call objDicPag.Add("CONSINSTALADOSRESUMOEXP.ASP","CONSINSTALADOSRESUMOEXP.ASP")
Call objDicPag.Add("CONSATIVOSLISTA.ASP","CONSATIVOSLISTA.ASP")
Call objDicPag.Add("CONSATIVOSRESUMOVEL.ASP","CONSATIVOSRESUMOVEL.ASP")

Call objDicPag.Add("RELAPENDCONSTACESSO.ASP","RELAPENDCONSTACESSO.ASP")
Call objDicPag.Add("RELATELCO.ASP","RELATELCO.ASP")

Call objDicPag.Add("CONSMOVIMENTACAO.ASP","CONSMOVIMENTACAO.ASP")
Call objDicPag.Add("CONSMOVIMENTACAOTEMP.ASP","CONSMOVIMENTACAOTEMP.ASP")

Call objDicPag.Add("FILEPOST.ASP","FILEPOST.ASP")

objAryPag = split(Request.ServerVariables("SCRIPT_NAME"),"/")
objAryPagAux = split(Ucase(objAryPag(Ubound(objAryPag))),"?")
strPagina = objAryPagAux(0)

if Not objDicPag.Exists(Trim(Ucase(strPagina))) and Not objDicPag.Exists(Trim(Ucase("SMBD_MAIN.ASP"))) then
		Response.Write "<script language=javascript >alert('USUÁRIO SEM PERMISSÃO!');window.location.reload('../index.asp');</script>"
		Response.End 
End if



Set objDicPag = Nothing
%>
<Html>
<Head>
<Title>CLA - Controle Local de Acesso</Title>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='JavaScript' src='../javascript/formatamenu.js'></script>
<script language='JavaScript' src='../javascript/montamenu.js'></script>
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>
<Script language='javascript'>
if (top.location != self.location)
{
	top.location = self.location;
}
try{
<%=strGeral%>
}
catch(e){
	alert(e.description)
}
</Script>
<SCRIPT language='javascript'>
  javascript:window.history.forward(1);
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" onLoad="resposta(<%if DBAction <> "" then response.write DBAction else response.write "0" end if%>,'main.asp');">
<table width="760" border="0" cellspacing="0" cellpadding="0">
<tr bgcolor="#f1f1f1"> 
	<td valign=top>
		<a href="main.asp"><img name="embratel" src="../imagens/topo_embratel.jpg" width=760px height=80px border="0"></a>
		<div style="position:absolute;left:550;top:0;">
			<table border=0 cellspacing=0 cellpadding="0">
				<tr>
					<td width=60><font color=white size=1>Servidor</font></td>
					<td><font color=white size=1><%=Request.ServerVariables("SERVER_NAME") %></font></td>
				</tr>	
				<tr>
					<td><font color=white size=1>Banco</font></td>
					<td><font color=white size=1><%=strBanco%></font></td>
				</tr>	
				<tr>
					<td><font color=white size=1>Usuário</font></td>
					<td><font color=white size=1><%=strUserName%></font></td>
				</tr>	
			</table>
		</div>
	</td>	
</tr>
<tr>
	<td height=12px bgcolor="#f1f1f1" valign=top>
		<div id=divMenu style="position:absolute;top:81;left:5">
			<table width="<%=190*(intTotalCelula)%>" border="0" cellspacing="0" cellpadding="0" >
				<tr>
					<%if blnAcessoLog then%>
					<td nowrap width=175px >
						<span id=spnAcessoLog onMouseOut="popDown('eMenu1')" onClick="showInput(false);popUp('eMenu1',event)" style='cursor:hand' ><font class=clsMenu >Acesso Lógico</font></span>
					</td>
					<%End if%>
					<%if blnAcessoFis and blnAcessoLog then%>
					<td nowrap width=175px >
						<span id=spnAcessoFis  onMouseOut="popDown('eMenu2')" onClick="showInput(false);popUp('eMenu2',event)" style='cursor:hand' ><font class=clsMenu >|&nbsp;Acesso Físico</font></span>
					</td>
					<%Else%>
						<%if blnAcessoFis and not blnAcessoLog then%>
						<td nowrap width=175px >
							<span id=spnAcessoFis onMouseOut="popDown('eMenu1')" onClick="showInput(false);popUp('eMenu1',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Acesso Físico</font></span>
						</td>
						<%End if%>
					<%End if%>
					<td nowrap width=175px >
						<span id=spnConsultas onMouseOut="popDown('eMenu<%=intMenuOrd%>')" onClick="showInput(false);popUp('eMenu<%=intMenuOrd%>',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Consultas</font></span>
					</td>
					<%if blnTabelas then %>
					<td nowrap width=175px >
						<span id=spnTabelas onMouseOut="popDown('eMenu<%=intMenuOrd+1%>')" onClick="showInput(false);popUp('eMenu<%=intMenuOrd+1%>',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Tabelas</font></span>
					</td>
					<%End if%>
					<td nowrap >
						<span id=spnCrms style='cursor:hand' onClick="window.open('http://<%=Request.ServerVariables("SERVER_NAME")%>/crmsf/inicio.htm','CRMSF','')" onmouseover="showtip(this,event,'Controle de Radio,Mux,Satélite e Fibra');" onmouseout="hidetip();"><font class=clsMenu >|&nbsp;CRMSF</font></span>
					</td>
				</tr>
			</table>
		</div>
	</td>
</tr>
</table>
<table valign="top" width="780" cellspacing="2" cellpadding="0">
<input type=hidden name=hdnUserHerder value="<%=strUserName%>">