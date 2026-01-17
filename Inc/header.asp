<%
                Response.ContentType = "text/html; charset=utf-8"
                Response.Charset = "UTF-8"
              %>
			  
<!--include file="data.asp"-->
<!--#include file="../inc/log.asp"-->
<%
on error resume next
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
Dim intPos
Dim intMenuOrd
Dim blnTabelas
Dim intTotalMenu
Dim objAryPag
Dim objAryPagAux
Dim strPagina
Dim strAcessoFis
Dim strGerencia
Dim strConsulta
Dim strTabela
dim strUserHerderPerfil

IF strLoginRede = "EDAR" or strLoginRede = "edar" or strLoginRede = "T3FRRP"  then
	'dblUsuId      = 869
	'strLoginRede  = "ducret"

	'dblUsuId      = 1076
	'strLoginRede  = "rccard"

	'dblUsuId      = 4098
	'strLoginRede  = "scesar"


END IF

Set objRS = db.execute("CLA_sp_sel_usuarioctfc2 null," & dblUsuId)
Set objDicCef = Server.CreateObject("Scripting.Dictionary")
Set objDicPag = Server.CreateObject("Scripting.Dictionary")

While Not objRS.Eof
	if Not objDicCef.Exists(Ucase(Trim(objRS("Age_Sigla")))) then
		Call objDicCef.Add (Ucase(Trim(objRS("Age_Sigla"))),Ucase(Trim(objRS("Ctfc_ID"))))
	End if
	objRS.MoveNext
Wend

'''@@@ Zera as Variaveis
strGeral 	 = ""
strAcessoFis = ""
strGerencia  = ""
strConsulta	 = ""
strTabela	 = ""
blnAcessoLog = false
blnGerencia = false
blnAcessoFis = false
blnTabelas 	 = false

intTotalCelula = 0

'PRSSILV - Autorização Acesso - 16/03/2009
Set objRSPerf = db.execute("CLA_sp_view_loginusuario '" & strLoginRede & "'")
If not objRSPerf.eof then
  var_Usu_PerfCadSenha = objRSPerf("Usu_PerfCadSenha")
  var_Usu_PerfAltDesig = objRSPerf("Usu_PerfAltDesig")
  
  'Reenvio Aprovisionador
  var_Usu_PerfReenvAprov = objRSPerf("Usu_PerfReenvAprov")
End if

''@@@@@@@@@@@@ ACESSO LOGICO

if objDicCef.Exists("E") or objDicCef.Exists("GE") or objDicCef.Exists("GAT") or objDicCef.Exists("AVL") then
	if objDicCef.Exists("E") or objDicCef.Exists("GE") then
		
		'PRSSILV/EDAR - Autorização Acesso - 14/08/2011
		''Call objDicPag.Add("PERFCADSENHA2M_MAIN.ASP","PERFCADSENHA2M_MAIN.ASP")
		''Call objDicPag.Add("INCLUIRAUTORIZARACESSO.ASP","INCLUIRAUTORIZARACESSO.ASP")
		''Call objDicPag.Add("CONSULTARAUTORIZARACESSO.ASP","CONSULTARAUTORIZARACESSO.ASP")
		''Call objDicPag.Add("ALTERARAUTORIZARACESSO.ASP","ALTERARAUTORIZARACESSO.ASP")
		'Call objDicPag.Add("CONSULTARAPROVARACESSO.ASP","CONSULTARAPROVARACESSO.ASP")

		'** OBS: Ao incluir um item no menu, incluir no IF e no ELSE **'
		
		'Reenvio Aprovisionador
'		if var_Usu_PerfReenvAprov = 1 then
		IF strLoginRede = "EDAR" or strLoginRede = "edar" or strLoginRede = "JCARTUS" or strLoginRede = "SCESAR" or strLoginRede = "T3FRRP"  or strLoginRede = "t3frrp" THEN
			''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
			strGeral = strGeral & "var aMenu1 = new Array('','5','95','','','','','','','Solicitação de Acesso - Aprovisionador','SolicitacaoAprov_main.asp',0,'Solicitação de Acesso - Ativação','Solicitacao.asp',0,'Solicitação de Acesso - Alteração','AcessoLogico.asp?acao=alt',0,'Solicitação de Acesso - Cancelamento','AcessoLogico.asp?acao=can',0,'Solicitação de Acesso - Desativação','AcessoLogico.asp?acao=des',0,'Solicitação de Reenvio - Aprovisionador','SolicitacaoReenvioAprov_main.asp',0,'Entregar/Ativar serviço','javascript:document.cookie = ""COOK_LOCATION = MONITORACAO"";window.location.replace(""monitoracaoEntregaAtivacao.asp"");',0);" & Chr(13)
			''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
			'strGeral = strGeral & "var aMenu1 = new Array('','5','95','','','','','','','Solicitação de Acesso - Aprovisionador','SolicitacaoAprov_main.asp',0,'Solicitação de Reenvio - Aprovisionador','SolicitacaoReenvioAprov_main.asp',0,'Solicitação de Acesso - Cancelamento','AcessoLogico.asp?acao=can',0,'Solicitação de Acesso - Desativação','AcessoLogico.asp?acao=des',0,'Entregar/Ativar serviço','javascript:document.cookie = ""COOK_LOCATION = MONITORACAO"";window.location.replace(""monitoracaoEntregaAtivacao.asp"");',0,'Habilitar Interface CLE','HabilitaCLE.asp',0,'Habilitar Interface APG','HabilitaAPG.asp',0);" & Chr(13)
			Call objDicPag.Add("HABILITACLE.ASP","HABILITACLE.ASP")
			Call objDicPag.Add("HABILITAAPG.ASP","HABILITAAPG.ASP")
		else
			''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
			strGeral = strGeral & "var aMenu1 = new Array('','5','95','','','','','','','Solicitação de Acesso - Aprovisionador','SolicitacaoAprov_main.asp',0,'Solicitação de Acesso - Ativação','Solicitacao.asp',0,'Solicitação de Acesso - Alteração','AcessoLogico.asp?acao=alt',0,'Solicitação de Acesso - Cancelamento','AcessoLogico.asp?acao=can',0,'Solicitação de Acesso - Desativação','AcessoLogico.asp?acao=des',0,'Solicitação de Reenvio - Aprovisionador','SolicitacaoReenvioAprov_main.asp',0,'Entregar/Ativar serviço','javascript:document.cookie = ""COOK_LOCATION = MONITORACAO"";window.location.replace(""monitoracaoEntregaAtivacao.asp"");',0);" & Chr(13)
			''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
			'strGeral = strGeral & "var aMenu1 = new Array('','5','95','','','','','','','Solicitação de Acesso - Aprovisionador','SolicitacaoAprov_main.asp',0,'Solicitação de Reenvio - Aprovisionador','SolicitacaoReenvioAprov_main.asp',0,'Solicitação de Acesso - Cancelamento','AcessoLogico.asp?acao=can',0,'Solicitação de Acesso - Desativação','AcessoLogico.asp?acao=des',0,'Entregar/Ativar serviço','javascript:document.cookie = ""COOK_LOCATION = MONITORACAO"";window.location.replace(""monitoracaoEntregaAtivacao.asp"");',0);" & Chr(13)
		end if
		
		Call objDicPag.Add("SOLICITACAOREENVIOAPROV_MAIN.ASP","SOLICITACAOREENVIOAPROV_MAIN.ASP")
'		else
'			strGeral = strGeral & "var aMenu1 = new Array('','5','95','','','','','','','Solicitação de Acesso - Aprov. (APG)','Solicitacao_main.asp',0,'Solicitação de Acesso - Aprovisionador','SolicitacaoAprov_main.asp',0,'Solicitação de Acesso - Ativação','Solicitacao.asp',0,'Solicitação de Acesso - Alteração','AcessoLogico.asp?acao=alt',0,'Solicitação de Acesso - Cancelamento','AcessoLogico.asp?acao=can',0,'Solicitação de Acesso - Desativação','AcessoLogico.asp?acao=des',0,'Solicitação de senha PIN','CONSULTARAUTORIZARACESSO.asp',0,'Modificações de Informações do Acesso','AlteracaoMain.asp',0,'Avaliação de Acesso','gerenciamento_main.ASP',0,'Entregar/Ativar serviço','javascript:document.cookie = ""COOK_LOCATION = MONITORACAO"";window.location.replace(""monitoracaoEntregaAtivacao.asp"");',0);" & Chr(13)
'		End if

		Call objDicPag.Add("SOLICITACAO_MAIN.ASP","SOLICITACAO_MAIN.ASP")
		Call objDicPag.Add("SOLICITACAOAPROV_MAIN.ASP","SOLICITACAOAPROV_MAIN.ASP")
		Call objDicPag.Add("ALTERACAOMAIN.ASP","ALTERACAOMAIN.ASP")
		Call objDicPag.Add("SOLICITACAO.ASP","SOLICITACAO.ASP")
		Call objDicPag.Add("ACESSOLOGICO.ASP","ACESSOLOGICO.ASP")
		Call objDicPag.Add("GERENCIAMENTO_MAIN.ASP","GERENCIAMENTO_MAIN.ASP")
	    Call objDicPag.Add("MONITORACAO.ASP","MONITORACAO.ASP")
		Call objDicPag.Add("MONITORACAOENTREGAATIVACAO.ASP","MONITORACAOENTREGAATIVACAO.ASP")
		Call objDicPag.Add("SOLICITACAOAPROV.ASP","SOLICITACAOAPROV.ASP")
		Call objDicPag.Add("ALTERACAOCADAPROV.ASP","ALTERACAOCADAPROV.ASP")

	''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
	'Elseif objDicCef.Exists("AVL") then
	'	  strGeral = strGeral & "var aMenu1 = new Array('','5','95','','','','','','','Avaliação de Acesso','gerenciamento_main.ASP',0,'Modificações de Informações do Acesso','AlteracaoMain.asp',0);" & Chr(13)
	'	  Call objDicPag.Add("GERENCIAMENTO_MAIN.ASP","GERENCIAMENTO_MAIN.ASP")
	''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
	else
		''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
		'strGeral = strGeral & "var aMenu1 = new Array('','5','95','','','','','','','Modificações de Informações do Acesso','AlteracaoMain.asp',0);" & Chr(13)
		''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
		strGeral = strGeral & "var aMenu1 = new Array('','5','95','','','','','','','Solicitação de Acesso - Aprovisionador','SolicitacaoAprov_main.asp',0,'Solicitação de Acesso - Cancelamento','AcessoLogico.asp?acao=can',0,'Solicitação de Acesso - Desativação','AcessoLogico.asp?acao=des',0,'Entregar/Ativar serviço','javascript:document.cookie = ""COOK_LOCATION = MONITORACAO"";window.location.replace(""monitoracaoEntregaAtivacao.asp"");',0);" & Chr(13)
		Call objDicPag.Add("ALTERACAOMAIN.ASP","ALTERACAOMAIN.ASP")
	End if
	blnAcessoLog = true
	intTotalCelula = intTotalCelula + 1
End if

''@@@@@@@@@@@@ AVALIADOR
if objDicCef.Exists("AVL") then
  Call objDicPag.Add("AVALIARACESSO.ASP","AVALIARACESSO.ASP")
end if 


'' ***  AMEDAWA - vai virar tarefas
''@@@@@@@@@@@@ TAREFAS


if objDicCef.Exists("GAT") OR  objDicCef.Exists("AVL") OR objDicCef.Exists("E") OR objDicCef.Exists("GAE") OR objDicCef.Exists("MT") OR objDicCef.Exists("MD") OR objDicCef.Exists("SMB") then
		strGerencia = strGerencia & "'Solicitação','Pendenciaterc.asp?hdnTipo=S',0,'Pedido Terceiro / EBT ADE','Pendenciaterc.asp?hdnTipo=T',0);" & Chr(13)
		Call objDicPag.Add("PENDENCIATERC.ASP","PENDENCIATERC.ASP")
		blnGerencia = true
end if

''@@@@@@@@@@@@ ACESSO FISICO

'Aceite
 TotalAcessoFis = 0
' strAcessoFis =  "'Embratel','',0,'Terceiro ADE/Cliente','',2);"


''reestru if objDicCef.Exists("ATE") or objDicCef.Exists("MD") or objDicCef.Exists("MT") or objDicCef.Exists("GE") then

	''reestru TotalAcessoFis =  TotalAcessoFis + 1
   ' ****  ALTEREI - AMEDAWA
	' strAcessoFisSub = strAcessoFisSub & "'Aceite','aceitacao_main.asp',0,"
	''reestru  strAcessoFis = strAcessoFis & "'Aceite','aceitacao_main.asp',0,"	

	''reestru Call objDicPag.Add("ACEITACAO_MAIN.ASP","ACEITACAO_MAIN.ASP")
	''reestru Call objDicPag.Add("ACEITACAO.ASP","ACEITACAO.ASP")

	''reestru blnAcessoFis = true
	''reestru intTotalCelula = intTotalCelula + 1
''reestru End if

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
'Agendamento de Alteração
'''if objDicCef.Exists("E") or objDicCef.Exists("GE") then
	''' TotalAcessoFis =  TotalAcessoFis + 1
	'strAcessoFisSub = strAcessoFisSub & "'Agendamento de Alteração do Serviço/Acesso','javascript:document.cookie = ""COOK_LOCATION = CONSULTA"";window.location.replace(""Monitoracao.asp"");',0,"
 	''' strAcessoFis = strAcessoFis & "'Agendamento de Alteração do Serviço/Acesso','javascript:document.cookie = ""COOK_LOCATION = CONSULTA"";window.location.replace(""Monitoracao.asp"");',0,"

	'''blnAcessoFis = true
'''End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
'''if objDicCef.Exists("PST") then
   ''' TotalAcessoFis 	= TotalAcessoFis + 1
	'strAcessoFisSub = strAcessoFisSub & "'Alteração na Alocação de Facilidade(Manobra)','Manobra.asp',0,"
'''	strAcessoFis 		= strAcessoFis & "'Alteração na Alocação de Facilidade(Manobra)','ManobraFac.asp',0,"

	'''Call objDicPag.Add("MANOBRAFAC.ASP","MANOBRAFAC.ASP")
	'''blnAcessoFis = true
'''End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
'Execução / Pendência de Manobra e Pendência de Retirada de Interligação
'''if objDicCef.Exists("MD") or objDicCef.Exists("MT") or objDicCef.Exists("AGE")or objDicCef.Exists("GE") then
   '''TotalAcessoFis 	=  TotalAcessoFis + 1
	'strAcessoFisSub 	= strAcessoFisSub & "'Execução','execucao_main.asp',0,"
	'''strAcessoFis 		= strAcessoFis & "'Execução','execucao_main.asp',0,"
'''End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
'''if objDicCef.Exists("GAT") or objDicCef.Exists("GE") then
   '''TotalAcessoFis 	=  TotalAcessoFis + 1
	'strAcessoFisSub 	= strAcessoFisSub & "'Liberação de Estoque','DesativacaoLote.asp',0,"
	'''strAcessoFis 		= strAcessoFis  & "'Liberação de Estoque','DesativacaoLote.asp',0,"
	'''Call objDicPag.Add("DESATIVACAOLOTE.ASP","DESATIVACAOLOTE.ASP")
'''End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
'''if objDicCef.Exists("MD") or objDicCef.Exists("MT") or objDicCef.Exists("AGE")or objDicCef.Exists("GE") then
   '''TotalAcessoFis =  TotalAcessoFis + 2
'	strAcessoFisSub = strAcessoFisSub & "'Pendência de Manobra','PENDENCIAMANOBRA.ASP',0,"
'	strAcessoFisSub = strAcessoFisSub & "'Pendência de Retirada de Interligação','PedenciaInterligacaoRet.asp',0,"

	'''strAcessoFis = strAcessoFis & "'Pendência de Manobra','PENDENCIAMANOBRA.ASP',0,"
	'''strAcessoFis = strAcessoFis & "'Pendência de Retirada de Interligação','PedenciaInterligacaoRet.asp',0,"

	'''Call objDicPag.Add("PEDENCIAINTERLIGACAORET.ASP","PEDENCIAINTERLIGACAORET.ASP")
	'''Call objDicPag.Add("EXECUCAO_MAIN.ASP","EXECUCAO_MAIN.ASP")
	'''Call objDicPag.Add("EXECUCAO.ASP","EXECUCAO.ASP")
	'''Call objDicPag.Add("PENDENCIAMANOBRA.ASP","PENDENCIAMANOBRA.ASP")
	'''call objdicpag.add("CONFIRMACAOMANOBRA.ASP","CONFIRMACAOMANOBRA.ASP")

	'''blnAcessoFis = true
	'''intTotalCelula = intTotalCelula + 1
'''End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

'Pedido
''reestruturacao
''if objDicCef.Exists("GAT") or objDicCef.Exists("GAE") or objDicCef.Exists("GE") then

if  objDicCef.Exists("MTE")  then
	TotalAcessoFis =  TotalAcessoFis + 3
	
	strAcessoFis = strAcessoFis & "'Alocação / Aceite','AlocacaoNew_main.ASP',0,"
	
	strAcessoFis = strAcessoFis & "'Alteração de Facilidades','ManutencaoAlocacaoNew_main.ASP',0,"

	strAcessoFis = strAcessoFis & "'Desalocação','DesalocacaoNew_main.ASP',0,"
	


	'strAcessoFis = strAcessoFis & "'Desalocação','DesalocacaoNew_main.ASP',0,"
	
	Call objDicPag.Add("MANUTENCAOALOCACAONEW_MAIN.ASP","MANUTENCAOALOCACAONEW_MAIN.ASP")
	Call objDicPag.Add("MANUTENCAOALOCACAONEW.ASP","MANUTENCAOALOCACAONEW.ASP")
	
	
	Call objDicPag.Add("ALOCACAONEW_MAIN.ASP","ALOCACAONEW_MAIN.ASP")
	Call objDicPag.Add("ALOCACAONEW.ASP","ALOCACAONEW.ASP")

	Call objDicPag.Add("DESALOCACAONEW_MAIN.ASP","DESALOCACAONEW_MAIN.ASP")
	Call objDicPag.Add("DESALOCACAONEW.ASP","DESALOCACAONEW.ASP")
	
	blnAcessoFis = true
	intTotalCelula = intTotalCelula + 1

End if 

if objDicCef.Exists("GAT") or objDicCef.Exists("GAE")  then
   TotalAcessoFis =  TotalAcessoFis + 2
	' strAcessoFisSub = strAcessoFisSub & "'Alocação de Facilidades','facilidade_main.asp',0,"
	''reestru strAcessoFis = strAcessoFis & "'Alocação de Facilidades','facilidade_main.asp',0,"

	'''retirada 260325
	'''strAcessoFis = strAcessoFis & "'Alocação / Aceite','AlocacaoNew_main.ASP',0,"

	'''strAcessoFis = strAcessoFis & "'Desalocação','DesalocacaoNew_main.ASP',0,"

	''' fim retirada 260325
	strAcessoFis = strAcessoFis & "'Solicitação de Pedido SNOA','SolicPedSNOA_main.asp',0,"

	'strAcessoFis = strAcessoFis & "'Solicitação de Pedido SNOA Manual','solicPedSNOA_CancManual.asp',0,"

	strAcessoFis = strAcessoFis & "'Tratar Providencia SNOA','TratarProvidenciaSNOA_main.asp',0,"

  	strAcessoFis = strAcessoFis & "'Criação/Envio de Carta ao Provedor','CartaProvedor.asp',0,"
  
  	call objdicpag.add("CARTAPROVEDOR.ASP","CARTAPROVEDOR.ASP")
	'Call objDicPag.Add("FACILIDADE_MAIN.ASP","CARTAPROVEDOR.ASP")


	call objdicpag.add("RELATORIOACESSOFISICO_MAIN.ASP","RELATORIOACESSOFISICO_MAIN.ASP")

	
	''reestru Call objDicPag.Add("FACILIDADE_MAIN.ASP","FACILIDADE_MAIN.ASP")
	''reestru Call objDicPag.Add("FACILIDADE.ASP","FACILIDADE.ASP")

	'''retirada 260325
	'''Call objDicPag.Add("ALOCACAONEW_MAIN.ASP","ALOCACAONEW_MAIN.ASP")
	'''Call objDicPag.Add("ALOCACAONEW.ASP","ALOCACAONEW.ASP")

	'''Call objDicPag.Add("DESALOCACAONEW_MAIN.ASP","DESALOCACAONEW_MAIN.ASP")
	'''Call objDicPag.Add("DESALOCACAONEW.ASP","DESALOCACAONEW.ASP")
	
	'''fim retirada 260325



	Call objDicPag.Add("FACILIDADE_NEW.ASP","FACILIDADE_NEW.ASP")
	Call objDicPag.Add("FACILIDADE_NEW_CNS.ASP","FACILIDADE_NEW_CNS.ASP")

	Call objDicPag.Add("SOLICPEDSNOA_MAIN.ASP","SOLICPEDSNOA_MAIN.ASP")
	'Call objDicPag.Add("FACILIDADE.ASP","FACILIDADE.ASP")

	'Call objDicPag.Add("SOLICPEDSNOA_CANCMANUAL.ASP","SOLICPEDSNOA_CANCMANUAL.ASP")

	Call objDicPag.Add("TRATARPROVIDENCIASNOA_MAIN.ASP","TRATARPROVIDENCIASNOA_MAIN.ASP")

	Call objDicPag.Add("EMAILPROVEDOR.ASP","EMAILPROVEDOR.ASP")
	Call objDicPag.Add("EMAILATT.ASP","EMAILATT.ASP")
	blnAcessoFis = true
	intTotalCelula = intTotalCelula + 1
End if


'if  objDicCef.Exists("MTE")  then
'	TotalAcessoFis =  TotalAcessoFis + 2
	
'	strAcessoFis = strAcessoFis & "'Manutencao de Alocação,'ManutencaoAlocacaoNew_main.ASP',0,"

	'strAcessoFis = strAcessoFis & "'Desalocação','DesalocacaoNew_main.ASP',0,"
	
''	Call objDicPag.Add("MANUTENCAOALOCACAONEW_MAIN.ASP","MANUTENCAOALOCACAONEW_MAIN.ASP")
''	Call objDicPag.Add("MANUTENCAOALOCACAONEW.ASP","MANUTENCAOALOCACAONEW.ASP")

	'Call objDicPag.Add("DESALOCACAONEW_MAIN.ASP","DESALOCACAONEW_MAIN.ASP")
	'Call objDicPag.Add("DESALOCACAONEW.ASP","DESALOCACAONEW.ASP")
	
''	blnAcessoFis = true
''	intTotalCelula = intTotalCelula + 1

''End if 

'Pendente OS Provedor
if objDicCef.Exists("GAT") or objDicCef.Exists("GAE") or objDicCef.Exists("GE") then
   TotalAcessoFis =  TotalAcessoFis + 1
	strAcessoFis = strAcessoFis & "'Pendente OS Provedor','CadastraOSProvedor.asp',0,"


	Call objDicPag.Add("CADASTRAOSPROVEDOR.ASP","CADASTRAOSPROVEDOR.ASP")
	Call objDicPag.Add("CADASTRAOS.ASP","CADASTRAOS.ASP")
	Call objDicPag.Add("REALIZACADASTROOS.ASP","REALIZACADASTROOS.ASP")
	blnAcessoFis = true
	intTotalCelula = intTotalCelula + 1
End if

' ************* AMEDAWA ***********

''@@@@@@@@@@@@ RELATORIOS

   	   Call objDicPag.Add("RELATORIOS.ASP","RELATORIOS.ASP")
	   strRelatorios = "'Acesso por Estado','../relatorios/mapa.ASP',0,'Acesso por Diretoria','../relatorios/mapa_diretoria.asp',0,'Acesso por Cliente','../relatorios/mapa_cliente.asp',0,'Acesso por Logradouro','../relatorios/mapa_endereco.asp',0"

'JC***
'		strTabela = strTabela & _
'				"'Histórico - Migração HFC BSoD','historico_migracaoBsod.asp',0,"
'		Call objDicPag.Add("HISTORICO_MIGRACAOBSOD.ASP","HISTORICO_MIGRACAOBSOD.ASP")
''Histórico - Migração HFC BSoD','historico_migracaoBsod.asp',0,

Call objDicPag.Add("HISTORICO_MIGRACAOBSOD.ASP","HISTORICO_MIGRACAOBSOD.ASP")

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
if objDicCef.Exists("GAT") then
	'Inclusão da nova Tela "Consulta OS Provedor" no menu - T3GREYN - ACCENTURE - 17/09/2009
	
	''' ReCLA
	'''strConsulta = strConsulta & "'Acessos Aceitos','ConsAcessosAceitos.asp',0,'Acessos Ativados Comercialmente','RelaAtivComercialmente.asp', 0 , 'Acessos Ativos (Lista)','consAtivosLista.asp',0,'Acessos Ativos por Velocidade','consAtivosResumoVel.asp',0,'Acessos Físicos','ConsultaAcessoFisico.asp',0,'Acessos Instalados (Lista)','consInstaladosLista.asp',0,'Acessos Instalados (Resumo)','consInstaladosResumo.asp',0,'Acessos Instalados (Sem Expurgo)','consInstaladosResumoExp.asp',0,'Acompanhamento','consacompanhamento.asp',0,'Controle de Acessos de Serviços Ativados','ConsControleAcesso.asp',0,'Controle de Rede Determinística','ConsRedeDet.asp',0,'Cartas ao Provedor','CartasProvedor.asp',0,'Desempenho do GLA - A. Aceitos','cons_GLA.asp',0,'Expressa','Expressa_main.asp',0,'Facilidades','ConsultarFacilidadeRec.asp',0,'Geral','consulta_main.asp',0,'Histórico de Pedido','historico_main.asp',0,'Histórico de Pedido SNOA','historicopedidosnoa_main.asp',0,'Histórico de Facilidades','cons_historicofacilidade.asp',0,'Monitoração','javascript:document.cookie = ""COOK_LOCATION = MONITORACAO"";window.location.replace(""Monitoracao.asp"");',0,'OS Provedor','ConsultaOSProvedor.asp',0," & _ 
	'''	"'Mais...','',22);"
		
		
'   "'Histórico - Migração HFC BSoD','historico_migracaoBsod.asp',0,'Mais...','',22);"


	strConsulta = strConsulta & "'Cartas ao Provedor','CartasProvedor.asp',0," & _
		"'Geral','consulta_main.asp',0," & _
		"'Histórico de Pedido','historico_main.asp',0,"& _
		"'Histórico de Pedido SNOA','historicopedidosnoa_main.asp',0," & _
		"'Histórico de Facilidades','cons_historicofacilidade.asp',0," & _
		"'Relatório de Pesquisa Acesso Fisico','RelatorioAcessoFisico_main.asp',0);" 

Else
	''' ReCLA
	'''strConsulta = strConsulta & "'Acessos Aceitos','ConsAcessosAceitos.asp',0,'Acessos Ativados Comercialmente','RelaAtivComercialmente.asp', 0 ,'Acessos Ativos (Lista)','consAtivosLista.asp',0,'Acessos Ativos por Velocidade','consAtivosResumoVel.asp',0,'Acessos Físicos','ConsultaAcessoFisico.asp',0,'Acessos Instalados (Lista)','consInstaladosLista.asp',0,'Acessos Instalados (Resumo)','consInstaladosResumo.asp',0,'Acessos Instalados(Sem Expurgo)','consInstaladosResumoExp.asp',0,'Acompanhamento','consacompanhamento.asp',0,'Controle de Acessos de Serviços Ativados','ConsControleAcesso.asp',0,'Controle de Rede Determinística','ConsRedeDet.asp',0,'Desempenho do GLA - A. Aceitos','cons_GLA.asp',0,'Facilidades','ConsultarFacilidadeRec.asp',0,'Expressa','Expressa_main.asp',0,'Geral','consulta_main.asp',0,'Histórico de Pedido','historico_main.asp',0,'Histórico de Pedido SNOA','historicopedidosnoa_main.asp',0,'Histórico de Facilidades','cons_historicofacilidade.asp',0,'Monitoração','javascript:document.cookie = ""COOK_LOCATION = MONITORACAO"";window.location.replace(""Monitoracao.asp"");',0,'Pendência de Cabo Interno','ConsPendCaboInterno.asp',0,'OS Provedor','ConsultaOSProvedor.asp',0," & _
	'''  "'Mais...','',22);"
	  
	 strConsulta = strConsulta & "'Cartas ao Provedor','CartasProvedor.asp',0," & _
		"'Geral','consulta_main.asp',0," & _
		"'Histórico de Pedido','historico_main.asp',0,"& _
		"'Histórico de Pedido SNOA','historicopedidosnoa_main.asp',0," & _
		"'Histórico de Facilidades','cons_historicofacilidade.asp',0," & _
		"'Relatório de Pesquisa Acesso Fisico','RelatorioAcessoFisico_main.asp',0);" 
	 ' "'Histórico - Migração HFC BSoD','historico_migracaoBsod.asp',0,'Mais...','',22);"
End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

'''strConsulta = strConsulta & "'Expressa','Expressa_main.asp',0,'Geral','consulta_main.asp',0,'Histórico de Pedido','historico_main.asp',0,'Histórico de Pedido SNOA','historicopedidosnoa_main.asp',0,'OS Provedor','ConsultaOSProvedor.asp',0," & _  

'''"'Mais...','',22);"
''@@@@@@@@@@@@ TABELA


if objDicCef.Exists("PST")  or objDicCef.Exists("REC")  or objDicCef.Exists("GE") or objDicCef.Exists("GAT") or objDicCef.Exists("MD") or objDicCef.Exists("MT") or objDicCef.Exists("E") or objDicCef.Exists("GAE") or objDicCef.Exists("ML") or objDicCef.Exists("SMB") then
	blnTabelas = true
	strUserHerderPerfil = "REC"
End if

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
'''if objDicCef.Exists("SMB") and (strloginrede = "EDAR" or strloginrede = "PRSSILV" or strloginrede = "T3FRRP" ) then
'''	strTabela = strTabela & "'SMBD','smbd_main.asp',0,"
'''	Call objdicpag.add("SMBD_MAIN.ASP","SMBD_MAIN.ASP")
'''	intTotalMenu = intTotalMenu + 1
'''End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
'********************* Migração BSOD
'''If objDicCef.Exists("PST") or objDicCef.Exists("GAT") or objDicCef.Exists("GAE") Then
		
		'''strTabela = strTabela & _
		'''		"'Migração HFC BSoD','migracaoBsod.asp',0,"
		'''Call objdicpag.add("MIGRACAOBSOD.ASP","MIGRACAOBSOD.ASP")
		
'		strTabela = strTabela & _
'				"'Histórico - Migração HFC BSoD','historico_migracaoBsod.asp',0,"
'		Call objDicPag.Add("HISTORICO_MIGRACAOBSOD.ASP","HISTORICO_MIGRACAOBSOD.ASP")


		
		'''intTotalMenu = intTotalMenu + 1'2
		
'''End If
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA


''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
'''if objDicCef.Exists("GE") then
	'''strTabela = strTabela & _
	'''			"'Agente','agente_main.asp',0," & _
	'''			"'Agente do Pedido','agentepedido_main.asp',0,"

	'''Call objdicpag.add("AGENTE_MAIN.ASP","AGENTE_MAIN.ASP")
	'''Call objdicpag.add("AGENTE.ASP","AGENTE.ASP")
	'''Call objDicPag.Add("AGENTEPEDIDO_MAIN.ASP","AGENTEPEDIDO_MAIN.ASP")
	'''Call objDicPag.Add("AGENTEPEDIDO.ASP","AGENTEPEDIDO.ASP")

	'''intTotalMenu = intTotalMenu + 2

'''End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
'''if objDicCef.Exists("GE") or objDicCef.Exists("GAT") or objDicCef.Exists("E") or objDicCef.Exists("REC") then

	'''strTabela = strTabela &  "'Associação Designação de Tronco 2M','DesignacaoTronco2M_main.asp',0," 
	strTabela = strTabela &  "'Consulta Estação/Associação Estação CF','Assoc_estacaoCtf_main.asp',0," 
					
		
		Call objdicpag.add("MANTER_ESTACAO.ASP","MANTER_ESTACAO.ASP")	
		Call objDicPag.Add("DESIGNACAOTRONCO2M_MAIN.ASP","DESIGNACAOTRONCO2M_MAIN.ASP")
		Call objDicPag.Add("DESIGNACAOTRONCO2M.ASP","DESIGNACAOTRONCO2M.ASP")
		Call objDicPag.Add("INCLUIRDESIGNACAOTRONCO2M.ASP","INCLUIRDESIGNACAOTRONCO2M.ASP")
	
	  Call objdicpag.add("ASSOC_ESTACAOCTF_MAIN.ASP","ASSOC_ESTACAOCTF_MAIN.ASP")
	  Call objdicpag.add("ASSOC_ESTACAOCTF.ASP","ASSOC_ESTACAOCTF.ASP")
	  
	'''intTotalMenu = intTotalMenu + 2
		intTotalMenu = intTotalMenu + 1
'''End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

'PRSSILV - Autorização Acesso - 16/03/2009

'''Set objRSPerf = db.execute("CLA_sp_view_loginusuario '" & strLoginRede & "'")
'''If not objRSPerf.eof then
'''  var_Usu_PerfCadSenha = objRSPerf("Usu_PerfCadSenha")
'''  var_Usu_PerfAltDesig = objRSPerf("Usu_PerfAltDesig")
'''End if
'''if var_Usu_PerfCadSenha = 1 then
'''	strTabela = strTabela &  "'Autorização de Acessos Terceiros','CONSULTARAPROVARACESSO.asp',0," 
	'Call objDicPag.Add("PERFCADSENHA2M_MAIN.ASP","PERFCADSENHA2M_MAIN.ASP")
	'Call objDicPag.Add("INCLUIRAUTORIZARACESSO.ASP","INCLUIRAUTORIZARACESSO.ASP")
	'Call objDicPag.Add("CONSULTARAUTORIZARACESSO.ASP","CONSULTARAUTORIZARACESSO.ASP")
	'''Call objDicPag.Add("CONSULTARAPROVARACESSO.ASP","CONSULTARAPROVARACESSO.ASP")
	'Call objDicPag.Add("ALTERARAUTORIZARACESSO.ASP","ALTERARAUTORIZARACESSO.ASP")
	'Call objDicPag.Add("INCLUIRAUTORIZARACESSO.ASP","INCLUIRAUTORIZARACESSO.ASP")
	'''intTotalMenu = intTotalMenu + 1

'''End if


''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
''if objDicCef.Exists("GAT") OR  objDicCef.Exists("E") OR objDicCef.Exists("PST") OR objDicCef.Exists("C") OR objDicCef.Exists("SMB") then
	'''strTabela = strTabela &  "'Alteração de Designação - Conciliação','ALTERARDESIGNACAO.asp',0," 
	'''Call objDicPag.Add("ALTERARDESIGNACAO.ASP","ALTERARDESIGNACAO.ASP")
	'''Call objDicPag.Add("LISTARDESIGNACAO.ASP","LISTARDESIGNACAO.ASP")
	'''Call objDicPag.Add("PROCESSOLISTARDESIGNACAO.ASP","PROCESSOLISTARDESIGNACAO.ASP")
	'''Call objDicPag.Add("PROCESSOALTERARDESIGNACAO.asp.ASP","PROCESSOALTERARDESIGNACAO.ASP")
	
	'intTotalMenu = intTotalMenu + 1


	'CH-75110JLS- Inicio'
	
	'''if objDicCef.Exists("E") OR  objDicCef.Exists("C")   then '- Perfil GICL e GICN -'
		
	'''	if  not objDicCef.Exists("GE") then '- Perfil GER.USUARIO -'
	'''		strTabela = strTabela &  "'Estação','estacao_main.asp',0,"

	'''		Call objdicpag.add("ESTACAO_MAIN.ASP","ESTACAO_MAIN.ASP")
	'''		Call objdicpag.add("MANTER_ESTACAO.ASP","MANTER_ESTACAO.ASP")

	'''		Call objdicpag.add("ESTACAO.ASP","ESTACAO.ASP")

	'''	end if
	
	'''end if

	'''intTotalMenu = intTotalMenu + 1

	'CH-75110JLS- Fim'
	

'''End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
if objDicCef.Exists("GE") then
	strTabela = strTabela & _
				
				"'Associação Usuário com CF','AssocUserCef_main.asp',0," & _
				"'Associação de Serviço com Modalidade', 'AssocServMod_main.asp',0," & _
				"'Associação de Serviço com Velocidade', 'AssocServVel_main.asp',0," & _
				"'Associação de Tecnologia com Velocidade', 'AssocTecVel_main.asp',0," & _
			
			    "'Centro Funcional','CentroFuncional_main.asp',0," & _
			

				"'Estação','estacao_main.asp',0,"

	
	Call objDicPag.Add("ACAO.ASP","ACAO.ASP")
	Call objdicpag.add("ASSOCUSERCEF_MAIN.ASP","ASSOCUSERCEF_MAIN.ASP")
	Call objdicpag.add("ASSOCUSERCEF.ASP","ASSOCUSERCEF.ASP")
	Call objdicpag.add("ASSOCSERVMOD_MAIN.ASP","ASSOCSERVMOD_MAIN.ASP")
	Call objdicpag.add("ASSOCSERVMOD.ASP","ASSOCSERVMOD.ASP")
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
	Call objdicpag.add("MANTER_ESTACAO.ASP","MANTER_ESTACAO.ASP")

	Call objdicpag.add("ESTACAO.ASP","ESTACAO.ASP")

	
	'''intTotalMenu = intTotalMenu + 10
	
	intTotalMenu = intTotalMenu + 6

	strUserHerderPerfil = "GE"

End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
'GPON
'''if objDicCef.Exists("PST") then
	'''if strloginrede = "PRSSILV" OR strloginrede = "EDAR" OR strloginrede = "MONIB" OR strloginrede = "RUTH" or strloginrede = "T3FRRP" then
	'''	strTabela = strTabela &  "'Empreiteira','Empreiteira_main.asp',0,"
	'''	Call objdicpag.add("EMPREITEIRA_MAIN.ASP","EMPREITEIRA_MAIN.ASP")
	'''	intTotalMenu = intTotalMenu + 1
	'''end if
'''end if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
'''if objDicCef.Exists("GE") or objDicCef.Exists("GAT") or objDicCef.Exists("GAE") then
	''xib
	'strTabela = strTabela & "'Facilidade','../asp/CadastrarFacilidade.asp',0,'Grupos','../asp/grupo_main.asp',0,'Origem Solicitação','../asp/origemsol_main.asp',0,"
	
'''	strTabela = strTabela & "'Grupos','../asp/grupo_main.asp',0,'Origem Solicitação','../asp/origemsol_main.asp',0,"
	''strTabela = strTabela & "'Facilidade','CadastrarFacilidade.asp',0,"
'''	Call objdicpag.add("GRUPO_MAIN.ASP","GRUPO_MAIN.ASP")
'''	Call objdicpag.add("GRUPO.ASP","GRUPO.ASP")
	'Call objdicpag.add("CADASTRARFACILIDADE.ASP","CADASTRARFACILIDADE.ASP")
'''	Call objdicpag.add("origemsol_main.asp","origemsol_main.asp")
'''	Call objdicpag.add("origemsol.asp","origemsol.asp")
	
'''	intTotalMenu = intTotalMenu + 2
''' End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

if objDicCef.Exists("REC")  then
	strTabela = strTabela & "'Facilidade','../asp/CadastrarFacilidade.asp',0,"
	''strTabela = strTabela & "'Facilidade','CadastrarFacilidade.asp',0,"
	'Call objdicpag.add("GRUPO_MAIN.ASP","GRUPO_MAIN.ASP")
	'Call objdicpag.add("GRUPO.ASP","GRUPO.ASP")
	Call objdicpag.add("CADASTRARFACILIDADE.ASP","CADASTRARFACILIDADE.ASP")
	'Call objdicpag.add("origemsol_main.asp","origemsol_main.asp")
	'Call objdicpag.add("origemsol.asp","origemsol.asp")
	
	intTotalMenu = intTotalMenu + 1
End if

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
'GPON
'''if objDicCef.Exists("PST") then
	'''strTabela = strTabela &  "'ONT Fabricante','FabricanteONT_main.asp',0,"
	'''strTabela = strTabela &  "'ONT Modelo','ModeloONT_main.asp',0,"
	'''Call objdicpag.add("MODELOONT_MAIN.ASP","MODELOONT_MAIN.ASP")
	'''Call objdicpag.add("FABRICANTEONT_MAIN.ASP","FABRICANTEONT_MAIN.ASP")
	'''Call objdicpag.add("CAD_FABRICANTEONT.ASP","CAD_FABRICANTEONT.ASP")
	'''Call objdicpag.add("CAD_MODELOONT.ASP","CAD_MODELOONT.ASP")
	
	'''intTotalMenu = intTotalMenu + 2
'''end if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

if objDicCef.Exists("PST") then
	strTabela = strTabela & "'Holding','Holding_main.asp',0,"
	Call objdicpag.add("HOLDING_MAIN.ASP","HOLDING_MAIN.ASP")
	Call objdicpag.add("HOLDING.ASP","HOLDING.ASP")
	intTotalMenu = intTotalMenu + 1
End if

''Raio x
'if objDicCef.Exists("GE") or objDicCef.Exists("GAT") then
'	strTabela = strTabela & "'Logradouro','logradouro_main.asp',0,"
'	Call objdicpag.add("LOGRADOURO_MAIN.ASP","LOGRADOURO_MAIN.ASP")
'	Call objdicpag.add("LOGRADOURO.ASP","LOGRADOURO.ASP")

'	intTotalMenu = intTotalMenu + 1
'End if
''Raio x

if objDicCef.Exists("GE") then
	strTabela = strTabela & "'Localidade Subordinada ao CF','LocalidadeSubCef_main.asp',0,"
	Call objdicpag.add("LOCALIDADESUBCEF_MAIN.ASP","LOCALIDADESUBCEF_MAIN.ASP")
	Call objdicpag.add("LOCALIDADESUBCEF.ASP","LOCALIDADESUBCEF.ASP")
	intTotalMenu = intTotalMenu + 1
End if

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
'''if objDicCef.Exists("GE") then
'''	strTabela = strTabela & "'Modalidade','Modalidade_main.asp',0,"
'''	Call objdicpag.add("MODALIDADE_MAIN.ASP","MODALIDADE_MAIN.ASP")
'''	Call objdicpag.add("MODALIDADE.ASP","MODALIDADE.ASP")
'''	intTotalMenu = intTotalMenu + 1
'''End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
''xib
'if objDicCef.Exists("GE") or objDicCef.Exists("GAT") or objDicCef.Exists("GAE") or objDicCef.Exists("MD") or objDicCef.Exists("MT") then

if objDicCef.Exists("REC") then
	intTotalMenu = intTotalMenu + 1
	strTabela = strTabela & "'Manutenção de Facilidade','ManutencaoFacilidade.asp',0,"
	Call objdicpag.add("MANUTENCAOFACILIDADE.ASP","MANUTENCAOFACILIDADE.ASP")
End if

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
'''if objDicCef.Exists("GE") then

'''	intTotalMenu = intTotalMenu + 2
'''	strTabela = strTabela & "'Órgão','orgao_main.asp',0,'Mais...',''," & intTotalMenu & ");"
'SubMenu - Tabela
'''	strTabelaSub_1 = strTabelaSub_1 & "'Parâmetros CF','CefConfig_main.asp',0,"
'''	strTabelaSub_1 = strTabelaSub_1 & "'Plataforma','Plataforma_main.asp',0,"


'''	Call objdicpag.add("ORGAO_MAIN.ASP","ORGAO_MAIN.ASP")
'''	Call objdicpag.add("ORGAO.ASP","ORGAO.ASP")
'''	call objdicpag.add("CEFCONFIG_MAIN.ASP","CEFCONFIG_MAIN.ASP")
'''	call objdicpag.add("CEFCONFIG.ASP","CEFCONFIG.ASP")
'''	call objdicpag.add("PLATAFORMA_MAIN.ASP","CEFCONFIG_MAIN.ASP")
'''	call objdicpag.add("PLATAFORMA.ASP","CEFCONFIG.ASP")
'''End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA


'if objDicCef.Exists("PST") or objDicCef.Exists("GE")then

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
if objDicCef.Exists("MPR") or objDicCef.Exists("GE")then

'''strTabelaSub_1 = strTabelaSub_1 & "'Promoção','Promocao_main.asp',0,"
strTabelaSub_1 = strTabelaSub_1 & "'Provedor','provedor_main.asp',0,"
strTabelaSub_1 = strTabelaSub_1 & "'Provedor Contrato','provedorcontrato_main.asp',0,"

	call objdicpag.add("PROMOCAO_MAIN.ASP","PROMOCAO_MAIN.ASP")
	call objdicpag.add("PROMOCAO.ASP","PROMOCAO.ASP")
	call objdicpag.add("PROVEDOR_MAIN.ASP","PROVEDOR_MAIN.ASP")
	call objdicpag.add("PROVEDOR.ASP","PROVEDOR.ASP")
	call objdicpag.add("PROVEDORCONTRATO_MAIN.ASP" ,"PROVEDORCONTRATO_MAIN.ASP")
	call objdicpag.add("PROVEDORCONTRATO.ASP" ,"PROVEDORCONTRATO.ASP")

End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA


''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA

''' if objDicCef.Exists("MPR") or objDicCef.Exists("GE")then

if  objDicCef.Exists("TEC") then

'''strTabelaSub_1 = strTabelaSub_1 & "'Promoção','Promocao_main.asp',0,"
strTabelaSub_1 = strTabelaSub_1 & "'Tecnologia','tecnologia_main.asp',0,"
strTabelaSub_1 = strTabelaSub_1 & "'Associacao Tecnologia com Facilidade','assocTecnologiaFacilidade_main.asp',0,"

strTabelaSub_1 = strTabelaSub_1 & "'Combo','combo_main.asp',0,"
strTabelaSub_1 = strTabelaSub_1 & "'Estrutura Combo','estruturaCombo_main.asp',0,"

strTabelaSub_1 = strTabelaSub_1 & "'Estrutura Tecnologia','estruturaTecnologia_main.asp',0,"
'strTabelaSub_1 = strTabelaSub_1 & "'Ordenação Estrutura Tecnologia','ordenacaoEstruturaTecnologia_main.asp',0,"

	call objdicpag.add("TECNOLOGIA_MAIN.ASP","TECNOLOGIA_MAIN.ASP")
	call objdicpag.add("TECNOLOGIA.ASP","TECNOLOGIA.ASP")
	call objdicpag.add("ASSOCTECNOLOGIAFACILIDADE_MAIN.ASP","ASSOCTECNOLOGIAFACILIDADE_MAIN.ASP")
	call objdicpag.add("ASSOCTECNOLOGIAFACILIDADE.ASP","ASSOCTECNOLOGIAFACILIDADE.ASP")
	
	call objdicpag.add("COMBO_MAIN.ASP","COMBO_MAIN.ASP")
	call objdicpag.add("COMBO.ASP","COMBO.ASP")
	
	call objdicpag.add("ESTRUTURACOMBO_MAIN.ASP" ,"ESTRUTURACOMBO_MAIN.ASP")
	call objdicpag.add("ESTRUTURACOMBO.ASP" ,"ESTRUTURACOMBO.ASP")
	call objdicpag.add("CST_ESTRUTURACOMBO.ASP" ,"CST_ESTRUTURACOMBO.ASP")
	
	call objdicpag.add("ESTRUTURATECNOLOGIA_MAIN.ASP" ,"ESTRUTURATECNOLOGIA_MAIN.ASP")
	call objdicpag.add("ESTRUTURATECNOLOGIA.ASP" ,"ESTRUTURATECNOLOGIA.ASP")
	call objdicpag.add("CST_ESTRUTURATECNOLOGIA.ASP" ,"CST_ESTRUTURATECNOLOGIA.ASP")
	'call objdicpag.add("ORDENACAOESTRUTURATECNOLOGIA_MAIN.ASP" ,"ORDENACAOESTRUTURATECNOLOGIA_MAIN.ASP")
	'call objdicpag.add("ORDENACAOESTRUTURATECNOLOGIA.ASP" ,"ORDENACAOESTRUTURATECNOLOGIA.ASP")

End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
'''if objDicCef.Exists("PST") then
'''strTabelaSub_1 = strTabelaSub_1 &	"'Preço-Provedor','postFile.asp',0,"

'''	Call objDicPag.Add("POSTFILE.ASP","POSTFILE.ASP")
'''End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
'''if objDicCef.Exists("GE") or objDicCef.Exists("GAT") then
''xib
'strTabelaSub_1 = strTabelaSub_1 &	"'Recurso','recurso_main.asp',0," & _
'							"'Redirecionar Solicitação','redirsolicitacao_main.asp',0,"
							
'''strTabelaSub_1 = strTabelaSub_1 &	"'Redirecionar Solicitação','redirsolicitacao_main.asp',0,"
	'Call objdicpag.add("RECURSO_MAIN.ASP","RECURSO_MAIN.ASP")
	'Call objdicpag.add("RECURSO.ASP","RECURSO.ASP")
'''	Call objdicpag.add("REDIRSOLICITACAO_MAIN.ASP","REDIRSOLICITACAO_MAIN.ASP")
'''	Call objdicpag.add("REDIRSOLICITACAO.ASP","REDIRSOLICITACAO.ASP")
'''End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

if objDicCef.Exists("REC")  then
strTabelaSub_1 = strTabelaSub_1 &	"'Recurso','recurso_main.asp',0," 
	Call objdicpag.add("RECURSO_MAIN.ASP","RECURSO_MAIN.ASP")
	Call objdicpag.add("RECURSO.ASP","RECURSO.ASP")
	'Call objdicpag.add("REDIRSOLICITACAO_MAIN.ASP","REDIRSOLICITACAO_MAIN.ASP")
	'Call objdicpag.add("REDIRSOLICITACAO.ASP","REDIRSOLICITACAO.ASP")
End if

if objDicCef.Exists("PST") then
strTabelaSub_1 = strTabelaSub_1 & "'Regime de Contrato','regimecontrato_main.asp',0,"
	call objdicpag.add("REGIMECONTRATO_MAIN.ASP","REGIMECONTRATO_MAIN.ASP")
	call objdicpag.add("REGIMECONTRATO.ASP","REGIMECONTRATO.ASP")
End if

if objDicCef.Exists("GE") or objDicCef.Exists("MD") or objDicCef.Exists("MT") or objDicCef.Exists("ML") or objDicCef.Exists("GAE") then
strTabelaSub_1 = strTabelaSub_1 & "'Rede Interna','redeInterna.asp',0,"
	Call objDicPag.Add("REDEINTERNA.ASP","REDEINTERNA.ASP")
End if

if objDicCef.Exists("GE") then
strTabelaSub_1 = strTabelaSub_1 & "'Serviço','servico_main.asp',0," & _
						"'Status','status_main.asp',0,"

	call objdicpag.add("SERVICO_MAIN.ASP","SERVICO_MAIN.ASP")
	call objdicpag.add("SERVICO.ASP","SERVICO.ASP")
	call objdicpag.add("STATUS_MAIN.ASP","STATUS_MAIN.ASP")
	call objdicpag.add("STATUS.ASP","STATUS.ASP")

End if

if objDicCef.Exists("GE") or objDicCef.Exists("PST") then
strTabelaSub_1 = strTabelaSub_1 & "'Tipo de Vigência de Contrato','tipocontrato_main.asp',0,"
	call objdicpag.add("TIPOCONTRATO_MAIN.ASP","TIPOCONTRATO_MAIN.ASP")
	call objdicpag.add("TIPOCONTRATO.ASP","TIPOCONTRATO.ASP")
End if

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
'''if objDicCef.Exists("GE") or objDicCef.Exists("GAE") then
'''strTabelaSub_1 = strTabelaSub_1 & "'Tipo de Rádio','Cadtiporadio_main.asp',0,"
'''	call objdicpag.add("CADTIPORADIO_MAIN.ASP","CADTIPORADIO_MAIN.ASP")
'''	call objdicpag.add("CADTIPORADIO.ASP","CADTIPORADIO.ASP")
'''End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
'''if objDicCef.Exists("GE") or objDicCef.Exists("GAT") or objDicCef.Exists("E") then
'''strTabelaSub_1 = strTabelaSub_1 & "'Transferência de Carteira','TransferenciaCarteira.asp',0," 
'''	call objdicpag.add("TRANSFERENCIACARTEIRA.ASP","TRANSFERENCIACARTEIRA.ASP")
'''End if
''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA

if objDicCef.Exists("GE") then
strTabelaSub_1 = strTabelaSub_1 & "'Usuário','usuario_main.asp',0," & _
							"'Velocidade','velocidade_main.asp',0,"
	call objdicpag.add("USUARIO_MAIN.ASP","USUARIO_MAIN.ASP")
	call objdicpag.add("USUARIO.ASP","USUARIO.ASP")
	call objdicpag.add("VELOCIDADE_MAIN.ASP","VELOCIDADE_MAIN.ASP")
	call objdicpag.add("VELOCIDADE.ASP","VELOCIDADE.ASP")
End if


Call objDicPag.Add("CONSPENDCABOINTERNO.ASP","CONSPENDCABOINTERNO.ASP")
Call objDicPag.Add("RELAATIVCOMERCIALMENTE.ASP" , "RELAATIVCOMERCIALMENTE.ASP")
Call objDicPag.Add("CONSACESSOSACEITOS.ASP","CONSACESSOSACEITOS.ASP")
Call objDicPag.Add("CONSULTA_MAIN.ASP","CONSULTA_MAIN.ASP")
Call objDicPag.Add("EXPRESSA_MAIN.ASP","EXPRESSA_MAIN.ASP")
Call objDicPag.Add("EXPRESSACLIENTE.ASP","EXPRESSACLIENTE.ASP")
Call objDicPag.Add("EXPRESSAENDERECO.ASP","EXPRESSAENDERECO.ASP")
Call objDicPag.Add("EXPRESSAORDERENTRY.ASP","EXPRESSAORDERENTRY.ASP")
Call objDicPag.Add("EXPRESSALISTARORDERENTRY.ASP","EXPRESSALISTARORDERENTRY.ASP")
Call objDicPag.Add("EXPRESSALISTARENDERECOCLIENTE.ASP","EXPRESSALISTARENDERECOCLIENTE.ASP")
Call objDicPag.Add("EXPRESSAENDERECOCLIENTE.ASP","EXPRESSAENDERECOCLIENTE.ASP")
	Call objDicPag.Add("CONSULTACLAAPROVISIONADOR.ASP","CONSULTACLAAPROVISIONADOR.ASP")
Call objDicPag.Add("CONSULTAACESSOFISICO.ASP","CONSULTAACESSOFISICO.ASP")
Call objDicPag.Add("CONSULTAACESSOFISICOENDERECO.ASP","CONSULTAACESSOFISICOENDERECO.ASP")
if objDicCef.Exists("GAT") then
	Call objDicPag.Add("CARTASPROVEDOR.ASP","CARTASPROVEDOR.AS")
End if
Call objDicPag.Add("CONSACOMPANHAMENTO.ASP","CONSACOMPANHAMENTO.ASP")
'Call objDicPag.Add("MONITORACAO.ASP","MONITORACAO.ASP")
Call objDicPag.Add("CONS_ACA.ASP","CONS_ACA.ASP")
Call objDicPag.Add("CONS_PROACA.ASP","CONS_PROACA.ASP")
Call objDicPag.Add("CONS_PROSTP.ASP","CONS_PROSTP.ASP")
Call objDicPag.Add("CONS_GLA.ASP","CONS_GLA.ASP")
Call objDicPag.Add("CONS_PROACAQTD.ASP","CONS_PROACAQTD.ASP")
Call objDicPag.Add("CONS_PROQTD.ASP","CONS_PROQTD.ASP")
Call objDicPag.Add("CONS_PRODATA.ASP","CONS_PRODATA.ASP")
Call objDicPag.Add("CONSVELSERVICO.ASP","CONSVELSERVICO.ASP")
Call objDicPag.Add("HISTORICO_MAIN.ASP","HISTORICO_MAIN.ASP")
Call objDicPag.Add("HISTORICOPEDIDOSNOA_MAIN.ASP","HISTORICOPEDIDOSNOA_MAIN.ASP")
Call objDicPag.Add("CONS_HISTORICOFACILIDADE.ASP","CONS_HISTORICOFACILIDADE.ASP")
Call objDicPag.Add("HISTORICO.ASP","HISTORICO.ASP")
Call objDicPag.Add("CONSULTARFACILIDADEREC.ASP","CONSULTARFACILIDADEREC.ASP")

call objdicpag.add("FACILIDADEDET.ASP","FACILIDADEDET.ASP")
call objdicpag.add("ALTERACAOCAD.ASP","ALTERACAOCAD.ASP")
call objdicpag.add("CONSREDEDET.ASP","CONSREDEDET.ASP")
call objdicpag.add("CONSCONTROLEACESSO.ASP","CONSCONTROLEACESSO.ASP")
'Inclusão da nova Tela "Consulta OS Provedor" no menu - T3GREYN - ACCENTURE - 17/09/2009
call objdicpag.add("CONSULTAOSPROVEDOR.ASP","CONSULTAOSPROVEDOR.ASP")
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


strRelatorioSub =""


if blnAcessoLog and blnGerencia and blnAcessoFis then

	strGerencia	 	= "var aMenu2 = new Array('','132','95','','','','','',''," & strGerencia
	strAcessoFis 	   =  "var aMenu3 = new Array('','258','95','','','','','',''," & strAcessoFis

	strAcessoFisSub  =  "var aMenu3_2 = new Array(" + strAcessoFisSub + ");"



	strConsulta	 	= "var aMenu4 = new Array('','385','95','','','','','',''," & strConsulta
	if objDicCef.Exists("GAT") then
		strConsultaSub 	= " var aMenu4_22= new Array('Acessos Físicos por Endereço','ConsultaAcessoFisicoEndereco.asp',0,'Histórico - Migração HFC BSoD','historico_migracaoBsod.asp',0,'Movimentação de Acessos (Lista)','consMovimentacao.asp',0,'Movimentação de Acessos Temporários (Lista)','consMovimentacaoTemp.asp',0,'Pendência de Cabo Interno','ConsPendCaboInterno.asp',0,'Pedidos por Ação','cons_Aca.asp',0,'Pedidos por Provedor e Ação','cons_ProAca.asp',0,'Pedidos por Provedor e Status','cons_ProStp.asp',0,'Pendentes de Instalação (Lista)','consPendInstalaLista.asp',0,'Pendentes de Instalação (Resumo)','consPendInstalaResumo.asp',0,'Percentual de Acessos com Serviços Ativados','consPercentAcesso.asp',0,'Preço-Provedor','javascript:window.open(""../Excel/PrecoProvedor/PrecoProvedor.xls"");',0,'Relatório de Backlogs Provedor','relapendconstacesso.asp',0,'Relatório Telco','RelaTelco.asp',0,'Qtde. de Pedidos por Provedor e Ação','cons_ProAcaQtd.asp',0,'Qtde. de Pedidos Aceitos por Provedor','cons_ProQtd.asp',0,'Qtde. de Pedidos Pendentes por Provedor','cons_ProData.asp',0,'Velocidade por Serviço','consVelServico.asp',0);"
	else
		strConsultaSub 	= " var aMenu4_22= new Array('Acessos Físicos por Endereço','ConsultaAcessoFisicoEndereco.asp',0,'Histórico - Migração HFC BSoD','historico_migracaoBsod.asp',0,'Movimentação de Acessos (Lista)','consMovimentacao.asp',0,'Movimentação de Acessos Temporários (Lista)','consMovimentacaoTemp.asp',0,'Pendência de Cabo Interno','ConsPendCaboInterno.asp',0,'Pedidos por Ação','cons_Aca.asp',0,'Pedidos por Provedor e Ação','cons_ProAca.asp',0,'Pedidos por Provedor e Status','cons_ProStp.asp',0,'Pendentes de Instalação (Lista)','consPendInstalaLista.asp',0,'Pendentes de Instalação (Resumo)','consPendInstalaResumo.asp',0,'Percentual de Acessos com Serviços Ativados','consPercentAcesso.asp',0,'Preço-Provedor','javascript:window.open(""../Excel/PrecoProvedor/PrecoProvedor.xls"");',0,'Relatório de Backlogs Provedor','relapendconstacesso.asp',0,'Relatório Telco','RelaTelco.asp',0,'Qtde. de Pedidos por Provedor e Ação','cons_ProAcaQtd.asp',0,'Qtde. de Pedidos Aceitos por Provedor','cons_ProQtd.asp',0,'Qtde. de Pedidos Pendentes por Provedor','cons_ProData.asp',0,'Velocidade por Serviço','consVelServico.asp',0);"
	end if
	if blnTabelas then
		if right(strTabela,3) = "" then
			strTabela	 = "var aMenu5 = new Array('','510','95','','','','','',''," & strTabelaSub_1
			strTabelaSub_1 = ""
		else
			strTabela	 = "var aMenu5 = new Array('','510','95','','','','','',''," & strTabela
			if right(strTabela,3) <> ",0," then
				strTabelaSub = " var aMenu5_" & intTotalMenu & " = new Array("
			end if
		end if
	end if
	 strRelatorioSub  =  "var aMenu6 = new Array('','637','95','','','','','',''," + strRelatorios + ");" 

elseif blnAcessoLog and blnGerencia and not blnAcessoFis then

	strGerencia	 	= "var aMenu2 = new Array('','132','95','','','','','',''," & strGerencia
	strConsulta	 	= "var aMenu3 = new Array('','258','95','','','','','',''," & strConsulta
	if objDicCef.Exists("GAT") then
		strConsultaSub 	= " var aMenu4_22 = new Array('Acessos Físicos por Endereço','ConsultaAcessoFisicoEndereco.asp',0,'Histórico - Migração HFC BSoD','historico_migracaoBsod.asp',0,'Movimentação de Acessos (Lista)','consMovimentacao.asp',0,'Movimentação de Acessos Temporários (Lista)','consMovimentacaoTemp.asp',0,'Pendência de Cabo Interno','ConsPendCaboInterno.asp',0,'Pedidos por Ação','cons_Aca.asp',0,'Pedidos por Provedor e Ação','cons_ProAca.asp',0,'Pedidos por Provedor e Status','cons_ProStp.asp',0,'Pendentes de Instalação (Lista)','consPendInstalaLista.asp',0,'Pendentes de Instalação (Resumo)','consPendInstalaResumo.asp',0,'Percentual de Acessos com Serviços Ativados','consPercentAcesso.asp',0,'Preço-Provedor','javascript:window.open(""../Excel/PrecoProvedor/PrecoProvedor.xls"");',0,'Relatório de Backlogs Provedor','relapendconstacesso.asp',0,'Relatório Telco','RelaTelco.asp',0,'Qtde. de Pedidos por Provedor e Ação','cons_ProAcaQtd.asp',0,'Qtde. de Pedidos Aceitos por Provedor','cons_ProQtd.asp',0,'Qtde. de Pedidos Pendentes por Provedor','cons_ProData.asp',0,'Velocidade por Serviço','consVelServico.asp',0);"
	else
		strConsultaSub 	= " var aMenu4_22 = new Array('Acessos Físicos por Endereço','ConsultaAcessoFisicoEndereco.asp',0,'Histórico - Migração HFC BSoD','historico_migracaoBsod.asp',0,'Movimentação de Acessos (Lista)','consMovimentacao.asp',0,'Movimentação de Acessos Temporários (Lista)','consMovimentacaoTemp.asp',0,'Pendência de Cabo Interno','ConsPendCaboInterno.asp',0,'Pedidos por Ação','cons_Aca.asp',0,'Pedidos por Provedor e Ação','cons_ProAca.asp',0,'Pedidos por Provedor e Status','cons_ProStp.asp',0,'Pendentes de Instalação (Lista)','consPendInstalaLista.asp',0,'Pendentes de Instalação (Resumo)','consPendInstalaResumo.asp',0,'Percentual de Acessos com Serviços Ativados','consPercentAcesso.asp',0,'Preço-Provedor','javascript:window.open(""../Excel/PrecoProvedor/PrecoProvedor.xls"");',0,'Relatório de Backlogs Provedor','relapendconstacesso.asp',0,'Relatório Telco','RelaTelco.asp',0,'Qtde. de Pedidos por Provedor e Ação','cons_ProAcaQtd.asp',0,'Qtde. de Pedidos Aceitos por Provedor','cons_ProQtd.asp',0,'Qtde. de Pedidos Pendentes por Provedor','cons_ProData.asp',0,'Velocidade por Serviço','consVelServico.asp',0);"
	end if
	if blnTabelas then
		if right(strTabela,3) = "" then
			strTabela	 = "var aMenu4 = new Array('','385','95','','','','','',''," & strTabelaSub_1
			strTabelaSub_1 = ""
		else
			strTabela	 = "var aMenu4 = new Array('','385','95','','','','','',''," & strTabela
			if right(strTabela,3) <> ",0," then
				strTabelaSub = " var aMenu4_" & intTotalMenu & " = new Array("
			end if
		end if
	end if
	 strRelatorioSub  =  "var aMenu5 = new Array('','510','95','','','','','',''," + strRelatorios + ");" 

elseif blnAcessoLog and not blnGerencia and blnAcessoFis then
	strAcessoFis 	= "var aMenu2 = new Array('','132','95','','','','','',''," & strAcessoFis
	strAcessoFisSub  =  "var aMenu2_2 = new Array( "& strAcessoFisSub &");"
'	strAcessoFis = strAcessoFis & strAcessoFisSub  
	strConsulta	 	= "var aMenu3 = new Array('','258','95','','','','','',''," & strConsulta
	strConsultaSub 	= " var aMenu3_22 = new Array('Acessos Físicos por Endereço','ConsultaAcessoFisicoEndereco.asp',0,'Histórico - Migração HFC BSoD','historico_migracaoBsod.asp',0,'Movimentação de Acessos (Lista)','consMovimentacao.asp',0,'Movimentação de Acessos Temporários (Lista)','consMovimentacaoTemp.asp',0,'Pendência de Cabo Interno','ConsPendCaboInterno.asp',0,'Pedidos por Ação','cons_Aca.asp',0,'Pedidos por Provedor e Ação','cons_ProAca.asp',0,'Pedidos por Provedor e Status','cons_ProStp.asp',0,'Pendentes de Instalação (Lista)','consPendInstalaLista.asp',0,'Pendentes de Instalação (Resumo)','consPendInstalaResumo.asp',0,'Percentual de Acessos com Serviços Ativados','consPercentAcesso.asp',0,'Preço-Provedor','javascript:window.open(""../Excel/PrecoProvedor/PrecoProvedor.xls"");',0,'Relatório de Backlogs Provedor','relapendconstacesso.asp',0,'Relatório Telco','RelaTelco.asp',0,'Qtde. de Pedidos por Provedor e Ação','cons_ProAcaQtd.asp',0,'Qtde. de Pedidos Aceitos por Provedor','cons_ProQtd.asp',0,'Qtde. de Pedidos Pendentes por Provedor','cons_ProData.asp',0,'Velocidade por Serviço','consVelServico.asp',0);"
	if blnTabelas then
		if right(strTabela,3) = "" then
			strTabela	 = "var aMenu4 = new Array('','385','95','','','','','',''," & strTabelaSub_1
			strTabelaSub_1 = ""
		else
			strTabela	 = "var aMenu4 = new Array('','385','95','','','','','',''," & strTabela
			if right(strTabela,3) <> ",0," then
				strTabelaSub = " var aMenu4_" & intTotalMenu & " = new Array("
			end if
		end if
	end if
	strRelatorioSub  =  "var aMenu5 = new Array('','510','95','','','','','',''," + strRelatorios + ");" 
elseif not blnAcessoLog and blnGerencia and blnAcessoFis then

	strGerencia	 	= "var aMenu1 = new Array('','5','95','','','','','',''," & strGerencia
	strAcessoFis 		= "var aMenu2 = new Array('','132','95','','','','','',''," & strAcessoFis
	strAcessoFisSub  =  "var aMenu2_2 = new Array( "& strAcessoFisSub &");"
	strConsulta	 	= "var aMenu3 = new Array('','258','95','','','','','',''," & strConsulta
	strConsultaSub 	= "var aMenu3_22 = new Array('Acessos Físicos por Endereço','ConsultaAcessoFisicoEndereco.asp',0,'Histórico - Migração HFC BSoD','historico_migracaoBsod.asp',0,'Movimentação de Acessos (Lista)','consMovimentacao.asp',0,'Movimentação de Acessos Temporários (Lista)','consMovimentacaoTemp.asp',0,'Pendência de Cabo Interno','ConsPendCaboInterno.asp',0,'Pedidos por Ação','cons_Aca.asp',0,'Pedidos por Provedor e Ação','cons_ProAca.asp',0,'Pedidos por Provedor e Status','cons_ProStp.asp',0,'Pendentes de Instalação (Lista)','consPendInstalaLista.asp',0,'Pendentes de Instalação (Resumo)','consPendInstalaResumo.asp',0,'Percentual de Acessos com Serviços Ativados','consPercentAcesso.asp',0,'Preço-Provedor','javascript:window.open(""../Excel/PrecoProvedor/PrecoProvedor.xls"");',0,'Relatório de Backlogs Provedor','relapendconstacesso.asp',0,'Relatório Telco','RelaTelco.asp',0,'Qtde. de Pedidos por Provedor e Ação','cons_ProAcaQtd.asp',0,'Qtde. de Pedidos Aceitos por Provedor','cons_ProQtd.asp',0,'Qtde. de Pedidos Pendentes por Provedor','cons_ProData.asp',0,'Velocidade por Serviço','consVelServico.asp',0);"
	if blnTabelas then
		if right(strTabela,3) = "" then
			strTabela	 = "var aMenu4 = new Array('','385','95','','','','','',''," & strTabelaSub_1
			strTabelaSub_1 = ""
		else
			strTabela	 = "var aMenu4 = new Array('','385','95','','','','','',''," & strTabela
			if right(strTabela,3) <> ",0," then
				strTabelaSub = " var aMenu4_" & intTotalMenu & " = new Array("
			end if
		end if
	end if
	 strRelatorioSub  =  "var aMenu5 = new Array('','510','95','','','','','',''," + strRelatorios + ");" 

elseif not blnAcessoLog and not blnGerencia and blnAcessoFis then

	strAcessoFis 	= "var aMenu1 = new Array('','5','95','','','','','',''," & strAcessoFis
	strAcessoFisSub  =  "var aMenu2_2 = new Array( "& strAcessoFisSub &");"
	strConsulta	 	= "var aMenu2 = new Array('','132','95','','','','','',''," & strConsulta
	strConsultaSub 	= " var aMenu2_22 = new Array('Acessos Físicos por Endereço','ConsultaAcessoFisicoEndereco.asp',0,'Histórico - Migração HFC BSoD','historico_migracaoBsod.asp',0,'Movimentação de Acessos (Lista)','consMovimentacao.asp',0,'Movimentação de Acessos Temporários (Lista)','consMovimentacaoTemp.asp',0,'Pendência de Cabo Interno','ConsPendCaboInterno.asp',0,'Pedidos por Ação','cons_Aca.asp',0,'Pedidos por Provedor e Ação','cons_ProAca.asp',0,'Pedidos por Provedor e Status','cons_ProStp.asp',0,'Pendentes de Instalação (Lista)','consPendInstalaLista.asp',0,'Pendentes de Instalação (Resumo)','consPendInstalaResumo.asp',0,'Percentual de Acessos com Serviços Ativados','consPercentAcesso.asp',0,'Preço-Provedor','javascript:window.open(""../Excel/PrecoProvedor/PrecoProvedor.xls"");',0,'Relatório de Backlogs Provedor','relapendconstacesso.asp',0,'Relatório Telco','RelaTelco.asp',0,'Qtde. de Pedidos por Provedor e Ação','cons_ProAcaQtd.asp',0,'Qtde. de Pedidos Aceitos por Provedor','cons_ProQtd.asp',0,'Qtde. de Pedidos Pendentes por Provedor','cons_ProData.asp',0,'Velocidade por Serviço','consVelServico.asp',0);"
	if blnTabelas then
		if right(strTabela,3) = "" then
			strTabela	 = "var aMenu3 = new Array('','258','95','','','','','',''," & strTabelaSub_1
			strTabelaSub_1 = ""
		else
			strTabela	 = "var aMenu3 = new Array('','258','95','','','','','',''," & strTabela
			if right(strTabela,3) <> ",0," then
				strTabelaSub = " var aMenu3_" & intTotalMenu & " = new Array("
			end if
		end if
	end if
	strRelatorioSub  =  "var aMenu4 = new Array('','385','95','','','','','',''," + strRelatorios + ");" 

  
elseif not blnAcessoLog and blnGerencia and not blnAcessoFis then

	strGerencia	 	= "var aMenu1 = new Array('','5','95','','','','','',''," & strGerencia
	strConsulta	 	= "var aMenu2 = new Array('','132','95','','','','','',''," & strConsulta
	strConsultaSub 	= " var aMenu2_22 = new Array('Acessos Físicos por Endereço','ConsultaAcessoFisicoEndereco.asp',0,'Histórico - Migração HFC BSoD','historico_migracaoBsod.asp',0,'Movimentação de Acessos (Lista)','consMovimentacao.asp',0,'Movimentação de Acessos Temporários (Lista)','consMovimentacaoTemp.asp',0,'Pendência de Cabo Interno','ConsPendCaboInterno.asp',0,'Pedidos por Ação','cons_Aca.asp',0,'Pedidos por Provedor e Ação','cons_ProAca.asp',0,'Pedidos por Provedor e Status','cons_ProStp.asp',0,'Pendentes de Instalação (Lista)','consPendInstalaLista.asp',0,'Pendentes de Instalação (Resumo)','consPendInstalaResumo.asp',0,'Percentual de Acessos com Serviços Ativados','consPercentAcesso.asp',0,'Preço-Provedor','javascript:window.open(""../Excel/PrecoProvedor/PrecoProvedor.xls"");',0,'Relatório de Backlogs Provedor','relapendconstacesso.asp',0,'Relatório Telco','RelaTelco.asp',0,'Qtde. de Pedidos por Provedor e Ação','cons_ProAcaQtd.asp',0,'Qtde. de Pedidos Aceitos por Provedor','cons_ProQtd.asp',0,'Qtde. de Pedidos Pendentes por Provedor','cons_ProData.asp',0,'Velocidade por Serviço','consVelServico.asp',0);"

	if blnTabelas then
		if right(strTabela,3) = "" then
			strTabela	 = "var aMenu3 = new Array('','385','95','','','','','',''," & strTabelaSub_1
			strTabelaSub_1 = ""
		else
			strTabela	 = "var aMenu3 = new Array('','385','95','','','','','',''," & strTabela
			if right(strTabela,3) <> ",0," then
				strTabelaSub = " var aMenu3_" & intTotalMenu & " = new Array("
			end if
		end if
	end if
	 strRelatorioSub  =  "var aMenu4 = new Array('','510','95','','','','','',''," + strRelatorios + ");" 

elseif not blnAcessoLog and not blnGerencia and not blnAcessoFis then

	strConsulta	 	= "var aMenu1 = new Array('','5','95','','','','','',''," & strConsulta
	strConsultaSub 	= " var aMenu1_22 = new Array('Acessos Físicos por Endereço','ConsultaAcessoFisicoEndereco.asp',0,'Histórico - Migração HFC BSoD','historico_migracaoBsod.asp',0,'Movimentação de Acessos (Lista)','consMovimentacao.asp',0,'Movimentação de Acessos Temporários (Lista)','consMovimentacaoTemp.asp',0,'Pendência de Cabo Interno','ConsPendCaboInterno.asp',0,'Pedidos por Ação','cons_Aca.asp',0,'Pedidos por Provedor e Ação','cons_ProAca.asp',0,'Pedidos por Provedor e Status','cons_ProStp.asp',0,'Pendentes de Instalação (Lista)','consPendInstalaLista.asp',0,'Pendentes de Instalação (Resumo)','consPendInstalaResumo.asp',0,'Percentual de Acessos com Serviços Ativados','consPercentAcesso.asp',0,'Preço-Provedor','javascript:window.open(""../Excel/PrecoProvedor/PrecoProvedor.xls"");',0,'Relatório de Backlogs Provedor','relapendconstacesso.asp',0,'Relatório Telco','RelaTelco.asp',0,'Qtde. de Pedidos por Provedor e Ação','cons_ProAcaQtd.asp',0,'Qtde. de Pedidos Aceitos por Provedor','cons_ProQtd.asp',0,'Qtde. de Pedidos Pendentes por Provedor','cons_ProData.asp',0,'Velocidade por Serviço','consVelServico.asp',0);"
	if blnTabelas then
		if right(strTabela,3) = "" then
			strTabela	 = "var aMenu2 = new Array('','132','95','','','','','',''," & strTabelaSub_1
			strTabelaSub_1 = ""
		else
			strTabela	 = "var aMenu2 = new Array('','132','95','','','','','',''," & strTabela
			if right(strTabela,3) <> ",0," then
				strTabelaSub = " var aMenu2_" & intTotalMenu & " = new Array("
			end if
		end if
		strRelatorioSub  =  "var aMenu3 = new Array('','258','95','','','','','',''," + strRelatorios + ");" 
    else
		strRelatorioSub  =  "var aMenu2 = new Array('','132','95','','','','','',''," + strRelatorios + ");" 
		strTabela = ""
	end if

end if

'if blnGerencia then
	if strGerencia <> "" and right(strGerencia,3) = ",0," then
			strGerencia 	 = Left(strGerencia,len(strGerencia)-1) & ");" & Chr(13)
	end if
'end if

'if blnAcessoFis then
	if strAcessoFis <> "" and right(strAcessoFis,3) = ",0," then
			strAcessoFis 	 = Left(strAcessoFis,len(strAcessoFis)-1) & ");" & Chr(13)
	end if	
'end if

 if strAcessoFisSub <> "" and right(strAcessoFisSub ,3) = ",);" then
     	strAcessoFisSub 	 = Left(strAcessoFisSub ,len(strAcessoFisSub )-3) & ");" & Chr(13)
 end if


'if blnATabelas then
	if strTabela <> "" and strTabelaSub_1 = "" and right(strTabela,3) = ",0," then
		strTabela 	 = Left(strTabela,len(strTabela)-1) & ");" & Chr(13)
	end if
	if strTabelaSub <> "" and right(strTabelaSub,3) = ",0," then
		strTabelaSub 	 = Left(strTabelaSub,len(strTabelaSub)-1) & ");" & Chr(13)
	end if
	if strTabelaSub_1 <> "" and right(strTabelaSub_1,3) = ",0," then
	strTabelaSub_1 	 = Left(strTabelaSub_1,len(strTabelaSub_1)-1) & ");" & Chr(13)
	end if
'end if

	'''strConsultaSub = ""
	
  strGeral = strGeral & strGerencia & strAcessoFis & strConsulta & strConsultaSub & strTabela & strTabelaSub & strTabelaSub_1 & strRelatorioSub

	''' strGeral = strGeral & strGerencia & strAcessoFis & strConsulta & strTabela & strTabelaSub & strTabelaSub_1 
 


objAryPag = split(Request.ServerVariables("SCRIPT_NAME"),"/")
objAryPagAux = split(Ucase(objAryPag(Ubound(objAryPag))),"?")
strPagina = objAryPagAux(0)
'Response.Write "<script language=javascript >alert('"&strPagina&"');</script>"
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
	alert("ddd");
}
</Script>
<SCRIPT language='javascript'>
  javascript:window.history.forward(1);
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" onLoad="resposta(<%if DBAction <> "" then response.write DBAction else response.write "0" end if%>,'main.asp');">
<table width="790" border="0" cellspacing="0" cellpadding="0">
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
		<div id=divMenu style="position: absolute; top: 81; left: 5; width: 760; height: 5">
			<table width="790" border="0" cellspacing="0" cellpadding="0" >
			<%
			if blnAcessoLog and blnGerencia and blnAcessoFis then
			%>
			<tr>
				<td nowrap width=<%=760/6%>>
				<span id=spnAcessoLog onMouseOut="popDown('eMenu1')" onClick="showInput(false);popUp('eMenu1',event)" style='cursor:hand' ><font class=clsMenu >Acesso Lógico</font></span>
				</td>
				<!-- ''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
				<td nowrap width=<%=760/6%>>
				<span id=spnGerencia onMouseOut="popDown('eMenu2')" onClick="showInput(false);popUp('eMenu2',event)" style='cursor:hand' ><font class=clsMenu >|&nbsp;Tarefas</font></span>
				</td>
				''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
				-->
				<td nowrap width=<%=760/6%> >
					<span id=spnAcessoFis  onMouseOut="popDown('eMenu3')" onClick="showInput(false);popUp('eMenu3',event)" style='cursor:hand' ><font class=clsMenu >|&nbsp;Acesso Físico</font></span>
				</td>
				<td nowrap width=<%=760/6%>>
					<span id=spnConsultas onMouseOut="popDown('eMenu4')" onClick="showInput(false);popUp('eMenu4',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Consultas</font></span>
				</td>
				<td nowrap width=<%=760/6%>>
					<span id=spnTabelas onMouseOut="popDown('eMenu5')" onClick="showInput(false);popUp('eMenu5',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Tabelas</font></span>
				</td>
				<!-- ''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
				<td nowrap >
				<span id=spnRelatorios onMouseOut="popDown('eMenu6')" onClick="showInput(false);popUp('eMenu6',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Relatorios</font></span>
				</td>
				<td nowrap >
					<span id=spnCrms style='cursor:hand' onClick="window.open('http://<%=Request.ServerVariables("SERVER_NAME")%>/crmsf/inicio.htm','CRMSF','')" onmouseover="showtip(this,event,'Controle de Radio,Mux,Satélite e Fibra');" onmouseout="hidetip();"><font class=clsMenu >|&nbsp;CRMSF</font></span>
				</td>
				''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
				-->

			</tr>
			<%
			elseif blnAcessoLog and blnGerencia and not blnAcessoFis then
			%>
			<tr>
				<td nowrap width=<%=760/6%>>
				<span id=spnAcessoLog onMouseOut="popDown('eMenu1')" onClick="showInput(false);popUp('eMenu1',event)" style='cursor:hand' ><font class=clsMenu >Acesso Lógico</font></span>
				</td>
				<!-- ''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
				<td nowrap width=<%=760/6%>>
				<span id=spnGerencia onMouseOut="popDown('eMenu2')" onClick="showInput(false);popUp('eMenu2',event)" style='cursor:hand' ><font class=clsMenu >|&nbsp;Tarefas</font></span>
				</td>
				''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
				-->
				<td nowrap width=<%=760/6%>>
					<span id=spnConsultas onMouseOut="popDown('eMenu3')" onClick="showInput(false);popUp('eMenu3',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Consultas</font></span>
				</td>
				<td nowrap width=<%=760/6%>>
					<span id=spnTabelas onMouseOut="popDown('eMenu4')" onClick="showInput(false);popUp('eMenu4',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Tabelas</font></span>
				</td>
				<!-- ''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
				<td nowrap >
				<span id=spnRelatorios onMouseOut="popDown('eMenu5')" onClick="showInput(false);popUp('eMenu5',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Relatorios</font></span>
				</td>

				<td nowrap >
					<span id=spnCrms style='cursor:hand' onClick="window.open('http://<%=Request.ServerVariables("SERVER_NAME")%>/crmsf/inicio.htm','CRMSF','')" onmouseover="showtip(this,event,'Controle de Radio,Mux,Satélite e Fibra');" onmouseout="hidetip();"><font class=clsMenu >|&nbsp;CRMSF</font></span>
				</td>
				''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
				-->
			</tr>
			<%
			elseif blnAcessoLog and not blnGerencia and blnAcessoFis then
			%>
			<tr>
				<td nowrap width=<%=760/6%>>
				<span id=spnAcessoLog onMouseOut="popDown('eMenu1')" onClick="showInput(false);popUp('eMenu1',event)" style='cursor:hand' ><font class=clsMenu >Acesso Lógico</font></span>
				</td>
				<td nowrap width=<%=760/6%> >
					<span id=spnAcessoFis  onMouseOut="popDown('eMenu2')" onClick="showInput(false);popUp('eMenu2',event)" style='cursor:hand' ><font class=clsMenu >|&nbsp;Acesso Físico</font></span>
				</td>
				<td nowrap width=<%=760/6%>>
					<span id=spnConsultas onMouseOut="popDown('eMenu3')" onClick="showInput(false);popUp('eMenu3',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Consultas</font></span>
				</td>
				<%if blnTabelas then %>
				<td nowrap width=<%=760/6%>>
					<span id=spnTabelas onMouseOut="popDown('eMenu4')" onClick="showInput(false);popUp('eMenu4',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Tabelas</font></span>
				</td>
				<%End if%>
				<!-- ''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
				<td nowrap >
				<span id=spnRelatorios onMouseOut="popDown('eMenu5')" onClick="showInput(false);popUp('eMenu5',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Relatorios</font></span>
				</td>
				<td nowrap >
					<span id=spnCrms style='cursor:hand' onClick="window.open('http://<%=Request.ServerVariables("SERVER_NAME")%>/crmsf/inicio.htm','CRMSF','')" onmouseover="showtip(this,event,'Controle de Radio,Mux,Satélite e Fibra');" onmouseout="hidetip();"><font class=clsMenu >|&nbsp;CRMSF</font></span>
				</td>
				''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
				-->
			</tr>


			<%
			elseif not blnAcessoLog and blnGerencia and blnAcessoFis then
			%>
			<tr>
				<!-- ''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
				<td nowrap width=<%=760/6%>>
				<span id=spnGerencia onMouseOut="popDown('eMenu1')" onClick="showInput(false);popUp('eMenu1',event)" style='cursor:hand' ><font class=clsMenu >|&nbsp;Tarefas</font></span>
				</td>
				''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
				-->
				<td nowrap width=<%=760/6%> >
					<span id=spnAcessoFis  onMouseOut="popDown('eMenu2')" onClick="showInput(false);popUp('eMenu2',event)" style='cursor:hand' ><font class=clsMenu >|&nbsp;Acesso Físico</font></span>
				</td>
				<td nowrap width=<%=760/6%>>
					<span id=spnConsultas onMouseOut="popDown('eMenu3')" onClick="showInput(false);popUp('eMenu3',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Consultas</font></span>
				</td>
				<%if blnTabelas then %>
				<td nowrap width=<%=760/6%>>
					<span id=spnTabelas onMouseOut="popDown('eMenu4')" onClick="showInput(false);popUp('eMenu4',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Tabelas</font></span>
				</td>
				<%End if%>
				<!-- ''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
				<td nowrap >
				<span id=spnRelatorios onMouseOut="popDown('eMenu5')" onClick="showInput(false);popUp('eMenu5',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Relatorios</font></span>
				</td>

				<td nowrap >
					<span id=spnCrms style='cursor:hand' onClick="window.open('http://<%=Request.ServerVariables("SERVER_NAME")%>/crmsf/inicio.htm','CRMSF','')" onmouseover="showtip(this,event,'Controle de Radio,Mux,Satélite e Fibra');" onmouseout="hidetip();"><font class=clsMenu >|&nbsp;CRMSF</font></span>
				</td>
				''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
				-->
			</tr>
			<%
			elseif not blnAcessoLog and blnGerencia and not blnAcessoFis then
			%>
			<tr>
				<!-- ''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
				<td nowrap width=<%=760/6%>>
				<span id=spnGerencia onMouseOut="popDown('eMenu1')" onClick="showInput(false);popUp('eMenu1',event)" style='cursor:hand' ><font class=clsMenu >|&nbsp;Tarefas</font></span>
				</td>
				''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
				-->
				<td nowrap width=<%=760/6%>>
					<span id=spnConsultas onMouseOut="popDown('eMenu2')" onClick="showInput(false);popUp('eMenu2',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Consultas</font></span>
				</td>
				<%if blnTabelas then %>
				<td nowrap width=<%=760/6%>>
					<span id=spnTabelas onMouseOut="popDown('eMenu3')" onClick="showInput(false);popUp('eMenu3',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Tabelas</font></span>
				</td>
				<%End if%>
				<!-- ''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
				<td nowrap >
				<span id=spnRelatorios onMouseOut="popDown('eMenu4')" onClick="showInput(false);popUp('eMenu4',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Relatorios</font></span>
				</td>

				''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
				-->
				<td nowrap >
					<span id=spnCrms style='cursor:hand' onClick="window.open('http://<%=Request.ServerVariables("SERVER_NAME")%>/crmsf/inicio.htm','CRMSF','')" onmouseover="showtip(this,event,'Controle de Radio,Mux,Satélite e Fibra');" onmouseout="hidetip();"><font class=clsMenu >|&nbsp;CRMSF</font></span>
				</td>
				''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
				-->
		
			</tr>
			<%
			elseif not blnAcessoLog and not blnGerencia and blnAcessoFis then
            %>
			<tr>
				<td nowrap width=<%=760/6%> >
					<span id=spnAcessoFis  onMouseOut="popDown('eMenu1')" onClick="showInput(false);popUp('eMenu1',event)" style='cursor:hand' ><font class=clsMenu >|&nbsp;Acesso Físico</font></span>
				</td>
				<td nowrap width=<%=760/6%>>
					<span id=spnConsultas onMouseOut="popDown('eMenu2')" onClick="showInput(false);popUp('eMenu2',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Consultas</font></span>
				</td>
				<%if blnTabelas then %>
				<td nowrap width=<%=760/6%>>
					<span id=spnTabelas onMouseOut="popDown('eMenu3')" onClick="showInput(false);popUp('eMenu3',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Tabelas</font></span>
				</td>
				<%End if%>
				<!-- ''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
				<td nowrap >
				<span id=spnRelatorios onMouseOut="popDown('eMenu4')" onClick="showInput(false);popUp('eMenu4',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Relatorios</font></span>
				</td>
				''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
				-->
				<td nowrap >
					<span id=spnCrms style='cursor:hand' onClick="window.open('http://<%=Request.ServerVariables("SERVER_NAME")%>/crmsf/inicio.htm','CRMSF','')" onmouseover="showtip(this,event,'Controle de Radio,Mux,Satélite e Fibra');" onmouseout="hidetip();"><font class=clsMenu >|&nbsp;CRMSF</font></span>
				</td>
				''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
				-->
			</tr>
			<%elseif not blnAcessoLog and not blnGerencia and not blnAcessoFis then%>
			<tr>
				<td nowrap width=<%=760/6%>>
					<span id=spnConsultas onMouseOut="popDown('eMenu1')" onClick="showInput(false);popUp('eMenu1',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Consultas</font></span>
				</td>
				<%if blnTabelas then %>
				<td nowrap width=<%=760/6%>>
					<span id=spnTabelas onMouseOut="popDown('eMenu2')" onClick="showInput(false);popUp('eMenu2',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Tabelas</font></span>
				</td>
				<!-- ''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
				<td nowrap >
				<span id=spnRelatorios onMouseOut="popDown('eMenu3')" onClick="showInput(false);popUp('eMenu3',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Relatorios</font></span>
				</td>
				''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
				-->
				<%End if%>
			   
			   <!-- ''' INICIO RETIRADA PROJETO REESTRUTURAÇÃO CLA
				<td nowrap >
				<span id=spnRelatorios onMouseOut="popDown('eMenu2')" onClick="showInput(false);popUp('eMenu2',event)" style='cursor:hand'><font class=clsMenu >|&nbsp;Relatorios</font></span>
				</td>			   
				
				<td nowrap >
					<span id=spnCrms style='cursor:hand' onClick="window.open('http://<%=Request.ServerVariables("SERVER_NAME")%>/crmsf/inicio.htm','CRMSF','')" onmouseover="showtip(this,event,'Controle de Radio,Mux,Satélite e Fibra');" onmouseout="hidetip();"><font class=clsMenu >|&nbsp;CRMSF</font></span>
				</td>
					''' FIM RETIRADA PROJETO REESTRUTURAÇÃO CLA
				-->
			</tr>
			<%end if%>
			</table>
		</div>
	</td>
</tr>
</table>
<table valign="top" width="790" cellspacing="2" cellpadding="0">
<input type=hidden name=hdnUserHerder value="<%=strUserName%>">
<input type=hidden name=hdnUserHerderPerfil value="<%=strUserHerderPerfil%>">