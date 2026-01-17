<%
Response.ContentType = "text/html; charset=utf-8"
Response.Charset = "UTF-8"
%>
<!--#include file="../inc/log.asp"-->
<%
on error resume next

Dim objDicCef, objDicPag
Dim strGeral, blnAcessoLog, blnAcessoFis, blnTabelas, blnGerencia
Dim strAcessoFis, strGerencia, strConsulta, strTabela
Dim strUserHerderPerfil

Set objRS = db.execute("CLA_sp_sel_usuarioctfc2 null," & dblUsuId)
Set objDicCef = Server.CreateObject("Scripting.Dictionary")
Set objDicPag = Server.CreateObject("Scripting.Dictionary")

While Not objRS.Eof
    if Not objDicCef.Exists(Ucase(Trim(objRS("Age_Sigla")))) then
        Call objDicCef.Add (Ucase(Trim(objRS("Age_Sigla"))),Ucase(Trim(objRS("Ctfc_ID"))))
    End if
    objRS.MoveNext
Wend

strGeral = ""
strAcessoFis = ""
strGerencia = ""
strConsulta = ""
strTabela = ""
blnAcessoLog = false
blnGerencia = false
blnAcessoFis = false
blnTabelas = false

' Simulacao para teste - REMOVER em producao
blnAcessoLog = true
blnAcessoFis = true
blnTabelas = true

Dim objAryPag, objAryPagAux, strPagina
objAryPag = split(Request.ServerVariables("SCRIPT_NAME"),"/")
objAryPagAux = split(Ucase(objAryPag(Ubound(objAryPag))),"?")
strPagina = objAryPagAux(0)

Set objDicPag = Nothing
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <title>CLA - Controle Local de Acesso</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Arial, Helvetica, sans-serif; background-color: #f5f5f5; min-width: 320px; }
        .header { background-color: #003366; position: relative; padding: 15px 20px; }
        .header-content { max-width: 1200px; margin: 0 auto; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; }
        .logo { height: 60px; cursor: pointer; }
        .system-info { color: white; font-size: 12px; text-align: right; }
        .system-info div { margin: 3px 0; }
        .nav-container { background-color: #f1f1f1; border-bottom: 3px solid #003366; position: sticky; top: 0; z-index: 999; }
        .nav-wrapper { max-width: 1200px; margin: 0 auto; position: relative; }
        .nav-menu { display: flex; list-style: none; margin: 0; padding: 0; }
        .nav-item { position: relative; }
        .nav-link { display: block; padding: 16px 24px; color: #003366; text-decoration: none; font-weight: bold; font-size: 14px; white-space: nowrap; background-color: transparent; border: none; cursor: pointer; transition: all 0.2s; border-right: 1px solid #d0d0d0; }
        .nav-link:hover, .nav-link.active { background-color: #003366; color: white; }
        .dropdown-menu { visibility: hidden; opacity: 0; position: absolute; top: 100%; left: 0; background-color: white; min-width: 300px; box-shadow: 0 8px 16px rgba(0,0,0,0.2); border: 1px solid #ccc; border-top: 3px solid #003366; max-height: 450px; overflow-y: auto; z-index: 1000; transition: opacity 0.2s, visibility 0.2s; }
        .dropdown-menu.active { visibility: visible; opacity: 1; }
        .dropdown-item { display: block; padding: 12px 20px; color: #333; text-decoration: none; border-bottom: 1px solid #f0f0f0; font-size: 13px; transition: background-color 0.2s; }
        .dropdown-item:last-child { border-bottom: none; }
        .dropdown-item:hover { background-color: #e8f0f8; color: #003366; }
        .mobile-menu-toggle { display: none; background-color: white; border: 2px solid white; color: #003366; font-size: 24px; cursor: pointer; padding: 8px 15px; border-radius: 4px; }
        .content-wrapper { max-width: 1200px; margin: 20px auto; padding: 20px; background: white; min-height: 400px; border-radius: 4px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        /* Estilos para formularios e tabelas do sistema */
        .text, input[type="text"], select { padding: 6px 10px; border: 1px solid #ccc; border-radius: 3px; font-size: 13px; }
        .button, input[type="button"], input[type="submit"] { padding: 8px 16px; background-color: #003366; color: white; border: none; border-radius: 3px; cursor: pointer; font-size: 13px; }
        .button:hover, input[type="button"]:hover, input[type="submit"]:hover { background-color: #004488; }
        table { border-collapse: collapse; width: 100%; }
        th { background-color: #003366; color: white; padding: 10px 8px; text-align: left; font-size: 12px; }
        td { padding: 8px; font-size: 12px; border-bottom: 1px solid #ddd; }
        .clsSilver { background-color: #f9f9f9; }
        .clsSilver2 { background-color: #ffffff; }
        .clsSilver td, .clsSilver2 td { vertical-align: middle; }
        a { color: #003366; }
        a:hover { color: #0055aa; }
        @media (max-width: 968px) {
            .header-content { flex-direction: column; gap: 15px; }
            .logo { height: 50px; }
            .system-info { text-align: center; font-size: 11px; width: 100%; }
            .mobile-menu-toggle { display: block; position: absolute; top: 15px; right: 20px; }
            .nav-menu { flex-direction: column; width: 100%; max-height: 0; overflow: hidden; transition: max-height 0.3s ease-out; }
            .nav-menu.mobile-open { max-height: 2000px; }
            .nav-item { width: 100%; border-bottom: 1px solid #ddd; }
            .nav-link { width: 100%; border-right: none; text-align: left; }
            .dropdown-menu { position: static; box-shadow: none; border: none; border-left: 3px solid #003366; background-color: #f9f9f9; max-height: none; }
            .dropdown-item { padding-left: 40px; }
            .content-wrapper { margin: 10px; padding: 15px; }
            table { font-size: 11px; }
            th, td { padding: 6px 4px; }
        }
    </style>
</head>
<body>
    <header class="header">
        <div class="header-content">
            <img src="../imagens/topo_embratel.jpg" alt="Embratel" class="logo" onclick="window.location.href='main.asp'">
            <button class="mobile-menu-toggle" id="mobileMenuBtn">&#9776;</button>
            <div class="system-info">
                <div><strong>Servidor:</strong> <%=Request.ServerVariables("SERVER_NAME")%></div>
                <div><strong>Banco:</strong> <%=strBanco%></div>
                <div><strong>Usuario:</strong> <%=strUserName%></div>
            </div>
        </div>
    </header>

    <nav class="nav-container">
        <div class="nav-wrapper">
            <ul class="nav-menu" id="mainMenu">
                <%If blnAcessoLog Then%>
                <li class="nav-item">
                    <a href="javascript:void(0);" class="nav-link" onclick="toggleDropdown('drop1')">Acesso Logico</a>
                    <div class="dropdown-menu" id="drop1">
                        <a href="SolicitacaoAprov_main.asp" class="dropdown-item">Solicitacao de Acesso - Aprovisionador</a>
                        <a href="Solicitacao.asp" class="dropdown-item">Solicitacao de Acesso - Ativacao</a>
                        <!--<a href="AcessoLogico.asp?acao=alt" class="dropdown-item">Solicitacao de Acesso - Alteracao</a>
                        <a href="AcessoLogico.asp?acao=can" class="dropdown-item">Solicitacao de Acesso - Cancelamento</a>
                        <a href="AcessoLogico.asp?acao=des" class="dropdown-item">Solicitacao de Acesso - Desativacao</a>
                        <a href="SolicitacaoReenvioAprov_main.asp" class="dropdown-item">Solicitacao de Reenvio - Aprovisionador</a>
                        -->
						<a href="monitoracaoEntregaAtivacao.asp" class="dropdown-item">Entregar/Ativar servico</a>
                    </div>
                </li>
                <%End If%>
                <%If blnAcessoFis Then%>
                <li class="nav-item">
                    <a href="javascript:void(0);" class="nav-link" onclick="toggleDropdown('drop2')">Acesso Fisico</a>
                    <div class="dropdown-menu" id="drop2">
                        <a href="AlocacaoNew_main.ASP" class="dropdown-item">Alocacao / Aceite</a>
						<!--
                        <a href="ManutencaoAlocacaoNew_main.ASP" class="dropdown-item">Alteracao de Facilidades</a>
                        -->
						<a href="DesalocacaoNew_main.ASP" class="dropdown-item">Desalocacao</a>
                        <a href="SolicPedSNOA_main.asp" class="dropdown-item">Solicitacao de Pedido SNOA</a>
                        <a href="TratarProvidenciaSNOA_main.asp" class="dropdown-item">Tratar Providencia SNOA</a>
                        <a href="CartaProvedor.asp" class="dropdown-item">Criacao/Envio de Carta ao Provedor</a>
                        <a href="CadastraOSProvedor.asp" class="dropdown-item">Pendente OS Provedor</a>
                    </div>
                </li>
                <%End If%>
                <li class="nav-item">
                    <a href="javascript:void(0);" class="nav-link" onclick="toggleDropdown('drop3')">Consultas</a>
                    <div class="dropdown-menu" id="drop3">
                        <a href="CartasProvedor.asp" class="dropdown-item">Cartas ao Provedor</a>
                        <a href="consulta_main.asp" class="dropdown-item">Geral</a>
                        <a href="historico_main.asp" class="dropdown-item">Historico de Pedido</a>
                        <a href="historicopedidosnoa_main.asp" class="dropdown-item">Historico de Pedido SNOA</a>
                        <!--<a href="cons_historicofacilidade.asp" class="dropdown-item">Historico de Facilidades</a> -->
                        <a href="RelatorioAcessoFisico_main.asp" class="dropdown-item">Relatorio de Pesquisa Acesso Fisico</a>
                       <!-- <a href="historico_migracaoBsod.asp" class="dropdown-item">Historico - Migracao HFC BSoD</a> -->
                    </div>
                </li>
                <%If blnTabelas Then%>
                <li class="nav-item">
                    <a href="javascript:void(0);" class="nav-link" onclick="toggleDropdown('drop4')">Tabelas</a>
                    <div class="dropdown-menu" id="drop4">
                        <a href="Assoc_estacaoCtf_main.asp" class="dropdown-item">Consulta Estacao/Associacao Estacao CF</a>
                        <a href="AssocUserCef_main.asp" class="dropdown-item">Associacao Usuario com CF</a>
                        <a href="AssocServMod_main.asp" class="dropdown-item">Associacao de Servico com Modalidade</a>
                        <a href="AssocServVel_main.asp" class="dropdown-item">Associacao de Servico com Velocidade</a>
                        <a href="AssocTecVel_main.asp" class="dropdown-item">Associacao de Tecnologia com Velocidade</a>
                        <a href="CentroFuncional_main.asp" class="dropdown-item">Centro Funcional</a>
                        <a href="estacao_main.asp" class="dropdown-item">Estacao</a>
						<!--
                        <a href="CadastrarFacilidade.asp" class="dropdown-item">Facilidade</a>
                        <a href="ManutencaoFacilidade.asp" class="dropdown-item">Manutencao de Facilidade</a-->
						
						<a href="tecnologia_main.asp" class="dropdown-item">Tecnologia</a>
						<a href="assocTecnologiaFacilidade_main.asp" class="dropdown-item">Associacao Tecnologia com Facilidade</a>
                        <a href="combo_main.asp" class="dropdown-item">Combo</a>
						<a href="estruturaCombo_main.asp" class="dropdown-item">Estrutura Combo</a>
                        <a href="estruturaTecnologia_main.asp" class="dropdown-item">Estrutura Tecnologia</a>
						
                        <a href="Holding_main.asp" class="dropdown-item">Holding</a>
                        <a href="LocalidadeSubCef_main.asp" class="dropdown-item">Localidade Subordinada ao CF</a>
                        <a href="provedor_main.asp" class="dropdown-item">Provedor</a>
                        <a href="provedorcontrato_main.asp" class="dropdown-item">Provedor Contrato</a>
                        
                        <a href="recurso_main.asp" class="dropdown-item">Recurso</a>
                        <a href="regimecontrato_main.asp" class="dropdown-item">Regime de Contrato</a>
                        <a href="redeInterna.asp" class="dropdown-item">Rede Interna</a>
                        <a href="servico_main.asp" class="dropdown-item">Servico</a>
                        <a href="status_main.asp" class="dropdown-item">Status</a>
                        <a href="tipocontrato_main.asp" class="dropdown-item">Tipo de Vigencia de Contrato</a>
                        <a href="usuario_main.asp" class="dropdown-item">Usuario</a>
                        <a href="velocidade_main.asp" class="dropdown-item">Velocidade</a>
                    </div>
                </li>
                <%End If%>
            </ul>
        </div>
    </nav>

    <div class="content-wrapper">
        <input type="hidden" name="hdnUserHerder" value="<%=strUserName%>">
        <input type="hidden" name="hdnUserHerderPerfil" value="<%=strUserHerderPerfil%>">
        <!-- CONTEUDO DA PAGINA COMECA AQUI -->