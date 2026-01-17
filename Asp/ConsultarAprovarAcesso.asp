<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%
'•ACCENTURE
'	- Sistema			: CLA
'	- Arquivo			: ConsultarAutorizarAcesso.asp
'	- Responsável		: Gustavo S. Reynaldo
'	- Descrição			: Tela de consulta da autorização de acesso a terceiros
%>
<html>
	<head>
		<script language='javascript' src="../javascript/xmlFacObjects.js"></script>
		<SCRIPT LANGUAGE=javascript>
		<!--
		var objXmlGeral = new ActiveXObject("Microsoft.XMLDOM")

		//Quando o campo Número de Senha está vazio completa com PIN
		function CompletarSenha()
		{
			if(document.getElementById('txtSenhaPIN').value == "")
			{
				document.getElementById('txtSenhaPIN').value = 'PIN';
			}
			else
			{
				document.getElementById('btnProcurar').disabled = false;
			}
		}

		//Quando o campo Número de Pedido está vazio completa com DM-
		function HabilitaProcurarSenha(obj)
		{
			return
		}

		function VazioSenha()
		{
			if((document.getElementById('txtSenhaPIN').value == "PIN" || document.getElementById('txtSenhaPIN').value == "") && document.getElementById('cboSistemaOrderEntry').value  == 0
			&& document.getElementById('txtAno').value  == "" && document.getElementById('txtNumero').value  == "" 
			&& document.getElementById('txtItem').value  == ""	&& document.getElementById('txtSolicitacao').value  == "" 
			&& document.getElementById('txtAcessoFisico').value  == "")
			{
				return true;
			}
			
			return false;
		}


		//Executa a busca de senha
		function ProcurarSenha(){
			
			with (document.forms[0])
			{
			
				document.getElementById("IFrmLista").height = 370
				document.getElementById("BarraAzul").style.height = 0
				document.getElementById("BarraAzul").style.visibility = "hidden";
					
				//if (!ValidarDM(txtPedido)) return;
				target = "IFrmLista"
				action = "ListaSenha.asp"
				submit()
			}
		}

		function onseleciona_1() 
		{ 
		if (document.getElementById("cboTpAprovacao").value == "2")
		  { 
		  	with (document.forms[0])
			{
			var strNome = "Autorização de Acessos Terceiros"			
			//var objJanela = window.open()
			//objJanela.name =  strNome
			target = "_self"
			action = "IncluirAutorizarAcesso.asp?tp=2"
			submit()
			}
		  }
		}

		//-->
		</SCRIPT>
	</head>
	<body>
		<form method="post" name=Form1 >
			<input type=hidden name="hdnAcao">
			<table border=0 cellspacing="1" cellpadding = 0 width="760" >
				<tr class=clsSilver>
					<th colspan=2><p align=center>Aprovação de Senhas</p></th>
				</tr>
				<tr>
					<td class=clsSilver colspan=2>&nbsp;</td>
				</tr>
				<tr class=clsSilver>
					<th class=clsSilver colspan=2><p>Tipo de Aprovação:&nbsp;
						<select size="1" ID="cboTpAprovacao" name="cboTpAprovacao" tabindex="1" onChange="onseleciona_1()">
							<option value="1" <%if strCboTpAprovacao = 1 then%>selected<%end if%>>Circuito Terceiro</option>
							<option value="2" <%if strCboTpAprovacao = 2 then%>selected<%end if%>>Projeto Especial de Acesso</option>
						</select></p> 
					</th>
				</tr>
				<tr>
					<td class=clsSilver colspan=2>&nbsp;</td>
				</tr>
				<tr class=clsSilver>
					<td>&nbsp;&nbsp;&nbsp;Senha PIN:</td>
					<td>
						<input type="text" class="text" name="txtSenhaPIN" onKeyUP="ValidarTipo(this,2);HabilitaProcurarSenha(this)" onChange="CompletarSenha()" value="<%if request("txtSenhaPIN") <> "" then response.write ucase(request("txtSenhaPIN")) else response.write "PIN" end if%>" maxlength="11" size="12">
					</td>
				</tr>	
				<tr class=clsSilver>
					<td>&nbsp;&nbsp;&nbsp;OE Sistema:</td>
					<td>
						<select name="cboSistemaOrderEntry" onChange="HabilitaProcurarSenha(this)" <%if strOrigem="APG" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
							<Option value="">Todos</Option>
							<Option value="APG"					<%if strOrderEntrySis = "APG" then Response.Write " selected " End If%>>APG</Option>
							<Option value="CFD"					<%if strOrderEntrySis = "CFD" then Response.Write " selected " End If%>>CFD</Option>
							<Option value="SGA VOZ 0300"		<%if strOrderEntrySis = "SGA VOZ 0300" then Response.Write " selected " End If%>>SGA VOZ 0300</Option>
							<Option value="SGA VOZ 0800 FASE 1"	<%if strOrderEntrySis = "SGA VOZ 0800 FASE 1" then Response.Write " selected " End If%>>SGA VOZ 0800 FASE 1</Option>
							<Option value="SGA VOZ VIP'S"		<%if strOrderEntrySis = "SGA VOZ VIP'S" then Response.Write " selected " End If%>>SGA VOZ VIP'S</Option>
							<Option value="SGA DADOS"			<%if strOrderEntrySis = "SGA DADOS" then Response.Write " selected " End If%>>SGA DADOS</Option>
							<Option value="SGA PLUS"			<%if strOrderEntrySis = "SGA PLUS" then Response.Write " selected " End If%>>SGA PLUS</Option>
							<Option value="ADFAC"				<%if strOrderEntrySis = "ADFAC" then Response.Write " selected " End If%>>ADFAC</Option>
							<Option value="CFM"					<%if strOrderEntrySis = "CFM" then Response.Write " selected " End If%>>CFM</Option>
							<Option value="CFT"					<%if strOrderEntrySis = "CFT" then Response.Write " selected " End If%>>CFT</Option>
						</Select>
						&nbsp;&nbsp;Ano:
						<input type="text" class="text" name="txtAno" value="" style="visibility: visible;" TIPO="N" onBlur="CompletarCampo(this)" onKeyUp="ValidarTipo(this,0);HabilitaProcurarSenha(this)" maxlength="4" size="4">
						&nbsp;&nbsp;Número:
						<input type="text" class="text" name="txtNumero" value="" style="visibility: visible;" TIPO="N" onBlur="CompletarCampo(this)" onKeyUp="ValidarTipo(this,0);HabilitaProcurarSenha(this)" maxlength="7" size="7">
						&nbsp;&nbsp;Item:
						<input type="text" class="text" name="txtItem" value="" style="visibility: visible;" TIPO="N" onBlur="CompletarCampo(this)" onKeyUp="ValidarTipo(this,0);HabilitaProcurarSenha(this)" maxlength="3" size="4">
					</td>
				</tr>
				<tr class=clsSilver>
					<td>&nbsp;&nbsp;&nbsp;N° Solicitação:</td>
					<td>
						<input type="text" class="text" name="txtSolicitacao" onKeyUp="ValidarTipo(this,0);HabilitaProcurarSenha(this)" maxlength="9" size="11">
					</td>
				</tr>
				<tr class=clsSilver>
					<td>&nbsp;&nbsp;&nbsp;Acesso Fisico:</td>
					<td>
						<input type="text" class="text" name="txtAcessoFisico" onKeyUp="ValidarTipo(this,2);HabilitaProcurarSenha(this)" maxlength="15" size="18">
						&nbsp;&nbsp;&nbsp;&nbsp;
						<input type="button" class="button" name="btnProcurar" value="Procurar" accesskey="P" onMouseOver="showtip(this,event,'Procurar (Alt+P)');" style="width:100px" onClick="ProcurarSenha()">
					</td>
				</tr>
				<tr class=clsSilver>
					<th id="BarraAzul" height=20 colspan=2></th>
				</tr>
			</table>
			<br>
			<table border="0" cellspacing="1" cellpadding="0" width="760">
				<tr>
					<td>
						<iframe	id			= "IFrmLista"
								name        = "IFrmLista" 
								width       = "100%"
								height      = "0"
								frameborder = "0"
								border		= "0"
								scrolling   = "no">
						</iFrame>
					</td>
				</tr>
				<tr>
					<td align="center" height=30px>
						<input type="button" class="button" name="btnSair" value=" Sair " accesskey="B" onmouseover="showtip(this,event,'Sair (Alt+B)');" onClick="javascript:window.location.replace('main.asp')" style="width:100px" accesskey="X" onmouseover="showtip(this,event,'Sair (Alt+X)');">&nbsp;
					</td>
				</tr>
			</table>
			<div id=divXls style="display:none;POSITION:relative">
				<table border=0 width=760><tr><td colspan=2 align=right></table>
			</div>
			<input type=hidden name=hdnCheck> 
			<input type=hidden name=hdnNomeCons value="ConsultaOSProvedor">
			<input type=hidden name=hdnTipoProcesso value="<%
			Set objRS = db.execute("CLA_sp_sel_tipoprocessoDesCan")
				While not objRS.Eof 
					Response.Write Trim(objRS("Tprc_ID"))
					objRS.MoveNext
					if not objRS.Eof then
						Response.Write ","
					end if	
				Wend
			%>">
		</form>
		<iframe	id			= "IFrmLista"
			    name        = "IFrmLista" 
			    width       = "800"
			    height      = "0"
			    frameborder = "0"
			    scrolling   = "no" 
			    align       = "left">
		</iFrame>
	</body>
	<script>
		ProcurarSenha()
	</script>
</html>
