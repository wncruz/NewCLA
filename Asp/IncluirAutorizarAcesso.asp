<!--#include file="../inc/data.asp"-->
<!--#include file="../inc/header.asp"-->
<%

strCboTpAprovacao = request("cboTpAprovacao")

dim strSql

if request("hdnAcao")="Gravar" then
	strContrato = null
	strSer_Id = null
	strVel_ID = null
	strProvedor_ID = null
	strCboSistemaOrderEntry = null
	strOeNumero = null
	strOeAno = null
	strOeItem = null
	strPed_Numero = null
	strCusto = null
	strSenha = null
    
    if (strCboTpAprovacao <> "") then
	   if strCboTpAprovacao = 1 then
		 strContrato = Replace(request("txtContrato"),"'","")
		 strSer_Id = request("cboServicoPedido")
		 strCboSistemaOrderEntry = request("cboSistemaOrderEntry")
		 strVel_ID = request("cboVelAcesso")
		 strProvedor_ID = request("cboCLA_Provedor")
		 strOeNumero = request("txtOeNumero")
		 strOeAno = request("txtOeAno")
		 strOeItem = request("txtOeItem")
		 strSenha = request("txtSenha")
	   else
	     strPed_Numero = request("txtPedNum")
		 strCusto = request("txtCusto")
	   end if
		
	    Vetor_Campos(1)="adInteger,1,adParamInput,"  & strCboTpAprovacao
		Vetor_Campos(2)="adWChar,30,adParamInput,"   & strContrato
		Vetor_Campos(3)="adInteger,8,adParamInput,"  & strSer_Id
		Vetor_Campos(4)="adInteger,8,adParamInput,"  & strVel_ID
		Vetor_Campos(5)="adInteger,8,adParamInput,"  & strProvedor_ID
		Vetor_Campos(6)="adWChar,20,adParamInput,"   & strCboSistemaOrderEntry
	    Vetor_Campos(7)="adWChar,7,adParamInput,"    & strOeNumero
		Vetor_Campos(8)="adWChar,4,adParamInput,"    & strOeAno
		Vetor_Campos(9)="adWChar,8,adParamInput,"    & strOeItem
		Vetor_Campos(10)="adWChar,13,adParamInput,"  & strPed_Numero
		Vetor_Campos(11)="adWChar,20,adParamInput,"  & strCusto
		Vetor_Campos(12)="adWChar,10,adParamInput,"  & strSenha
		Vetor_Campos(13)="adWChar,30,adParamInput,"  & strLoginRede
		Vetor_Campos(14)="adInteger,4,adParamOutput, null"
		Vetor_Campos(15)="adWChar,100,adParamOutput, null"
		
		Call APENDA_PARAM("CLA_sp_ins_AprovAcesso",15,Vetor_Campos)
		
		ObjCmd.Execute'pega dbaction
		DBAction = ObjCmd.Parameters("RET").value
		DBAction2 = trim(ObjCmd.Parameters("RET1").value)
		if 	DBAction = 1 then
			Response.write "<script>alert('"&DBAction2&"')</script>"
			strContrato = null
			strSer_Id = null
			strVel_ID = null
			strProvedor_ID = null
			strOeNumero = null
			strOeAno = null
			strOeItem = null
			strPed_Numero = null
			strCusto = null
			Response.write "<script>window.location='CONSULTARAUTORIZARACESSO.asp'</script>"
		else
		    %>
		    <script language="VBscript">
        	  MsgBox "Erro ao gravar a senha PIN.",16,"Erro <%=DBAction%>"
            </script>
			<%
		end if
	End if
End if

dblacao = ""
%>


<script language='javascript' src="../javascript/cla.js"></script>

<form name="Form_1" action="IncluirAutorizarAcesso.asp" method="post" >
<SCRIPT LANGUAGE="JavaScript">
function incluir()
{
	with (document.forms[0])
	{
		if (cboTpAprovacao.value == 1)
		{
			if (txtContrato.value != "" && cboServicoPedido.value != "" && cboCLA_Provedor.valeu != "" && cboVelAcesso.value != "" && txtOeNumero.value != "" && txtOeAno.value != "" && txtOeItem.value != "" && txtSenha.value != "" && cboSistemaOrderEntry.value != "")
			{
			  hdnAcao.value = "Gravar";
			  submit();
			}
			else
			{
				alert("Informe os campos obrigatórios :: ")
				txtContrato.focus();
			}

		}
	    else
	    {
	    
	    	if (txtPedNum.value != "" && txtCusto.value != "" && txtSenha.value != "")
			{
			  if(txtPedNum.value.length > 4)
			  {
	   	  		if (!ValidarDM(txtPedNum)) return
	   	  		
	   	  		hdnAcao.value = "Gravar";
			    submit();
			  }
			}
			else
			{
				alert("Informe os campos obrigatórios :: ")
				txtPedNum.focus();
			}

	    }
	}
}


//PRSSILV - AJAX
var xmlhttp = null;

function GerarSenha() {
    try { 
        xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); 
    } catch (e) { 
        try { 
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); 
        } catch (E) { 
            xmlhttp = false; 
        } 
    } 

    if  (!xmlhttp && typeof  XMLHttpRequest != 'undefined' ) { 
        try  { 
            xmlhttp = new  XMLHttpRequest(); 
        } catch  (e) { 
            xmlhttp = false ; 
        } 
    }

    if (xmlhttp) {
        xmlhttp.onreadystatechange = processadorMudancaEstado;
        xmlhttp.open("GET", "../Ajax/AJX_Resgatar_Senha.asp");
        xmlhttp.setRequestHeader('Content-Type','text/xml');
        xmlhttp.setRequestHeader('encoding','ISO-8859-1');
        xmlhttp.send(null);
    }
}

function processadorMudancaEstado () {
    if ( xmlhttp.readyState == 4) { // Completo 
        if ( xmlhttp.status == 200) { // resposta do servidor OK 
            document.getElementById("txtSenha").value = xmlhttp.responseText;           
        } else { 
            alert( "Erro: " + xmlhttp.statusText );  
        } 
    }
}

function detalhes(var_detalhes,var_span,var_span2)
{
if (var_detalhes == 'EXIBIR')
  {
  var_id_exibir = 'SPAN_EXIBIR'+var_span;
  var_id_ocultar = 'SPAN_OCULTAR'+var_span;
  var_id_exibir2 = 'SPAN_EXIBIR'+var_span2;
  var_id_ocultar2 = 'SPAN_OCULTAR'+var_span2;
  document.all[var_id_exibir].style.display = "block";
  document.all[var_id_ocultar].style.display = "none";
  document.all[var_id_exibir2].style.display = "block";
  document.all[var_id_ocultar2].style.display = "none";
  }
if (var_detalhes == 'OCULTAR')
  {
  var_id_exibir = 'SPAN_EXIBIR'+var_span;
  var_id_ocultar = 'SPAN_OCULTAR'+var_span;
  var_id_exibir2 = 'SPAN_EXIBIR'+var_span2;
  var_id_ocultar2 = 'SPAN_OCULTAR'+var_span2;

  document.all[var_id_exibir].style.display = "none";
  document.all[var_id_ocultar].style.display = "block";
  document.all[var_id_exibir2].style.display = "none";
  document.all[var_id_ocultar2].style.display = "block";
  }
}

function VerificaCboAproRedirect() 
{
if (document.getElementById("cboTpAprovacao").value == "1")
  { 	
    <%
	if request.querystring("tp") = 1 then
		%>
		window.location.href="ConsultarAutorizarAcesso.asp"
		<%
	else
		%>
		window.location.href="ConsultarAprovarAcesso.asp"
	<%end if%>
  }
if (document.getElementById("cboTpAprovacao").value == "2")
  { 
	detalhes('OCULTAR','1','2');
  }
}
function onseleciona_1() 
{ 
if (document.getElementById("cboTpAprovacao").value == "1")
  { 
    detalhes('EXIBIR','1','2');
  
  }
if (document.getElementById("cboTpAprovacao").value == "2")
  { 
  detalhes('OCULTAR','1','2');
  }
}

function ValidarItemOE(campo)
  {
    if (campo.value == "0")
    {
	  campo.value = "001"
	}							    
}
</script>
<input type=hidden name=hdnAcao>
<tr>
	<td >
		<table border="0" cellspacing="1" cellpadding=0 width="760">
			<tr>
				<th colspan=2><p align=center>Autorização de Acessos Terceiros.</p></td>
			</tr>

			<tr class=clsSilver>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<table border="0" cellspacing="1" cellpadding=0 width="760">
				<tr>
					<td width="83">
						<tr class=clsSilver>
							<th colspan=3 width="651"><p>Tipo de 
							Aprovação:&nbsp;
							<select size="1" ID="cboTpAprovacao" name="cboTpAprovacao" tabindex="1" onChange="VerificaCboAproRedirect()">
							<option value="1" <%if strCboTpAprovacao = 1 then%>selected<%end if%>>Circuito Terceiro</option>
							
							<%
							Set objRSPerf = db.execute("CLA_sp_view_loginusuario '" & strLoginRede & "'")
							If not objRSPerf.eof then
								var_Usu_PerfCadSenha = objRSPerf("Usu_PerfCadSenha")
								var_Usu_PerfAltDesig = objRSPerf("Usu_PerfAltDesig")
							End if

							if var_Usu_PerfCadSenha = 1 then
							%>
							<option value="2" <%if strCboTpAprovacao = 2 then%>selected<%end if%>>Projeto Especial de Acesso</option>
							<%end if%>
							</select></p>
						</tr>
							<tr class=clsSilver>
							<td width="83">&nbsp;</td>
							<td width="658" colspan="2">&nbsp;</td>
						</tr>
		</table>
		<span id='SPAN_EXIBIR1' style='display:block'>
		<table border="0" cellspacing="1" cellpadding=0 width="760">
						<tr class=clsSilver>
							<td width="83"><font class="clsObrig">:: </font>Contrato</td>
							<td width="658" colspan="2">
							<input type="text" class="text" name="txtContrato" value="<%=strContrato%>" maxlength="30" size="35" tabindex="2" ></td>
						</tr>
						
						
					   	<tr class=clsSilver>
					 		<td width="83"><font class="clsObrig">:: </font>Serviço</td>
							<td width="658" colspan="2">
							<select name="cboServicoPedido" tabindex="3">
				<option ></option>
			<%
			    set objRS = db.execute("CLA_sp_sel_servico null,null,null,1")
			    
				While Not objRS.eof
					strItemSel = ""
					if Trim(strSer_Id) = Trim(objRS("Ser_ID")) then strItemSel = " Selected " End if
					Response.Write "<Option value='" & objRS("Ser_ID") & "'" & strItemSel & ">" & objRS("Ser_Desc") & "</Option>"
					objRS.MoveNext
				Wend
				strItemSel = ""
			%>
			</select>							
							</td>
						</tr>
																		
						<tr class=clsSilver>
			<td width="83" nowrap><font class="clsObrig">:: </font>Provedor</td>
			<td width="512" nowrap colspan="2">
				<select name="cboCLA_Provedor">
					<option value=""></option>
					<%'Provedores
					set objRS = db.execute("CLA_sp_sel_provedor")
					do while not objRS.eof
					strItemSel = ""
					if Trim(strProvedor_ID) = Trim(objRS("Pro_id")) then strItemSel = " Selected " End if
					%>
					<option value="<%=objRS("Pro_id")%>" <%=strItemSel%>><%=objRS("Pro_Nome")%></option>
					<%
						objRS.movenext
					loop
					%>
				</select>
			</td>
		</tr>
					   	<tr class=clsSilver>
					 		<td width="83" nowrap><font class="clsObrig">:: </font>Velocidade do Acesso:</td>
							<td width="512" nowrap colspan="2">
							  <select name="cboVelAcesso" style="width:150px">
				              <option ></option>
				<%
					set objRS = db.execute("CLA_sp_sel_velocidade ")
					While Not objRS.eof
						strItemSel = ""
						if Trim(strVel_ID) = Trim(objRS("Vel_ID")) then strItemSel = " Selected " End if
						Response.Write "<Option value='" & Trim(objRS("Vel_ID")) & "'" & strItemSel & ">" & objRS("Vel_Desc") & "</Option>"
						objRS.MoveNext
					Wend
					strItemSel = ""
				%>
			</select>
							</td>
						</tr>
																		
						
					   	<tr class=clsSilver>
					 		<td width="83" nowrap><font class="clsObrig">:: </font>OE</td>
							<td width="512" nowrap colspan="2"><select name="cboSistemaOrderEntry"	<%if strOrigem="APG" then%> <%=bbloqueia%> <%else%> <%=bdesbloqueia%> <% End if%>>
					<Option ></Option>
					<Option value="APG"			<%if strOrderEntrySis = "APG" then Response.Write " selected " End If%>>APG</Option>
					<Option value="CFD"			<%if strOrderEntrySis = "CFD" then Response.Write " selected " End If%>>CFD</Option>
					<Option value="SGA VOZ 0300"			<%if strOrderEntrySis = "SGA VOZ 0300" then Response.Write " selected " End If%>>SGA VOZ 0300</Option>
					<Option value="SGA VOZ 0800 FASE 1"		<%if strOrderEntrySis = "SGA VOZ 0800 FASE 1" then Response.Write " selected " End If%>>SGA VOZ 0800 FASE 1</Option>
					<Option value="SGA VOZ VIP'S"			<%if strOrderEntrySis = "SGA VOZ VIP'S" then Response.Write " selected " End If%>>SGA VOZ VIP'S</Option>
					<Option value="SGA DADOS"	<%if strOrderEntrySis = "SGA DADOS" then Response.Write " selected " End If%>>SGA DADOS</Option>
					<Option value="SGA PLUS"	<%if strOrderEntrySis = "SGA PLUS" then Response.Write " selected " End If%>>SGA PLUS</Option>
					<Option value="ADFAC"		<%if strOrderEntrySis = "ADFAC" then Response.Write " selected " End If%>>ADFAC</Option>
					<Option value="CFM"			<%if strOrderEntrySis = "CFM" then Response.Write " selected " End If%>>CFM</Option>
					<Option value="CFT"			<%if strOrderEntrySis = "CFT" then Response.Write " selected " End If%>>CFT</Option>
				</Select>
							Ano: 
							<input type="text" class="text" name="txtOeAno" value="<%=strOeAno%>" maxlength="4" size="4" tabindex="4" TIPO="N" onBlur="CompletarCampo(this)" onKeyUp="ValidarTipo(this,0)">&nbsp; Número: 
							 <input type="text" class="text" name="txtOeNumero" value="<%=strOeNumero%>" maxlength="7" size="7" tabindex="5" TIPO="N" onBlur="CompletarCampo(this)" onKeyUp="ValidarTipo(this,0)"> Item: 
							<input type="text" class="text" name="txtOeItem" value="<%=strOeItem%>" maxlength="3" size="3" tabindex="6" TIPO="N" onBlur="CompletarCampo(this);ValidarItemOE(this);" onKeyUp="ValidarTipo(this,0)"></td>
						</tr>
																		
						
					   	<tr class=clsSilver>
					 		<td width="83">&nbsp;</td>
							<td width="329">&nbsp;
							</td>
							<td width="463">&nbsp;
				</td>
						</tr>
		</table>
		
		<table border="0" cellspacing="1" cellpadding=0 width="760">			
					   	<tr class=clsSilver>
					 		<td width="83"><b>Senha</b></td>
							<td width="329">							
							<input type="text" class="text" ID="txtSenha" name="txtSenha" value="<%=strSenha%>" maxlength="10" size="11" readonly>
							<input type="button" class="button" name="btnGerar" value="Gerar Nova" style="cursor:hand;" onClick="GerarSenha()" accesskey="I" onMouseOver="showtip(this,event,'Gerar Senha');" tabindex="8"></td>
							<td width="463">&nbsp;
				</td>
						</tr>
																		
						
					   	<tr class=clsSilver>
					 		<td width="83">&nbsp;</td>
							<td width="658" colspan="2">&nbsp;</td>
						</tr>
																		
					</td>
				</tr>
			</table>

		</table>
		</span>
		<span id='SPAN_EXIBIR2' style='display:none'></span>
		
		<span id='SPAN_OCULTAR1' style="display:none">
		<table border="0" cellspacing="1" cellpadding=0 width="760">
		<tr class=clsSilver>
		  <td width="156"><font class="clsObrig">:: </font>Pedido de Acesso:</td>
		  <td>

			<input type="text" class="text" name="txtPedNum" value="<%if request("txtPedNum") <> "" then response.write ucase(request("txtPedNum")) else response.write "DM-" end if%>" maxlength="13" size="15"></td>
		</tr>
		<tr class=clsSilver>
		  <td width="156"><font class="clsObrig">:: </font>Custo (R$):</td>
		  <td>
			<input type="text" class="text" name="txtCusto" value="<%=strCusto%>" maxlength="14" size="20" tabindex="3" onKeyPress="return(MascaraMoeda(this,'.',',',event,14))"></td>
		</tr>
		<tr class=clsSilver>
		  <td width="156">&nbsp;</td>
		  <td>&nbsp;</td>
		</tr>
		</table>
		</span>
		<span id='SPAN_OCULTAR2' style="display:none"></span>	
																
		
		<table border="0" cellspacing="1" cellpadding=0 width="760">
		<tr>
			<td colspan=2 align="center"><br>
				<input type="button" class="button" name="Incluir" value="Gravar" onClick="incluir()" accesskey="I" onMouseOver="showtip(this,event,'Incluir (Alt+I)');" tabindex="9">&nbsp;
				<input type="button" class="button" name="btnLimpar" value="Limpar" onClick="LimparForm();setarFocus('cboTpAprovacao');document.Form_1.cboTpAprovacao.value='1';GerarSenha();onseleciona_1();" accesskey="L" onMouseOver="showtip(this,event,'Limpar (Alt+L)');">&nbsp;
				<input type="button" class="button" name="Voltar" value="Voltar" onClick="javascript:window.location.replace('ConsultarAutorizarAcesso.asp')" accesskey="B" onMouseOver="showtip(this,event,'Voltar (Alt+B)');">
				<input type="button" class="button" name="btnSair" value=" Sair " onClick="javascript:window.location.replace('main.asp')" accesskey="X" onMouseOver="showtip(this,event,'Sair (Alt+X)');">
			</td>
		</tr>
		</table>
		<table width="760" border=0>
		<tr>
			<td>
				<font class="clsObrig">:: </font> Campos de preenchimento obrigatório.
			</td>
		</tr>
		<tr>
			<td>&nbsp;
				</td>
		</tr>
		<tr>
			<td height="23">&nbsp;
				</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</body>
<script>
GerarSenha();
onseleciona_1();
</script>
</html>
<%
Set objRS = Nothing
DesconectarCla()
%>