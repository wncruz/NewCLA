<!--#include file="../inc/data.asp"-->
<Html>
<Head>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/solicitacao.js"></script>

<script language='javascript' src="../javascript/claMsg.js"></script>

</Head>
<Body topmargin=0 leftmargin=0>
<SCRIPT LANGUAGE=javascript>

function LimparMotivoPendencia()
{
	with (document.forms[0])
	{
		hdnAcao.value = "ResgatarLista"
		target = "IFrmLista"
		action = "ProcessoListaStatusSNOA.asp?strAcao=ResgatarLista&dblSolId=<%=Request.QueryString("dblSolId")%>"
		submit()
	//break	
	}
	//alert('teste')
}

function AceitarConclusaoSNOA()
{
	
	with (document.forms[0])
	{
		
		if (!ValidarCampos(cboAceitarSNOA,"Conclusão de Ordem do SNOA")) return

		if (cboAceitarSNOA.value == "REJEICAOORDEM") 
		{
			if (!ValidarCampos(txtMotivo,"Motivo de Rejeição")) return
		}
		
		if (cboAceitarSNOA.value == "PROVIDENCIAREJEITADA") 
		{
			if (!ValidarCampos(txtMotivo,"Motivo da Rejeição da Providência")) return
		}

		hdnAcao.value = "ConclusaoSNOA"
		target = "IFrmLista"
		action = "ProcessoListaStatusSNOA.asp?strAcao=ConclusaoSNOA&dblSolId=<%=Request.QueryString("dblSolId")%>"
		submit()
	//break	
	}
	//alert('teste')
}

</SCRIPT>
<form name=Form1 method=Post >
	<input type=hidden name=hdnAcao>
	<input type=hidden name=hdnSolId  		value="<%=Request.QueryString("dblSolId")%>">
	<input type=hidden name=hdnPedId  		value="<%=Request.QueryString("dblPedId")%>">
	<input type=hidden name=hdnLibera 		value="<%=Request.QueryString("dblLibera")%>">
	<input type=hidden name=hdnAcfId  		value="<%=Request.QueryString("dblAcfId")%>">
	<!--<input type=hidden name=gravarDireto  	value="<%=Request.QueryString("gravarDireto")%>">-->
	<iframe	id 	    = "IFrmProcessoMotivo"
		name        = "IFrmProcessoMotivo"
		width       = "0"
		height      = "0"
		frameborder = "0"
		scrolling   = "auto"
		align       = "left">
	</iFrame>

	<table border=0 cellspacing="1" cellpadding="0"width="100%">
		<tr>
			<th style="FONT-SIZE: 14px" colspan=2 >&nbsp;•&nbsp;Status Solicitação SNOA</th>
		</tr>
		
		<% if Request.QueryString("telaaceitar") = "1" then%>
			<tr class=clsSilver>
				<!--<td nowrap width=530px>-->
					
				<td nowrap colspan="1">
					&nbsp; Conclusão de Ordem do SNOA 
					&nbsp; 
					<span id="spnCboAceitarSNOA">
						<select name="cboAceitarSNOA" style="width:200px">
						<option value=""></option>
						<%	
							set objRS = db.execute("CLA_sp_sel_StatusSNOA " & Request.QueryString("telaaceitar") )
							
							While not objRS.Eof
								strItemSel = ""
								
								'if Trim(str_vel_idTransmissao) = Trim(objRS("StsSnoa")) then strItemSel = " Selected " End if
								Response.Write "<Option value=" & objRS("StsSnoa") & strItemSel & ">" & objRS("StsSnoaMacro") & "</Option>"
								objRS.MoveNext
							Wend
							strItemSel = ""
						%>
						</select>
					</span>

					&nbsp; 
					<input type=button name=btnConcluir value="Concluir" class=button onclick="AceitarConclusaoSNOA()" accesskey="A" onmouseover="showtip(this,event,'Aceita a Conclusão de Ordem do SNOA (Alt+A)');">

				</td>
				<td>
					<input type=button name=btnRecarregarLista value="Recarregar" class=button onclick="LimparMotivoPendencia()" accesskey="Y" onmouseover="showtip(this,event,'Recarregar a Lista de Solicitações do SNOA (Alt+Y)');">

				</td>
			</tr>

			<tr class=clsSilver>
				<td colspan="2">
					&nbsp;
					Justificativa
					&nbsp;
					<textarea name="txtMotivo" cols="107" rows="2" onKeyPress="MaxLength(this,1000)"></textarea>
				</td>
			</tr>
		<% elseif Request.QueryString("telaaceitar") = "3" then%>
		
			<tr class=clsSilver>
				<!--<td nowrap width=530px>-->
				<td nowrap colspan="1">
					&nbsp; Conclusão da Providência do SNOA 
					&nbsp; 
					<span id="spnCboAceitarSNOA">
						<select name="cboAceitarSNOA" style="width:200px">
						<option value=""></option>
						<%	
							set objRS = db.execute("CLA_sp_sel_StatusSNOA " & Request.QueryString("telaaceitar") )
							While not objRS.Eof
								strItemSel = ""
								'if Trim(str_vel_idTransmissao) = Trim(objRS("StsSnoa")) then strItemSel = " Selected " End if
								Response.Write "<Option value=" & objRS("StsSnoa") & strItemSel & ">" & objRS("StsSnoaMacro") & "</Option>"
								objRS.MoveNext
							Wend
							strItemSel = ""
						%>
						</select>
					</span>

					&nbsp; 
					<input type=button name=btnConcluir value="Concluir" class=button onclick="AceitarConclusaoSNOA()" accesskey="A" onmouseover="showtip(this,event,'Aceita a Conclusão de Ordem do SNOA (Alt+A)');">

				</td>
				<td>
					<input type=button name=btnRecarregarLista value="Recarregar" class=button onclick="LimparMotivoPendencia()" accesskey="Y" onmouseover="showtip(this,event,'Recarregar a Lista de Solicitações do SNOA (Alt+Y)');">

				</td>
			</tr>

			<tr class=clsSilver>
				<td colspan="2">
					&nbsp;
					Justificativa
					&nbsp;
					<textarea name="txtMotivo" cols="107" rows="2" onKeyPress="MaxLength(this,1000)"></textarea>
				</td>
			</tr>
			
		<% else%>

			<tr class=clsSilver>
				<td nowrap width=630px>
				&nbsp;
				</td>

				<td>
					<input type=button name=btnRecarregarLista value="Recarregar" class=button onclick="LimparMotivoPendencia()" accesskey="Y" onmouseover="showtip(this,event,'Recarregar a Lista de Solicitações do SNOA (Alt+Y)');">

				</td>
			</tr>
		<% end if%>


		<tr>
			<td colspan=2 valign=top align=left >
				<iframe	id	    = "IFrmLista"
					name        = "IFrmLista"
					width       = "100%"
					height      = "155px"
					frameborder = "0"
					scrolling   = "auto"
	    			src			= "ProcessoListaStatusSNOA.asp?strAcao=ResgatarLista&dblSolId=<%=Request.QueryString("dblSolId")%>"
					align       = "left">
				</iFrame>
			</td>
		</tr>
	</table>
</Form>
</Body>
</Html>
