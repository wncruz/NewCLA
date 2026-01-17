<!--#include file="../inc/data.asp"-->
<Html>

	<Head>
		<link rel=stylesheet type="text/css" href="../css/cla.css">
		<script language='javascript' src="../javascript/cla.js"></script>
		<script language='javascript' src="../javascript/solicitacao.js"></script>
		<script language='javascript' src="../javascript/claMsg.js"></script>
	</Head>

	<Body topmargin=0 leftmargin=0>

		<form name=Form1 method=Post >
			
			<input type=hidden name=hdnAcao>
			<input type=hidden name=hdnSolId  		value="<%=Request.QueryString("dblSolId")%>">
			<input type=hidden name=hdnPedId  		value="<%=Request.QueryString("dblPedId")%>">
			<input type=hidden name=hdnLibera 		value="<%=Request.QueryString("dblLibera")%>">
			<input type=hidden name=hdnAcfId  		value="<%=Request.QueryString("dblAcfId")%>">

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
					<th style="FONT-SIZE: 14px" colspan=2 >&nbsp;•&nbsp;Designação da Fornecedora SNOA</th>
				</tr>
				
				<!--
				<tr class=clsSilver>
					<td nowrap width=630px>
					&nbsp;
					</td>

					<td>
						<input type=button name=btnRecarregarLista value="Recarregar" class=button onclick="LimparMotivoPendencia()" accesskey="Y" onmouseover="showtip(this,event,'Recarregar a Lista de Solicitações do SNOA (Alt+Y)');">

					</td>
				</tr>
				-->

				<tr>
					<td colspan=2 valign=top align=left >
						<iframe	id	    = "IFrmLista"
							name        = "IFrmLista"
							width       = "100%"
							height      = "155px"
							frameborder = "0"
							scrolling   = "auto"
			    			src			= "ProcessoDesigFornecSNOA.asp?strAcao=ResgatarLista&dblSolId=<%=Request.QueryString("dblSolId")%>"
							align       = "left">
						</iFrame>
					</td>
				</tr>

			</table>

		</Form>

	</Body>

</Html>