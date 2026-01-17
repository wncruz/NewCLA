
<input type="hidden" name="hdnPagina" value="<%=Request.ServerVariables("SCRIPT_NAME")%>"> 
<input type="hidden" name="hdCurrentPage"  value="<%=intCurrentPage%>">  
<input type="hidden" name="hdTotalPages"  value="<%=intTotalPages%>">  
<!--********* -- Controles de Paginaçãoo -- ***************-->
<table border="0" width="758">
	<tr>
	<td align=right><% if intTotalPages <> 0 then Response.Write("Página " & intCurrentPage & " de " & intTotalPages )%><br>

	<%if (intCurrentPage <= intTotalPages) then %>
		Ver página Nº&nbsp;<input type="text" size="2" class=text  name="TbNroPag" onkeyup="{ValidarTipo(this,0)}">
		<input type="button" name="BtNro" value="Ir" class=button onclick="{ValidarPaginacao('PagNro')}" style="width:25px" accesskey="M" onmouseover="showtip(this,event,'P�gina Solicitada(Alt+M)');">
	<%End If%>

	<%if intCurrentPage > 1 then  'Bot�es de navega��o na pagina��o%>
		<input type="button" name="BtAnt" value="<< " class=button onclick="{ValidarPaginacao('PagAnt')}" style="width:25px" accesskey="," onmouseover="showtip(this,event,'P�gina Anterior(Alt+,)');">
	<%End If%>

	<%if (intCurrentPage <= intTotalPages) then %>
		<%if (intCurrentPage < intTotalPages) then %>
			<input type="button" name="BtProx" value=" >>" class=button onclick="{ValidarPaginacao('PagProx')}" style="width:25px" accesskey="." onmouseover="showtip(this,event,'Pr�xima P�gina(Alt+.)');">
		<%End If%>	
	<%End If%>
	</td>
	</tr>	
</table>