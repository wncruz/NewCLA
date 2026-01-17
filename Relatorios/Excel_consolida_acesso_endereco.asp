<!--#include file="../inc/data.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="monta-sql.asp"-->

<html>
<%		
		'*** busca Variaveis
			Filtro				= request.form("Filtro")		
			if Filtro="" then
				Filtro			= "0"
			end if

		   IDOrdena			= request.form("IDOrdena")
		   IDestado 			= request.form("IDestado")		
		   IDEnd_bairro 		= request.form("IDEnd_bairro")		
		   IDsiglaEnd 		= request.form("IDsiglaEnd")		
   		   IDQtde				= converte_inteiroLongo(request.form("IDQtde"),0)
   		   IDQtde1				= converte_inteiroLongo(request.form("IDQtde1"),0)

		strSQL = Monta_SQL_consolida_endereco_dispon_consol()
		SET rs = Server.CreateObject("ADODB.Recordset")

		rs.Open strSQL, db

%>


<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel=stylesheet type="text/css" href="../css/cla.css">
<title>CLA - Relatório de Acesso</title>
</head>
<body topmargin="0" leftmargin="0">
<!--************ MONTA A TABELA DE FILTROS ****************** !-->
<% Response.ContentType = "application/vnd.ms-excel" %>
<table width="100%" border="1">
<tr>
<td>
<table width="100%" border="0">
<tr>
<td>
<center><h3>CLA - Controle Local de Acesso</h3><center>
<h4 align="center">Relatório de Consolidado por Logradouro  - <%= date() %></h4>
<center>
<Form name="FormRelat" method="Post" action="detalhe_acesso_endereco_fisico_dispon.asp" target="_self" >
 <input type=hidden name="filtro" 				value="1" >
 <input type=hidden name="IDEstado" 			value="<%=IDEstado%>" >
 <input type=hidden name="IDQtde" 				value="<%=IDQtde%>" >
 <input type=hidden name="IDQtde1" 			value="<%=IDQtde1%>" > 
 <input type=hidden name="IDOrdena" 			value="<%=IDOrdena%>" >
 <input type=hidden name="IDEnd_sigla" 		value="<%=IDEnd_sigla%>">
 <input type=hidden name="IDEnd_Nome" 		value="<%=IDEnd_Nome%>">
 <input type=hidden name="IDlogradouro" 		value="<%=IDlogradouro%>">
 <input type=hidden name="IDEnd_bairro" 		value="<%=IDEnd_bairro%>"> 
<input type=hidden name="IDacf_id" 			value="<%=IDacf_id%>" > 
 <input type=hidden name="Npagina" 			value="<%=Npagina%>" >
</center>
<tr>
<td>
<br>
<!--************ MONTA A TABELA DE RELATÓRIO ****************** !-->

<table width="80%" border="1" align="center" class="TableLine">
<tr>
 <th align="center">#</th>
 <th>Estado</th>
 <th>Bairro</th>
 <th>Logradouro</th>
 <th>Qtde Acessos Físicos</th>
 <% if not  RS.eof then
   While Not RS.eof 
 
	 qtde  = qtde  + 1 
	 Tqtde = Tqtde + converte_inteiro(RS("qtde_fisico"),0)

  %>
<tr>    

 <td align="right"><% =qtde %></td> 
 <td><%=RS("estado")%></td>
 <td><%=RS("end_bairro")%></td>
 <td><%=trim(RS("End_tpl_sigla") & " " & RS("End_Nomelogr") )%></td>
 <td align="center"><%=formatnumber(RS("qtde_fisico"),0)%></td>    
</tr> 
<%    
     RS.MoveNext
  Wend 
  RS.close : set RS = nothing	%>
 		
<% end if %> 
<tr class=clsSilver>
<td colspan="4"></td>  
<td align="center"><%= formatnumber(Tqtde,0)%></td>  
</tr>	 	
</table>
</td>
</tr>

</form>
</body>











