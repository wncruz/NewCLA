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
		   
		   IDacf_id			= request.form("IDacf_id")
		   
		

			if NomeCli ="" then
			   NomeCli 			= Trim(request.form("txtCliente")) 
			end if  
			 
			if Cliente ="" then
			   Cliente 			= request.form("cbocliente") 
			end if   



		strSQL = Monta_SQL_detalhe_servico_cliente()
		SET rs = Server.CreateObject("ADODB.Recordset")


		rs.Open strSQL, db
	
%>


<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel=stylesheet type="text/css" href="../css/cla.css">
<title>CLA - Relatório de Acesso</title>
</head>
<body topmargin="0" leftmargin="0">
<!--************ MONTA A TABELA DE FILTROS ****************** !-->

<table width="100%" border="1">
<tr>
<td>
<center><h3>CLA - Controle Local de Acesso</h3><center>
<h4  align="center">Relatório de acessos lógicos  - <%= date() %></h4>
<Form name="FormRelat" method="Post" action="exemplos/detalhe.asp" target="_self">
 <input type=hidden name="filtro" 				value="1">
 <input type=hidden name="IDEnd_sigla" 			value="<%=IDEnd_sigla%>">
 <input type=hidden name="IDEnd_Nome" 			value="<%=IDEnd_Nome%>">
 <input type=hidden name="IDEnd_Numero" 		value="<%=IDEnd_Numero%>">
 <input type=hidden name="IDEnd_bairro" 		value="<%=IDEnd_bairro%>">
 <input type=hidden name="IDsigla" 				value="<%=IDsigla%>">
 <input type=hidden name="IDConta_corrente" 	value="<%=IDConta_corrente%>">
 <input type=hidden name="IDSubConta" 			value="<%=IDSubConta%>">
 <input type=hidden name="IDestado" 			value="<%=IDestado%>">
 <input type=hidden name="IDporte" 				value="<%=IDporte%>">
 <input type=hidden name="IDOrdena" 			value="<%=IDOrdena%>">
 <input type=hidden name="NomeCli" 				value="<%=NomeCli%>">
 <input type=hidden name="Cliente" 				value="<%=Cliente%>"> 
 <input type=hidden name="IDacf_id" 			value="<%=IDacf_id%>" >  
 <input type="hidden" name="Npagina"			value="<%=Npagina%>">
 
 

<br>
<!--************ MONTA A TABELA DE RELATÓRIO ****************** !-->
<%Response.ContentType = "application/vnd.ms-excel"  %>
<table width="100%" border="1" class="TableLine">
<tr>
 <th align="center">#</th>

 <th>Endereço</th>
 <th>Razão Social</th>
 <th>Conta-Corrente</th> 
 <th>Serviço</th>
 <th>Velocidade do Lógico</th>
 

<% %>

<%if NOT RS.eof  then
	qtde 		=0
	Tqtde 		=0
	Totvalor 	=0
   While Not RS.eof 
   		qtde = qtde +1
   
   %>
   
   
<tr>    
<!-- <td align="right"><%=RS("acl_idacessologico")%></td> !-->

 <td align="right"><%=qtde %></td>
 <td><%=trim(RS("Endereco_do_Fisico"))%></td>
 <td><%=RS("Razao_Social")%></td> 
 <td><%=RS("Conta_Corrente")%>-<%=RS("SubConta")%></td>
 <td><%=RS("Serviço_Nome")%></td> 
 <td><%=RS("Vel_DescLogico")%></td>  
 
 
</tr> 
<%    
	   
	   
      RS.MoveNext
  Wend 
  RS.close : set RS = nothing	%>
 		
<% end if %>  
	
</table>
</td>
</tr>
</table>
</form>
</body>


