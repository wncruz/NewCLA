<!--#include file="../inc/data.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="paginacao.js"-->
<!--#include file="RelatoriosCla.asp"-->
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

		rs.PageSize =  ContPagina  'Quantidades de registro por páginas
		rs.Cachesize = ContPagina 'Quantidades de registro por páginas
		rs.CursorLocation = 3

		rs.Open strSQL, db
		Npagina = define_pagina(Request.form("Npagina"),RS.PageCount)
	



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
<SCRIPT LANGUAGE="JavaScript">



function filtrar(){

	mform 		   				= document.Formfiltro;
	mform.action 			 	="detalhe_acesso_servico_cliente.asp"
	mform.filtro.value			="1";
	mform.target 				="_self";
	mform.method 				="post";
	mform.submit();
}


function ordenar(IDOrdena){
	mform 		           	= document.FormRelat;
	mform.action 			 	= "detalhe_acesso_servico_cliente.asp"
	mform.IDOrdena.value 	= IDOrdena;
	mform.filtro.value	 	= "1";
	mform.target 				= "_self";
	mform.method 				= "post";
	mform.submit();
}

function Imprimir()
{
	window.print()
}

function RelExcel(){
	mform 		           			= document.FormRelat;
	mform.action 						= "excel_detalhe_acesso_servico_cliente.asp"
	mform.IDacf_id.value 			= mform.IDacf_id.value;
	mform.filtro.value	 			= "1";
	mform.target 						= "_blank";
	mform.method 						= "post";
	mform.submit();
}

// --></script>

<table width="100%" border="1">
<tr>
<td>
<Form name="Formfiltro" method="Post" action="detalhe_acesso_servico_cliente.asp" target="_self">
<input type="hidden" name="filtro" value="<%=filtro%>"  >
<input type="hidden" name="Npagina"			value="<%=Npagina%>">
<table class="bordafiltro" width="100%"   >  
<tr>
<td>
  Total de Registros :  <%=RS.recordCount%> 
</td>
<td>
<table  width="100%">
<tr><td align="right" width="50%" >
<!--<a target=_self href=javascript:RelExcel()><img src='../imagens/excel.gif' border=0></a>!-->&nbsp
<td align="left" width="50%">
<a target=_self href="javascript:window.print()" ><img src='../imagens/impressora.gif' border=0></a></td>
</tr>
</table>
</td>
<% if RS.recordCount > 0 then %>
<td>
  Página : <%=Npagina%>  de <%=RS.PageCount%>   
</td> 
<td>
 <%'Vamos verificar se não é a página 1, para podermos colocar o link “anterior”. 
IF Npagina > 1 then %> 
    <a  target="_self" href="javascript:Anterior(<%=2%>)">Primeira</a> 
<% END IF %>
</td>    
<td>
 <%'Vamos verificar se não é a página 1, para podermos colocar o link “anterior”. 
IF Npagina > 1 then %> 
    <a  target="_self" href="javascript:Anterior(<%=Npagina%>)">Anterior</a>     
<% END IF %>
</td>
<td>
<%'Se não estivermos no último registro contado, então é mostrado o link p/ a próxima página 
IF (strcomp(Npagina,RS.PageCount) <> 0) or (Npagina < RS.PageCount) then %> 
    <a  target="_self" href="javascript:Proxima(<%=Npagina%>)">Próxima</a>     
<% END IF  %>
</td>
<td>
<% 'Se não estivermos no último registro contado, então é mostrado o link p/ a próxima página 
IF (strcomp(Npagina,RS.PageCount) <> 0) or (Npagina < RS.PageCount) then %> 
    <a  target="_self" href="javascript:Proxima(<%=RS.PageCount-1%>)">Ultima</a>    
<% END IF  %>
</td>
<% end if %>
</table  >    

<!--
<table width="100%">
<tr  class=clsSilver>
 <td ><font ></font>Cliente</td>
    <td>
	 <input type="text" class="text" name="txtCliente"  value="<%=NomeCli%>"  size="30"><input type="submit" name="ProcurarCliente" value="Buscar" class="button" >
    </td>
	  
      <td>
      <% if NomeCli<> "" then 
      	 strSQL = Monta_SQL_cliente()
      	if  strSQL<> "" then
	 		SET RSaux= Server.CreateObject("ADODB.Recordset")
			RSaux.Open strSQL,db	 %>
 			<select name="cbocliente">
				<option ></option>
				<% While Not RSaux.eof %>

					<Option value="<%=RSaux("cli_nome")%>" <% if Cliente = RSaux("cli_nome") then  %> selected <% end if %>><%=RSaux("cli_nome")%></Option>
				<% RSaux.MoveNext
				Wend 
				RSaux.close : set RSaux = nothing	%>
 		
	      	</select>
	    <% end if 
	    end if
	    %>  	
		</td>
	
</tr>
</table>
!-->
<table width="100%">
<tr class=clsSilver>
  <!-- <td >
        <p align="right">	<input type=button name="Filtrar" value="Filtrar" class="button" onclick="filtrar()" >
	</p>
	</td> !-->
   <td>
        <p align="center">&nbsp;
        <input type=button name="Voltar" value="Voltar" class="button" onclick="javascript:window.history.back(-1)" ></p>
    </td>		
</tr>
</table>

</form>
<h5  align="center">Relatório de acessos lógicos</h5>
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
<table width="100%" border="1" class="TableLine">
<tr>
 <th align="center">#</th>

 <th><a class="white" target="_self" href="javascript:ordenar('End_tpl_sigla,End_Nomelogr,End_NroLogr')">Endereço</a></th>
 <th><a class="white" target="_self" href="javascript:ordenar('razao_social')">Razão Social</A></th>
 <th><a class="white" target="_self" href="javascript:ordenar('Conta_Corrente,SubConta')">Conta-Corrente</A></th> 
 <th><a class="white" target="_self" href="javascript:ordenar('Serviço_Nome')">Serviço</a></th>
 <th><a class="white" target="_self" href="javascript:ordenar('Valor_total')">Velocidade</a>
  do Lógico</th>
 

<% %>

<%if NOT RS.eof  then
	qtde 		=0
	Tqtde 		=0
	Totvalor 	=0
   RS.AbsolutePage = converte_inteiro(Npagina,1)
   While Not RS.eof  and qtde < RS.PageSize
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


