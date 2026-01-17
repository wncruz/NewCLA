<!--#include file="../inc/data.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="RelatoriosCla.asp"-->

<html>
<%

	'*** busca Variaveis
			Filtro				= request.form("Filtro")		
			if Filtro="" then
				Filtro			= "0"
			end if

  		   IDOrdena			= request.form("IDOrdena")

  		   
  		   IDestado 			= request.form("IDestado")
  		   IDservico 			= trim(request.form("IDservico"))
		   IDQtde				= converte_inteiroLongo(request.form("IDQtde"),0)
   		   IDQtde1				= converte_inteiroLongo(request.form("IDQtde1"),0)

  		   if IDporte="" then
			   IDporte				= request.form("cboPorte")
			end if   

  		   if IDestado ="" then
			   IDestado 			= request.form("cboUF")		   
			end if   

  		   if IDservico ="" then
			   IDservico 			= trim(request.form("cboServico"))	
			end if   

			if IDQtde=0 then 
			   IDQtde				= converte_inteiroLongo(request.form("cboQtde_acessos"),0)
			end if   
		   

			if IDQtde1=0 then 
			   IDQtde1				= converte_inteiroLongo(request.form("cboQtde_acessos1"),0)
			end if   
  

  		   


%>

<!--#include file="monta-sql.asp"-->
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

	mform 		   			 	= document.Formfiltro;
	mform.action 			 	= "consolida_acesso_servico.asp"
	mform.IDOrdena.value  	= mform.IDOrdena.value;
	mform.filtro.value		="1";
	mform.target 				= "_self";
	mform.method 				= "post";
	mform.submit();

}

function enviar(IDEstado,IDservico) {

	mform 		   				= document.FormRelat;
	mform.filtro.value		= "1";
	mform.IDEstado.value 	= IDEstado;
	mform.IDservico.value	= IDservico;
	mform.IDOrdena.value 	= "";
	mform.action 				= "detalhe_acesso_endereco_servico.asp";
	mform.target 				= "_self";
	mform.method 				= "post"; 
	mform.submit();
}

function ordenar(IDOrdena){
	mform 		           	= document.FormRelat;
	mform.action 			 	= "consolida_acesso_servico.asp"
	mform.IDOrdena.value  	= IDOrdena;
	mform.filtro.value	 	= "1";
	mform.target 				= "_self";
	mform.method 				= "post";
	mform.submit();
}

function Imprimir()
{
	window.print()
}

// --></script>

<Form name="Formfiltro" method="Post" action="consolida_acesso_servico.asp" target="_self">
<input type="hidden" name="filtro"       value="1"  >
 <input type=hidden name="IDOrdena" 		value="<%=IDOrdena%>" >
<table width="100%">

<tr class=clsSilver>

 		<td ><font ></font>UF</td>
		<td>
  		  <% strSQL = Monta_SQL_estado()
			SET RS= Server.CreateObject("ADODB.Recordset")
			RS.Open strSQL,db	 %>
			<select name="cboUF" >
				<option ></option>
				<% While Not RS.eof %>
					<Option value="<%=RS("est_sigla")%>" <%if trim(IDestado)=trim(RS("est_sigla")) then %> selected <% end if %>><%=RS("est_sigla")%> - <%=RS("est_desc")%></Option>
				<% RS.MoveNext
				Wend 
				RS.close : set RS = nothing
				%>				
       </select>
		</td>
	
		<td class=clsSilver>Serviço</td>
		<td >	
		<%
	
       strSQL = Monta_SQL_Servico()
		SET RS= Server.CreateObject("ADODB.Recordset")
		RS.Open strSQL,db
		%><select name="cboServico" >
				<option ></option>
			
			<% While Not RS.eof %>
					<Option value="<%=RS("Ser_Desc")%>" <%if trim(IDservico)=trim(RS("Ser_Desc")) then %> selected <% end if %>> <%=RS("Ser_Desc")%></Option>
					<% RS.MoveNext
			Wend 
			RS.close : set RS = nothing			
			%>
		 </select> </td>


		
		
	  
	

</tr>
<tr class=clsSilver>
	<td  colspan="4"><font ></font>Qtde Acessos &gt;=
		
	  <% strSQL = Monta_SQL_qtde_acessos_servico_consolida()
	   if strSQL<> "" then
	 		SET RS= Server.CreateObject("ADODB.Recordset")
			RS.Open strSQL,db	 %>
			<select name="cboQtde_acessos">
				<option ></option>
				<% While Not RS.eof %>
					<Option value="<%=RS("qtde_Acesso")%> " <%if IDqtde=RS("qtde_Acesso") then %> selected <% end if %>><%=RS("qtde_Acesso")%></Option>
				<% RS.MoveNext
				Wend 				
		 %>
	        </select>
      e Qtde Acessos &lt;=<select name="cboQtde_acessos1">
				<option ></option>
				<% RS.movefirst
				  While Not RS.eof %>
					<Option value="<%=RS("qtde_Acesso")%> " <%if IDqtde1=RS("qtde_Acesso") then %> selected <% end if %>><%=RS("qtde_Acesso")%></Option>
				<% RS.MoveNext
				Wend 				
		 %>
	        </select>
        <% end if %>   
			</td>
</tr>
<tr class=clsSilver>
	   <td colspan="2">
        <p align="right"><input type=button name="Filtrar" value="Filtrar" class="button" onclick="filtrar()" >	</p>
	</td>
	   <td colspan="2">
        <p align="left"><input type=button name="Voltar" value="Voltar" class="button" onclick="javascript:window.history.back(-1)" ></p>
       </td>
		
</tr>
</table>
</form>
<table width="100%" border="0">
<tr>
<td>
<h5 align="center">Relatório Consolidado de Porte por Serviço</h5>
<center>
<Form name="FormRelat" method="Post" action="consolida_acesso_servico.asp" target="_self" >
 <input type=hidden name="filtro" 				value="1" >
 <input type=hidden name="IDEstado" 			value="<%=IDEstado%>" >
 <input type=hidden name="IDservico" 			value="<%=IDservico %>" >
 <input type=hidden name="IDOrdena" 			value="<%=IDOrdena%>" >
 <input type=hidden name="IDQtde" 				value="<%=IDQtde%>" >
 <input type=hidden name="IDQtde1" 			value="<%=IDQtde1%>" >

</center>
<tr>
<td>
<br>
<!--************ MONTA A TABELA DE RELATÓRIO ****************** !-->
<table width="60%" border="1" align="center" class="TableLine">
 <tr>
 <th align="right">#</th>
 
 <th ><a class="white" target="_self" href="javascript:ordenar('Estado')">UF</A></th>
 <th ><a class="white" target="_self" href="javascript:ordenar('servico_nome')">Serviço</A></th>
 <th ><a class="white" target="_self" href="javascript:ordenar('Valor_total')">Qtde Acessos</A>
  Físicos</th>

 
<% strSQL = Monta_SQL_consolida_porte_servico() %>

<%if strSQL<> "" then

   qtde =0 
	SET RS= Server.CreateObject("ADODB.Recordset")
	RS.Open strSQL,db	
   While Not RS.eof 
    qtde = qtde +1 
   	 Tqtde  = Tqtde + RS("Valor_total")   
   %>
<tr>    

 <td align="right"><% =qtde %></td>
 <td><%=RS("Estado")%></td>  
 <td><%=RS("servico_nome")%></td>  
 <td align="center"><a target="_self" href="javascript:enviar('<%=RS("Estado")%>','<%=RS("servico_nome")%>');"><%=formatnumber(RS("Valor_total"),0)%></a></td>    
</tr> 
<%   
     RS.MoveNext
  Wend 
  RS.close : set RS = nothing	%>
 		
<% end if %>  	
<tr class=clsSilver>
<td colspan="3"></td>  
<td align="center"><%=formatnumber(Tqtde,0) %></td>
</table>
</td>
</tr>
</table>
<table width="100%" border=0>
<tr>
	<td align="center">
		<input type="button" class="button" name="btnImprimir" value="Imprimir" onClick="Imprimir()">&nbsp;
	</td>
</tr>
</table>
</form>
</body>


