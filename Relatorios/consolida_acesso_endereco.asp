<!--#include file="../inc/data.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="RelatoriosCla.asp"-->
<!--#include file="paginacao.js"-->
<!--#include file="monta-sql.asp"-->
<html>
<%		
		'*** busca Variaveis
			Filtro				= request.form("Filtro")		
			if Filtro="" then
				Filtro			= "0"
			end if

		   IDOrdena			= request.form("IDOrdena")

   
		   IDestado 			= request("IDestado")	
		   if request("IDestado")<>""  then
			   Filtro				= "1"
		   end if
		   if request.form("IDestado")	<>"" then
			   IDestado 			= request.form("IDestado")		
			end if   
			
	 	    if request.form("IDEnd_bairro")	<>"" then
			   IDEnd_bairro 			= request.form("IDEnd_bairro")		
			end if  

			if request.form("IDsiglaEnd")	<>"" then
			   IDsiglaEnd 			= request.form("IDsiglaEnd")		
			end if  
	
   		   IDQtde				= converte_inteiroLongo(request.form("IDQtde"),0)
   		   IDQtde1				= converte_inteiroLongo(request.form("IDQtde1"),0)

			if request.form("cboBairro")	<>"" then
			   IDEnd_bairro 			= request.form("cboBairro")		
			end if  

		   IF request.form("cboUF")	<>"" THEN
			   IDestado 			= request.form("cboUF")		   
			END IF
			
			if IDQtde=0 then 
			   IDQtde				= converte_inteiroLongo(request.form("cboQtde_acessos"),0)
			end if   
		   
			if IDQtde1=0 then 
			   IDQtde1				= converte_inteiroLongo(request.form("cboQtde_acessos1"),0)
			end if   

			IDlogradouro 		= request.form("IDlogradouro")
			if IDlogradouro="" then
				IDlogradouro 		= request.form("txtlogradouro")
			end if	

		strSQL = Monta_SQL_consolida_endereco_dispon_consol()
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
	mform.action 		 		= "consolida_acesso_endereco.asp"
	mform.IDOrdena.value  		= mform.IDOrdena.value;
	mform.filtro.value			= "1";
	mform.target 				= "_self";
	mform.method 				= "post";
	mform.submit();
}

function enviar(IDEstado,IDEnd_sigla,IDEnd_Nome,IDEnd_bairro) {

	mform 		   						= document.FormRelat;
	mform.action 						= "detalhe_acesso_endereco_fisico_dispon.asp";
	mform.filtro.value					= "1";
	mform.IDEstado.value 				= IDEstado;
	mform.IDEnd_bairro.value 			= IDEnd_bairro;
	mform.IDEnd_sigla.value				= IDEnd_sigla;
	mform.IDEnd_Nome.value				= IDEnd_Nome;		
	mform.IDOrdena.value 				= "";
	mform.target = "_self";
	mform.method = "post"; 
	mform.submit();
}



function ordenar(IDOrdena){
	mform 		           = document.FormRelat;
	mform.action 			 ="consolida_acesso_endereco.asp"
	mform.IDOrdena.value  = IDOrdena;
	mform.filtro.value	 = "1";
	mform.target = "_self";
	mform.method = "post";
	mform.submit();
}

function Imprimir()
{
	window.print()
}

function RelExcel(){
	mform 		           			= document.FormRelat;
	mform.action 						= "excel_consolida_acesso_endereco.asp"
	mform.IDEstado.value 			= mform.IDEstado.value;
	mform.IDEnd_bairro.value 		= mform.IDEnd_bairro.value;
	mform.IDEnd_sigla.value			= mform.IDEnd_sigla.value;
	mform.IDEnd_Nome.value			= mform.IDEnd_Nome.value;		
	mform.IDQtde.value				= mform.IDQtde.value;
	mform.IDQtde1.value				= mform.IDQtde1.value;
	mform.filtro.value	 			= "1";
	mform.target 						= "_blank";
	mform.method 						= "post";
	mform.submit();
}

// --></script>

<table width="100%" border="1">
<tr>
<td>
<Form name="Formfiltro" method="Post" action="consolida_acesso_endereco.asp" target="_self">
<input type="hidden" name="filtro" value="1"  >
<input type=hidden name="IDOrdena" 		value="<%=IDOrdena%>" >
<input type="hidden" name="Npagina"		value="<%=Npagina%>">
<input type=hidden name="IDEstado" 			value="<%=IDEstado%>" >
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


<table width="100%">
<tr class=clsSilver>

 		<td ALIGN="LEFT"><font ></font>UF</td>
		<td >
  		  <% strSQL = Monta_SQL_estado()
			SET RSaux= Server.CreateObject("ADODB.Recordset")
			RSaux.Open strSQL,db	 %>
			<select name="cboUF" onchange="submit()">
				<option ></option>
				<% While Not RSaux.eof %>
					<Option value="<%=RSaux("est_sigla")%>" <%if IDestado=RSaux("est_sigla") then %> selected <% end if %>><%=RSaux("est_sigla")%> - <%=RSaux("est_desc")%></Option>
				<% RSaux.MoveNext
				Wend 
				RSaux.close : set RSaux = nothing
				%>				
       </select>
		</td>
	

		<td ALIGN="RIGHT"><font ></font>Bairro</td>
		<td >
  		  <% strSQL = Monta_SQL_bairro()
			SET RSaux= Server.CreateObject("ADODB.Recordset")
			RSaux.Open strSQL,db	 %>
			<select name="cboBairro" >
				<option ></option>
				<% While Not RSaux.eof %>
					<Option value="<%=RSaux("end_bairro")%>" <%if IDEnd_bairro=RSaux("end_bairro") then %> selected <% end if %>><%=RSaux("end_bairro")%></Option>
				<% RSaux.MoveNext
				Wend 
				RSaux.close : set RSaux = nothing
				%>				
       </select>
		</td>
</tr>
<tr class=clsSilver>
<td ><font ></font>Logradouro</td>
<td ><input type="text" class="text" name="txtlogradouro"  value="<%=IDlogradouro%>"  size="30"></td>
<td ><font ></font>Qtde Acessos &gt;=</td>
		<td >
	  		<select name="cboQtde_acessos">
				<option ></option>
				<option value=1 <%if IDqtde=1 then %> selected <% end if %>>1</option>
				<option value=5  <%if IDqtde=5 then %> selected <% end if %>>5</option>
				<option value=10  <%if IDqtde=10 then %> selected <% end if %>>10</option>				
				<option value=15  <%if IDqtde=15 then %> selected <% end if %>>15</option>
				<option value=20  <%if IDqtde=20 then %> selected <% end if %>>20</option>
				<option value=30  <%if IDqtde=30 then %> selected <% end if %>>30</option>
				<option value=40  <%if IDqtde=40 then %> selected <% end if %>>40</option>
				<option value=50  <%if IDqtde=50 then %> selected <% end if %>>50</option>
				<option value=100  <%if IDqtde=100 then %> selected <% end if %>>100</option>
				
				
	        </select>
      e Qtde Acessos &lt;=<select name="cboQtde_acessos1">
				<option ></option>
				<option value=1 <%if IDqtde1=1 then %> selected <% end if %>>1</option>
				<option value=5  <%if IDqtde1=5 then %> selected <% end if %>>5</option>
				<option value=10  <%if IDqtde1=10 then %> selected <% end if %>>10</option>				
				<option value=15  <%if IDqtde1=15 then %> selected <% end if %>>15</option>
				<option value=20  <%if IDqtde1=20 then %> selected <% end if %>>20</option>
				<option value=30  <%if IDqtde1=30 then %> selected <% end if %>>30</option>
				<option value=40  <%if IDqtde1=40 then %> selected <% end if %>>40</option>
				<option value=50  <%if IDqtde1=50 then %> selected <% end if %>>50</option>
				<option value=100  <%if IDqtde1=100 then %> selected <% end if %>>100</option>				
	
	        </select>
    
			</td>
</tr>
</table>
<table width="100%">
<tr class=clsSilver>
   <td >
        <p align="right">	<input type=button name="Filtrar" value="Filtrar" class="button" onclick="filtrar()" >
	</p>
	</td>
   <td>
        <p align="left">&nbsp;
        <input type=button name="Voltar" value="Voltar" class="button" onclick="javascript:window.history.back(-1)" ></p>
    </td>		
</tr>
</table>
</table>
</form>
<table width="100%" border="0">
<tr>
<td>
<h5 align="center">Relatório de Consolidado por Logradouro</h5>
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
 <th><a class="white" target="_self" href="javascript:ordenar('Estado')">Estado</a></th>
 <th><a class="white" target="_self" href="javascript:ordenar('End_bairro')">Bairro</a></th>
 <th><a class="white" target="_self" href="javascript:ordenar('End_tpl_sigla,End_Nomelogr,End_NroLogr')">Logradouro</a></th>
 <th><a class="white" target="_self" href="javascript:ordenar('qtde_fisico')">Qtde Acessos Físicos</A></th>
 <% if not  RS.eof then
   RS.AbsolutePage = converte_inteiro(Npagina,1)
   While Not RS.eof  and qtde < RS.PageSize
 
	 qtde = qtde +1 
	 Tqtde = Tqtde + converte_inteiro(RS("qtde_fisico"),0)
  %>
<tr>    

 <td align="right"><% =qtde %></td> 
 <td><%=RS("estado")%></td>
 <td><%=RS("end_bairro")%></td>
 <td><%=trim(RS("End_tpl_sigla") & " " & RS("End_Nomelogr") )%></td>
  <td align="center"><a  href="javascript:enviar('<%=RS("Estado")%>','<%=RS("End_tpl_sigla")%>','<%=RS("End_Nomelogr")%>','<%=RS("End_bairro")%>');" target="_self"><%=formatnumber(RS("qtde_fisico"),0)%></a></td>    
</tr> 
<%    
     RS.MoveNext
  Wend 
  RS.close : set RS = nothing	%>
 		
<% end if %> 
</tr>
<tr class=clsSilver>
<td colspan="4"></td>  
<td align="center"><%= formatnumber(Tqtde,0) %></td>  
</tr>	 	
</table>
</td>

</table>

</form>
</body>






