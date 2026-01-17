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

		   if IDestado	="" then
			   IDestado 			= request.form("IDestado")		
			end if   
			
 			IDEnd_sigla			= request.form("IDEnd_sigla")
			IDEnd_bairro			= request.form("IDEnd_bairro")		
			IDEnd_Nome				= request.form("IDEnd_Nome")	
			IDproprietario		= request.form("IDproprietario")
			IDtecnologia			= request.form("IDtecnologia")
			IDestacao				= request.form("IDestacao")
			IDTipoestacao			= request.form("IDTipoestacao")

			if IDestacao="" then
				IDestacao				= request.form("Cboestacao")
			end if	

			if IDestacao="" then
				IDTipoestacao			= request.form("CboTipoestacao")
			end if	

			if request.form("cboProprietario")<>"" then
				IDproprietario			= request.form("cboProprietario")
			end if
			IDConta_corrente		= request.form("IDConta_corrente")
			IDSubconta				= request.form("IDSubconta")

   		    'IDQtde					= converte_inteiroLongo(request.form("IDQtde"),0)
   		    'IDQtde1				= converte_inteiroLongo(request.form("IDQtde1"),0)

			if request.form("cboBairro")	<>"" then
			   IDEnd_bairro 		= request.form("cboBairro")		
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
			
			if IDtecnologia= "" then
				IDtecnologia  = request.form("CboTecnologia")
			end if


		strSQL =  Monta_SQL_consolida_endereco()
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
	mform.action 		 		= "detalhe_acesso_endereco_fisico_dispon.asp"
	mform.IDOrdena.value  		= mform.IDOrdena.value;
	mform.filtro.value			= "1";
	mform.target 				= "_self";
	mform.method 				= "post";
	mform.submit();
}

function enviar(IDacf_id) {

	mform 		   						= document.FormRelat;
	mform.action 						= "detalhe_acesso_servico_cliente.asp";
	mform.filtro.value					= "1";
	mform.IDacf_id.value 				= IDacf_id;
	mform.IDOrdena.value 				= "";
	mform.target = "_self";
	mform.method = "post"; 
	mform.submit();
}



function ordenar(IDOrdena){
	mform 		           = document.FormRelat;
	mform.action 			 ="detalhe_acesso_endereco_fisico_dispon.asp"
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
	mform.action 						= "Excel_detalhe_acesso_endereco_fisico_dispon.asp"
	mform.IDEstado.value 			= mform.IDEstado.value;
	mform.IDproprietario.value 		= mform.IDproprietario.value;	
	mform.IDEnd_bairro.value		= mform.IDEnd_bairro.value;
	mform.IDQtde.value				= mform.IDQtde.value;
	mform.IDQtde1.value				= mform.IDQtde1.value;
	mform.IDtecnologia.value		= mform.IDtecnologia.value;
	mform.IDestacao.value			= mform.IDestacao.value;
	mform.filtro.value	 			= "1";
	mform.target 						= "_blank";
	mform.method 						= "post";
	mform.submit();

}

// --></script>

<table width="100%" border="1">
<tr>
<td>
<Form name="Formfiltro" method="Post" action="detalhe_acesso_endereco_fisico_dispon.asp" target="_self">
<input type="hidden" name="filtro" value="1"  >
<input type=hidden name="IDOrdena" 			value="<%=IDOrdena%>" >
<input type="hidden" name="Npagina"			value="<%=Npagina%>">
<input type=hidden name="IDEstado" 			value="<%=IDEstado%>" >
<input type=hidden name="IDEnd_Nome" 			value="<%=IDEnd_Nome	%>" >
<input type=hidden name="IDEnd_sigla" 		value="<%=IDEnd_sigla%>" >
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

 		<td ALIGN="left"><font ></font>UF</td>
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
	<td ALIGN="left" ><font ></font>Proprietario</td>
		<td >
  		<select name="cboProprietario" >
				<option ></option>
				<Option value="EBT" <%if IDproprietario="EBT" then %> selected <% end if %>>Embratel</Option>
				<Option value="TER" <%if IDproprietario="TER" then %> selected <% end if %>>Terceiro</Option>
				<Option value="CLI" <%if IDproprietario="CLI" then %> selected <% end if %>>Cliente</Option>
       </select>
		</td>
<td ALIGN="left" ><font ></font>Tecnologia</td>	
	<td >
		
  	 <% strSQL = Monta_SQL_tecnologia()
			SET RSaux= Server.CreateObject("ADODB.Recordset")
			RSaux.Open strSQL,db	 %>
			<select name="CboTecnologia" >
				<option ></option>
				<% While Not RSaux.eof %>
					<Option value="<%=RSaux("tec_nome")%>" <%if IDtecnologia=RSaux("tec_nome") then %> selected <% end if %>><%=RSaux("tec_nome")%></Option>
				<% RSaux.MoveNext
				Wend 
				RSaux.close : set RSaux = nothing
				%>				
       </select>

		</td>
</tr>
<tr class=clsSilver>
	<td ALIGN="left" ><font ></font>Estacao</td>
		<td >
 	
  		 <% strSQL = Monta_SQL_estacao()
			SET RSaux= Server.CreateObject("ADODB.Recordset")
			RSaux.Open strSQL,db	 %>
			<select name="CboEstacao" >
				<option ></option>
				<% While Not RSaux.eof %>
					<Option value="<%=RSaux("EstacaoEntrega")%>" <%if IDestacao=RSaux("EstacaoEntrega") then %> selected <% end if %>><%=RSaux("EstacaoEntrega")%></Option>
				<% RSaux.MoveNext
				Wend 
				RSaux.close : set RSaux = nothing
				%>				
       </select>
		</td>
<td ALIGN="left" ><font ></font>Tipo Estacao</td>	
	<td >
		
  	 <% strSQL = Monta_SQL_Tipo_estacao()
			SET RSaux= Server.CreateObject("ADODB.Recordset")
			RSaux.Open strSQL,db	 %>
			<select name="CboTipoEstacao" >
				<option ></option>
				<% While Not RSaux.eof %>
					<Option value="<%=RSaux("TipoEstacaoEntrega")%>" <%if IDTipoestacao=RSaux("TipoEstacaoEntrega") then %> selected <% end if %>><%=RSaux("TipoEstacaoEntrega")%></Option>
				<% RSaux.MoveNext
				Wend 
				RSaux.close : set RSaux = nothing
				%>				
       </select>

		</td>
</tr>
<tr class=clsSilver>
<td ><font ></font>Qtde Acessos &gt;=</td>
		<td COLSPAN="3">
	  
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
</td>
</tr>
</form>
<table width="100%" border="0">
<tr>
<td>
<h5 align="center">Relatório de acessos físicos</h5>
<center>
<Form name="FormRelat" method="Post" action="detalhe_acesso_endereco.asp" target="_self" >
 <input type=hidden name="filtro" 				value="1" >
 <input type=hidden name="IDEstado" 			value="<%=IDEstado%>" >
 <input type=hidden name="IDQtde" 				value="<%=IDQtde%>" >
 <input type=hidden name="IDQtde1" 			value="<%=IDQtde1%>" > 
 <input type=hidden name="IDOrdena" 			value="<%=IDOrdena%>" >
 <input type=hidden name="IDEnd_sigla" 		value="<%=IDEnd_sigla%>" >
 <input type=hidden name="IDEnd_Nome" 		value="<%=IDEnd_Nome%>" >
 <input type=hidden name="IDEnd_bairro" 		value="<%=IDEnd_bairro%>" > 
 <input type=hidden name="IDproprietario" 	value="<%=IDproprietario%>" >  
 <input type=hidden name="IDtecnologia" 		value="<%=IDtecnologia%>" >  
 <input type=hidden name="IDestacao" 			value="<%=IDestacao%>" >  
 <input type=hidden name="IDTipoestacao" 		value="<%=IDTipoestacao%>" >   
 <input type=hidden name="IDacf_id" 			value="<%=IDacf_id%>" > 
 <input type=hidden name="Npagina" 			value="<%=Npagina%>" >
</center>
<tr>
<td>
<br>
<!--************ MONTA A TABELA DE RELATÓRIO ****************** !-->
<table width="100%" border="1" align="center" class="TableLine">
<tr>
 <th align="center">#</th>
 <th><a class="white" target="_self" href="javascript:ordenar('Estado')">Estado</a></th>
 <th><a class="white" target="_self" href="javascript:ordenar('End_bairro')">Bairro</a></th>
 <th><a class="white" target="_self" href="javascript:ordenar('End_tpl_sigla,End_Nomelogr,End_NroLogr')">Endereço</a></th>
 <th><a class="white" target="_self" href="javascript:ordenar('razao_social)"></A>Razao Social</th>
 <th><a class="white" target="_self" href="javascript:ordenar('Proprietario')"></A>Proprietário</th>
 <th><a class="white" target="_self" href="javascript:ordenar('Tecnologia')">Tecnologia</A></th>
 <th><a class="white" target="_self" href="javascript:ordenar('EstacaoEntrega')">Estação Entrega</A></th>
  <th><a class="white" target="_self" href="javascript:ordenar('TipoEstacaoEntrega')">Tipo Estação Entrega</A></th>  
 <th><a class="white" target="_self" href="javascript:ordenar('Vel_fisico')">Velocidade Físico</A></th>
 <th><a class="white" target="_self" href="javascript:ordenar('disponibilidade')">Disponibilidade Teorica</A></th>
<th><a class="white" target="_self" href="javascript:ordenar('qtde_logico')">Qtde Acessos
  Lógicos</A></th>
 <% if not  RS.eof then
   RS.AbsolutePage = converte_inteiro(Npagina,1)
   While Not RS.eof  and qtde < RS.PageSize
 
	 qtde = qtde +1 
	 Tqtde  = Tqtde + RS("qtde_logico")   
  %>
<tr>    

 <td align="right"><% =qtde %></td> 
 <td><%=RS("estado")%></td>
 <td><%=RS("end_bairro")%></td>
 <td><%=trim(RS("End_tpl_sigla") & " " & RS("End_Nomelogr") & " " & RS("End_NroLogr"))%></td>
 <td><%=rs("razao_social")%></td>
 <td><%=trim(RS("Proprietario")) %></td>
 <td><%=trim(RS("tecnologia")) %></td>
 <td><%=trim(RS("EstacaoEntrega")) %></td>
 <td><%=trim(RS("TipoEstacaoEntrega")) %></td>
 <td align="right"><%=RS("Vel_fisico") %></td>
 <td align="right"><%=RS("disponibilidade") %></td>
 <td align="center"><a  href="javascript:enviar('<%=RS("acf_id")%>');" target="_self"><%=formatnumber(RS("qtde_logico"),0)%></a></td>    
</tr> 
<%    
     RS.MoveNext
  Wend 
  RS.close : set RS = nothing	%>
 		
<% end if %> 
<tr class=clsSilver>
<td colspan="11"></td>  
<td align="center"><%= formatnumber(Tqtde,0) %></td>  
</tr>	 	
</table>
</td>
</tr>
</table>
</form>
</body>







