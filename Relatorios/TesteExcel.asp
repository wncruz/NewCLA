<!--#include file="../inc/data.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="monta-sql.asp"-->

<html>
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
</table>
</td>
</tr>

</form>
</body>











