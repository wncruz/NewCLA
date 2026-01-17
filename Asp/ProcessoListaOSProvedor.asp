<%
'•ACCENTURE
'	- Sistema			: CLA
'	- Arquivo			: ProcessoListaOSProvedor.ASP
'	- Responsável		: Gustavo S. Reynaldo
'	- Descrição			: Página que recebe e exibe o html do Grid das páginas ConsultaOSProvedor e CadastraOSProvedor.
%>
<HTML>
<HEAD>
<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>
</HEAD>
<BODY>
<Form name=Form1 method=Post onsubmit="return false">
<input type="hidden" name="hdstrHtmlRet"  value="<%Request.Form("hdstrHtmlRet")%>">
<%
	Response.Write Request.Form("hdstrHtmlRet")
%>

</HTML>
