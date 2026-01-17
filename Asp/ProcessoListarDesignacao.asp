<HTML>
<HEAD>

<link rel=stylesheet type="text/css" href="../css/cla.css">
<script language='javascript' src="../javascript/cla.js"></script>
<script language='javascript' src="../javascript/claMsg.js"></script>
</HEAD>
<BODY>
<Form name=Form1 method=Post onsubmit="return false">
<input type=hidden name=hdnAprovID value =""> 
<input type=hidden name=hdnAprov_Utilizado value =""> 
<%
	Response.Write Request.Form("hdstrHtmlRet")
%>
<script language="VBScript">
Sub btnAlt_OnCLick 
	varDesig = document.getElementById("txtDesig").value
	varDesig_Antiga = document.getElementById("hdnDesig").value
	
	returnvalue=MsgBox ("Confirma a alteração da designação de '"&varDesig_Antiga&"' para '"&varDesig&"' ?", 4643, "Pergunta")
	If returnvalue=6 Then
		document.Form1.target = "IFrmLista"
		document.Form1.action = "ProcessoAlterarDesignacao.asp"
		document.Form1.submit()
			
	End If
End Sub
</script>
</HTML>
