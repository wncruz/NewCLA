<!--#include file="../inc/data.asp"-->
<% 

dim objrsCarta

	if Request.Form("hdnPedId") <> 0 then 
		set objrsCarta = db.execute(" select * from cla_documento where ped_id = " &  Request.Form("hdnPedId") )
		if not objrsCarta.eof then 
		
			dim objFSO, objFile, strCaminho 
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			strCaminho = server.MapPath("..\")
			Set objFile = objFSO.CreateTextFile(strCaminho & "\asp\cartaconteudo.txt",  true)
			objFile.WriteLine(objrsCarta("Doc_Conteudo"))
			objFile.Close
			set objFSO = nothing 
			set objFile = nothing 
		else
			Response.Write ("<script language = JavaScript> alert(' Conteudo não localizado ')</script>")
		end if 
	end if 
%>
<HEAD>
<TITLE>Gera arquivo carta</TITLE>
<LINK HREF = "..\css\CLA.CSS" REL ="stylesheet"/>
<SCRIPT LANGUAGE = "JavaScript">

	function  EnviaEmail()
	{
		with (document.forms[0])
		{
			hdnPedId.value = 28596
			target = self.name 
			submit()
		}
	}

	
</SCRIPT>
</HEAD>
<BODY leftmargin="0" topmargin="0">
<FORM  method=post >
<input type="hidden" name="hdnPedId">
<input type="hidden" name="hdnSolId">
<TABLE border=0 cellPadding=0 cellSpacing=1 width="100%" >
<TR>
	<TD valign = "top" background=..\imagens\topo_embratel.jpg  height = 80 colspan = 2 ></TD>
</TR>
<TR>
	<th colspan=2><p align=center>Gera Arquivo da Carta</p></th>
	<!--<TD  colspan = 2 style = "font-size:10pt;COLOR:white;FONT-WEIGHT: bold" bgcolor=SteelBlue align  = center >Envio manual de e-mail</TD>-->
</TR>
</TABLE>
<P></P>
<TABLE WIDTH = 100%>
<TR>
	<td colspan = 2 align =CENTER ><input type="button" class="button" name="btnEmailPro" style="width:150px;text-align=center" value="Gerar Arquivo " onclick ="javascript:EnviaEmail();"></td>
</TR>
</TABLE>
</FORM>
</BODY>
<%
	set objRSProv = nothing
%>

