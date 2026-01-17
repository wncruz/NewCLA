<!--#include file="../inc/data.asp"-->
<%
	dblNroSev = Request.Form("hdnNroSev")
	pro_id 	  = Request.Form("hdnCboProvedor")
	
	Vetor_Campos(1)="adInteger,4,adParamInput," & dblNroSev
	Vetor_Campos(2)="adInteger,4,adParamInput," & pro_id
	Vetor_Campos(3)="adInteger,4,adParamOutput,0"

	Call APENDA_PARAM("CLA_sp_check_sevMestra",3,Vetor_Campos)
	ObjCmd.Execute'pega dbaction
	DBAction = ObjCmd.Parameters("RET").value
	'response.write "<script>alert('"&DBAction&"')</script>"

	if DBAction <> 0 then
		Response.Write "<script language=javascript>parent.resposta(" & Cint("0" & DBAction) & ",'');</script>"
		Response.End
	else
	  %>
	  <script language="JavaScript">
	  	parent.AdicionarAcessoLista()
	  </script>
	  <%
	end if
%>
