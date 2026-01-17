<!--#include file="../inc/data.asp"-->
<%
strIDLogico = request.form("txtNroLogico")
strhdnAcao = request.form("hdnAcaoAprov")
strOrisolID = request.form("hdnOriSolID")
strAprovisiId = request.form("hdnAprovisiID")

if (len(strIDLogico) < 10) Or (strhdnAcao="DES" and left(strIDLogico,3)="677") then
  response.write "<script>alert('Acesso Lógico Inválido.')</script>"
  Response.end
end if

Vetor_Campos(1)="adWchar,10,adParamInput, " & trim(strIDLogico)
Vetor_Campos(2)="adWChar,20,adParamInput, " & trim(strhdnAcao)
Vetor_Campos(3)="adInteger,10,adParamInput, " & strAprovisiId
Vetor_Campos(4)="adInteger,4,adParamOutput,0"
Vetor_Campos(5)="adWChar,100,adParamOutput,0"

Call APENDA_PARAM("CLA_sp_check_IDLogico",5,Vetor_Campos)
'response.write APENDA_PARAMSTR("CLA_sp_check_IDLogico",5,Vetor_Campos)
ObjCmd.Execute'pega dbaction

DBErro = ObjCmd.Parameters("RET").value
DBErroDesc = ObjCmd.Parameters("RET1").value

if DBErro > 0 then
  response.write "<script>alert('"& trim(DBErroDesc) &"')</script>"
else
  %>
  <form name="form_confirma" method="post" action="ProcessoAssociarLogico.asp">
    <input type="hidden" name="hdnAcessoLogico" value="<%=strIDLogico%>">
	<input type="hidden" name="hdnAprovisiID" value="<%=strAprovisiID%>">
	<input type="hidden" name="hdnOrisolID" value="<%=strOriSolID%>">
  </form>
  <script language="VBscript">
    returnvalue=MsgBox ("Confirma associação do Acesso Lógico ao Item de OE?",36,"Confirmação definitiva de associação de lógico.")
                
    If returnvalue=6 Then
      form_confirma.submit()  
    Else
                        
    End If
  </script>
  <%
end if
%>
