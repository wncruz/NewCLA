<!--#include file="../inc/data.asp"-->
<%
strIDLogico = request.form("txtNroLogico")
strhdnAcao = request.form("hdnAcaoAPG")
strSolAcessoId = request.form("hdnSolAcessoID")

if len(strIDLogico) < 10 then
  response.write "<script>alert('Acesso Lógico Inválido.')</script>"
  Response.end
end if

Vetor_Campos(1)="adWchar,10,adParamInput, " & trim(strIDLogico)
Vetor_Campos(2)="adWChar,20,adParamInput, " & trim(strhdnAcao)
Vetor_Campos(3)="adInteger,10,adParamInput, " & strSolAcessoId
Vetor_Campos(4)="adInteger,4,adParamOutput,0"
Vetor_Campos(5)="adWChar,100,adParamOutput,0"

Call APENDA_PARAM("CLA_sp_check_IDLogico",5,Vetor_Campos)
'response.write APENDA_PARAMSTR("CLA_sp_check_IDLogico",5,Vetor_Campos)
ObjCmd.Execute'pega dbaction

DBErro = ObjCmd.Parameters("RET").value
DBErroDesc = ObjCmd.Parameters("RET1").value

if DBErro = 1 then
  response.write "<script>alert('"& trim(DBErroDesc) &"')</script>"
else
  %>
  <form name="form_confirma" method="post" action="ProcessoAssociarLogicoApg.asp">
    <input type="hidden" name="hdnAcessoLogico" value="<%=strIDLogico%>">
	<input type="hidden" name="hdnSolAcessoID" value="<%=strSolAcessoId%>">
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
