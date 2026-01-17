<%
'Log de acesso ao CLA:
var_data = date()
var_hora = time()
var_ip = Request.ServerVariables("REMOTE_ADDR")
var_origem = Request.ServerVariables("Path_info")
var_tipo = "I" 'I/O

if session("ss_cla_log") = "" then
  Vetor_Campos(1)="adWChar,30,adParamInput,"&	strLoginRede 'Login
  Vetor_Campos(2)="adWChar,15,adParamInput,"&	var_ip 'IP
  Vetor_Campos(3)="adWChar,80,adParamInput,"&	var_origem 'Origem
  Vetor_Campos(4)="adWChar,1,adParamInput,"&	var_tipo 'Tipo
  Call APENDA_PARAM("CLA_sp_ins_LogAcesso",4,Vetor_Campos)
  ObjCmd.Execute
    
  session("ss_cla_log") = "registrado"
end if
'Log de acesso FIM:
%>