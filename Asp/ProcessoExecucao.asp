<%
'•IMPLEMENT SOFT - IMPLEMENTAÇÕES E SOLUÇÕES EM INFORMÁTICA
'	- Sistema			: CLA
'	- Arquivo			: ProcessoExecucao.asp
'	- Descrição			: Executa um Tronco/Par
%>
<!--#include file="../inc/data.asp"-->
<%
Function GravarExecucao()

	Set rec = db.execute("CLA_sp_view_recurso " & request("recurso"))
	agora = year(now) & "-" & right("00" & month(now),2) & "-" & right("00" & day(now),2) & " " & hour(now) & ":" & minute(now) & ":" & second(now) & ".000"
	redim Int_ID(9999)
	DBAction = 0
	for i = 1 to request("qtd_fac")
		set inter = db.execute("CLA_sp_sel_interorigem '" & request("txtCoordenada"&i) & "'," & rec("Esc_id") & "," & rec("Dst_id") & "," & request("facilidade"&i))
		DBAction = inter("ret")
		if DBAction = 0 then
			for x = 1 to (i - 1)
				if cdbl(Int_ID(x)) = cdbl(inter("Int_ID")) then
					DBAction = 104 'A MESMA COORDENADA NÃO PODE SER UTILIZADA EM PARES DIFERENTES
					exit for
				end if
			next
			Int_ID(i) = inter("Int_ID")
		end if
		if DBAction <> 0 then exit for
	Next
	if DBAction = 0 then
		for i = 1 to Request("qtd_fac")
				Vetor_Campos(1)="adInteger,10,adParamInput," & request("facilidade"&i)
				Vetor_Campos(2)="adInteger,10,adParamInput," & Int_ID(i)
				Vetor_Campos(3)="adWChar,300,adParamInput,"& request("posicaoobservacao"&i)
				Vetor_Campos(4)="adWchar,1,adParamInput," & Request("hdnOrigem")
				Vetor_Campos(5)="adInteger,2,adParamOutput,0"
				Call APENDA_PARAM("CLA_sp_execucao",5,Vetor_Campos)
				ObjCmd.Execute
				DBAction = ObjCmd.Parameters("RET").value
		Next
	End if 'DBAction = 0

	GravarExecucao = DBAction

End Function

Select Case Trim(Request.Form("hdnAcao"))

	Case "GravarExecucao"
		DBAction = GravarExecucao()
		if DBAction = 69 then
			Response.Write "<script language=javascript>parent.ExecucaoGravada(" & DBAction & ");</script>"
		Else
			Response.Write "<script language=javascript>parent.resposta("& DBAction & ",'');</script>"
		End if	
		
End Select
%>
