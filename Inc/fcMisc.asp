<%

function getUser
  dim user
    
  If Request.ServerVariables("LOGON_USER") = "" Then
    getUser = "CONSULTA"    

  else 
    user = Request.ServerVariables("LOGON_USER")
    user = ucase(mid(user,10,10))
    Session.Contents("USER") = user
    getUser = user
  End If 
end function

function TIRAENTER(texto)
	IF NOT ISNULL(texto) THEN
		TIRAENTER = replace(trim(texto),chr(10),"")
		TIRAENTER = replace(TIRAENTER,chr(13),"")
	ELSE
		TIRAENTER =""
	END IF 

end function 

FUNCTION TIRAPLIC(TEXTO)
	IF NOT ISNULL(TEXTO) THEN
		TIRAPLIC = REPLACE(TEXTO,"'","")
	ELSE
		TIRAPLIC = ""
	END IF 

END FUNCTION 


%>