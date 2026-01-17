
<% 
Session.LCID = 1046 
ContPagina 	= 50
server.ScriptTimeout = 90

Function Converte_inteiro(numero,default)
		  if numero<>"" then
		     numero = cint(numero)
		  else 
		     numero = default   
		  end if	
		  Converte_inteiro = numero 
end function

Function Converte_inteiroLongo(numero,default)
		  if numero<>"" then
		     numero = clng(numero)
		  else 
		     numero = default   
		  end if	
		  Converte_inteiroLongo = numero 
end function

Function define_pagina(pagina,TotalPaginas)
		IF pagina="" then 
      			Npagina=1 
   		ELSE
      		IF cint(pagina)<1 then
        		Npagina=1 
     	    ELSE
         		IF cint(pagina)> TotalPaginas then 
            			Npagina = TotalPaginas 
         	 	ELSE
            			Npagina=pagina
         	 	END IF
      		END IF
   		END IF
        define_pagina = Npagina
end function
%>