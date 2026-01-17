<script language="JavaScript"><!--
function Proxima(pag){
 var myform = document.Formfiltro;
 myform.Npagina.value = pag +1 ;  
 myform.filtro.value = 1;  
 myform.method = "post";
 myform.submit(); 

}

function Anterior(pag){
 var myform = document.Formfiltro; 
 myform.Npagina.value = pag - 1;  
 myform.filtro.value = 1;  
 myform.method = "post";
 myform.submit(); 

}

function mostrarfilro(mostrafiltro){
 var myform = document.Formfiltro;     
 myform.mostrafiltro.value=mostrafiltro;
 myform.method = "post";
 myform.submit(); 

}
// --></script>
