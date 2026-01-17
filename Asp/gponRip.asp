<!--#include file="../inc/data.asp"-->

<% 
 
dim strProEmail, strProNome, strParmProc
dim strServico, strAssunto

dim strFromEmail 	
dim	data			
dim	cliente 		
dim	endereco		
	'OE				
dim	designacao		
dim	velocidade		
dim	FabricanteONT	
dim	PE				
dim	ModeloONT		
dim	PortaPE		    
dim	DesignacaoONT	
dim	VLAN			
dim	PortaONT		
dim	Estacao		


set objXmlDoc = server.CreateObject("Microsoft.XMLDOM") 
	
	'Atribuição de valores para as variáveis 	
	
objXmlDoc.load(Request)
strCaminho = server.MapPath("..\")


set ndAcl  =  objXmlDoc.selectSingleNode("//acl")


dblAcessoLogico  = ndAcl.Text

	
Vetor_Campos(1)="adInteger,2,adParamInput,null" 
Vetor_Campos(2)="adInteger,2,adParamInput," & dblAcessoLogico
strSqlRet = APENDA_PARAMSTR("CLA_sp_sel_FacilidadeGPON",2,Vetor_Campos)

Set objRS = db.Execute(strSqlRet)

if Not objRS.Eof and Not objRS.bof then
	
	strFromEmail 	= "acessosrj@embratel.com.br"
	strProEmail 	= "CCRRJDC@embratel.com.br"
	strProNome		= "EMBRATEL - GPON"
	strServico		= "Ativação"
	
	data			= date()
	cliente 		= Trim(objRS("cli_nome"))
	endereco		= Trim(objRS("Endereco"))
	OE				= Trim(objRS("OE"))
	designacao		= Trim(objRS("Acl_DesignacaoServico"))
	velocidade		= Trim(objRS("vel_desc"))
	FabricanteONT	= Trim(objRS("Font_Nome"))
	PE				= Trim(objRS("OntVlan_PE"))
	ModeloONT		= Trim(objRS("Tont_Modelo"))
	PortaPE		    = Trim(objRS("OntVlan_PortaOLT"))
	DesignacaoONT	= Trim(objRS("Ont_Desig"))
	VLAN			= Trim(objRS("OntVlan_Nome"))
	PortaONT		= Trim(objRS("ONTPorta_Porta")) 
	Estacao			= Trim(objRS("est_config"))
	strAssunto		= "REAL IP - GPON " &  designacao & " - " & cliente 
	dblProId		= Trim(objRS("pro_id"))
	dblPedId		= Trim(objRS("ped_id"))
	intTipoProcesso	= Trim(objRS("Tprc_id"))
	if isnull(Trim(objRS("ped_id"))) then
		dblPedId		= Trim(objRS("ped_id_fisico"))
	end if 
	strParmProc 	= dblProId + "|" + dblPedId  + "|" + intTipoProcesso + "|" + " " 
	
end if 


DIM   strXML
	
	strXML = "<?xml version = ""1.0"" encoding =""ISO-8859-1""?><root>"
	strXML = strXML & "<arquivo>GponRip</arquivo>"
	strXML = strXML & "<parmproc>" & strParmProc & "</parmproc>"
	strXML = strXML & "<assunto>" & strAssunto & "</assunto>"
	strXML = strXML & "<proemail>" & strProEmail & "</proemail>"
	strXML = strXML & "<mailfrom>" & strFromEmail &  " </mailfrom>"
	strXML = strXML & "<pronome>"  & strProNome & "</pronome>"
	
	
	strXML = strXML &	"<numero>"
	
	strXML = strXML & "</numero>"
	
	strXML = strXML & "<familia>"
		
	strXML = strXML & "<data>" & data & "</data>"
	
	strXML = strXML & "</familia>"
	strXML = strXML & "<empresa>"


	strXML = strXML & "<cliente> " & cliente & "</cliente>" 
	strXML = strXML & "<endereco> " & endereco & "</endereco>" 
	strXML = strXML & "<OE> " & OE & "</OE>" 
	strXML = strXML & "<designacao> "& designacao & "</designacao>" 
	strXML = strXML & "<velocidade> "& velocidade & "</velocidade>" 
	
	strXML = strXML & "</empresa>"
	
	strXML = strXML & "<tecnico>"
	strXML = strXML & "<FabricanteONT>"& FabricanteONT &"</FabricanteONT>"	
	strXML = strXML & "<PE>"& PE &"</PE>"
	strXML = strXML & "<ModeloONT>"& ModeloONT &"</ModeloONT>"
	strXML = strXML & "<PortaPE>"& PortaPE &"</PortaPE>"
	strXML = strXML & "<DesignacaoONT>"&DesignacaoONT&"</DesignacaoONT>" 
	strXML = strXML & "<VLAN>"&VLAN&"</VLAN>"
	strXML = strXML & "<PortaONT>" & PortaONT & "</PortaONT>"
	strXML = strXML & "<Estacao>"&Estacao&"</Estacao>" 
		
	strXML = strXML & "</tecnico>"	

	strXML = strXML & "</root>"

	set objRS = nothing 
		
	Response.ContentType = "text/HTML;charset=ISO-8859-1"
	Response.Write (strXML)

%>