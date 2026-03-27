<html>
<head>

<title>X.Sol Arte em Aquecedores - Visualiza Cliente</title>
<link rel="STYLESHEET" type="text/css" href="css/forAll.css">

</head>
<body>
<%
filepath=Server.MapPath("dados.mdb")
Set con = Server.CreateObject("ADODB.CONNECTION")

con.open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & filepath
set rs=Server.CreateObject("ADODB.Recordset")

sql = "SELECT * FROM Clientes WHERE Codigo ="&Request.QueryString("Codigo")&""
rs.Open sql,con
On Error Resume Next

%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div align="center"><img src="figuras/fg_logo.jpg"></div></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><strong>Informa&ccedil;&otilde;es 
        do Cliente</strong></font></div></td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="1" cellpadding="1">
  <tr>
    <td class="forTexts" width="20%" bgcolor="#FFFF66"><div align="right">NOME:</div></td>
	<td class="forTexts" width="80%"><%Response.Write(rs.Fields("Nome_Cli"))%></td>
  </tr>
  <tr>
    <td class="forTexts" align="right" width="20%" bgcolor="#FFFF66"><div align="right">ENDERE&Ccedil;O:</div></td>
    <td class="forTexts" align="left" width="80%"><%Response.Write(rs.Fields("End"))%></td>
  </tr>
  <tr>
    <td class="forTexts" align="right" width="20%" bgcolor="#FFFF66"><div align="right">BAIRRO:</div></td>
    <td class="forTexts" align="left" width="80%"><%Response.Write(rs.Fields("Bairro"))%></td>
  </tr>
  <tr>
    <td class="forTexts" align="right" width="20%" bgcolor="#FFFF66"><div align="right">CIDADE:</div></td>
    <td class="forTexts" align="left" width="80%"><%Response.Write(rs.Fields("Cidade"))%>&nbsp;</td>
  </tr>
  <tr>
    <td class="forTexts" align="right" width="20%" bgcolor="#FFFF66"><div align="right">CEP:</div></td>
    <td class="forTexts" align="left" width="80%"><%Response.Write(rs.Fields("Cep"))%>&nbsp;</td>
  </tr>  
  <tr>
    <td class="forTexts" align="right" width="20%" bgcolor="#FFFF66"><div align="right">ZONA:</div></td>
    <td class="forTexts" align="left" width="80%"><%Response.Write(rs.Fields("Zona"))%>&nbsp;</td>
  </tr>  
  <tr>
    <td class="forTexts" align="right" width="20%" bgcolor="#FFFF66"><div align="right">FONE 
        RESIDENCIAL:</div></td>
    <td class="forTexts" align="left" width="80%"><%Response.Write(rs.Fields("Fone_Res"))%>&nbsp;</td>
  </tr>  
  <tr>
    <td class="forTexts" align="right" width="20%" bgcolor="#FFFF66"><div align="right">FONE 
        COMERCIAL:</div></td>
    <td class="forTexts" align="left" width="80%"><%Response.Write(rs.Fields("Fone_Com"))%>&nbsp;</td>
  </tr>  
  <tr>
    <td class="forTexts" align="right" width="20%" bgcolor="#FFFF66"><div align="right">FONE 
        CELULAR:</div></td>
    <td class="forTexts" align="left" width="80%"><%Response.Write(rs.Fields("Fone_Celular"))%>&nbsp;</td>
  </tr>  
  <tr>
    <td class="forTexts" align="right" width="20%" bgcolor="#FFFF66"><div align="right">TIPO 
        DE MODELO EQUP.:</div></td>
    <td class="forTexts" align="left" width="80%"><%Response.Write(rs.Fields("Tipo_Eq_Mod"))%>&nbsp;</td>
  </tr>  
  <tr>
    <td class="forTexts" align="right" width="20%" bgcolor="#FFFF66"><div align="right">EMPRESA 
        SOLICITANTE:</div></td>
    <td class="forTexts" align="left" width="80%"><%Response.Write(rs.Fields("Solicitante"))%>&nbsp;</td>
  </tr>  

</table>

<%
rs.close
set rs = nothing
conn.close
set conn=nothing
%>

<%
filepath=Server.MapPath("dados.mdb")
Set con = Server.CreateObject("ADODB.CONNECTION")

con.open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & filepath
set rs=Server.CreateObject("ADODB.Recordset")

sql = "SELECT * FROM Atendimento WHERE Cod_Cliente ="&Request.QueryString("Codigo")&""
rs.Open sql,con
On Error Resume Next

%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>Informa&ccedil;&otilde;es de Atendimento</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="1">
  <tr>
    <td class="forTexts" bgcolor="#FFFF66"><div align="center">DT. ATENDIMENTO</div></td>
    <td class="forTexts" bgcolor="#FFFF66"><div align="center">DT. SERVI&Ccedil;O</div></td>
    <td class="forTexts" bgcolor="#FFFF66"><div align="center">SERVI&Ccedil;O</div></td>
    <td class="forTexts" bgcolor="#FFFF66"><div align="center">FICHA</div></td>
  </tr>
   <%if rs.EOF then
       Response.Write ("Não existe registro")
     Else
       rs.MoveFirst
     Do While Not rs.EOF
   %> 
  <tr>
    <td class="forTexts"><a href="consulta.asp"><%Response.Write(rs.Fields("Dt_Atend"))%>&nbsp;</a></td>
    <td class="forTexts"><%Response.Write(rs.Fields("Dt_Atend_Serv"))%>&nbsp;</td>
    <td class="forTexts"><%Response.Write(rs.Fields("Tipo_Serv"))%>&nbsp;</td>
    <td class="forTexts"><%Response.Write(rs.Fields("Ficha"))%>&nbsp;</td>
  </tr>
<%  
  rs.MoveNext  
Loop

End if
rs.Close
Set rs = Nothing

cn.Close
Set cn = Nothing
%>    
</table>

</body>
</html>