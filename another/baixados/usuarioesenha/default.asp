<%
'---------------------------------------------------------------------------------------
'		Script by Fabio Franco
'	Email: fabio_franco@ofm.com.br ICQ: 164613668
'---------------------------------------------------------------------------------------

Dim SrtMsg, erro

SrtMsg = cstr(Request.QueryString("erro"))
Select Case SrtMsg

Case "1"
erro = "Digite usu�rio"

Case "2"
erro = "Digite senha"

Case "3"
erro = "Nome de usu�rio n�o encontrado."

Case "4"
erro = "Senha incorreta."

Case "5"
erro = "Voc� n�o est� logado."

end select
%>

<html><head><title>Teste em ASP com DB</title></head>

Digite abaixo seu login e senha para acessar a �rea restrita:<br>

<hr size="1" color="black">

<%if len(SrtMsg) <> 0 then%><br><font color="red"><%=erro%></font><br><%end if%>

<form action="verificar_usuario.asp" method="POST">

<table><tr><td>Usu�rio:</td><td align="right"><input type="text" name="usuario"></td></tr><tr><td>Senha:</td><td align="right"><input type="password" name="senha"></td></tr>

<tr><td align="center" colspan="2"><input type="submit" value="login"></td></tr></table>

<hr size="1" color="black">

</form></body></html>
<%
'---------------------------------------------------------------------------------------
'		Script by Fabio Franco
'	Email: fabio_franco@ofm.com.br ICQ: 164613668
'---------------------------------------------------------------------------------------
%>