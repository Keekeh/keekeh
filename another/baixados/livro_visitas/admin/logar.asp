<!--#include file = "config.asp"-->
<%
'Pega os dados digitados no formul�rio
usuario = request("usuario")
senha = request("senha")

'Verifica se os dados est�o corretos, se estiverem
if usuario = user and senha = password then
'Cria uma se��o com o nome do usu�rio
session("login") = true
'Redireciona ele para a administra��o
response.Redirect("admin.asp")
'Cao contr�rio
else
'Redireciona ele para a p�gina principal
response.Redirect("index.asp")
end if

%>