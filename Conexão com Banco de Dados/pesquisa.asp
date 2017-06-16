<%
	palavra = Request("palavra")
	
	ConnString="Provider=Microsoft.Jet.OLEDB.4.0; Data Source=C:\Inetpub\wwwroot\flash\pesquisa.mdb;"
	Set conexao = Server.CreateObject("ADODB.Connection")
	conexao.Open ConnString
	Set registros = Server.CreateObject("ADODB.Recordset")
	registros.Open "clientes", conexao

	set registros = conexao.execute(" select * from clientes where nome like '%"&(palavra)&"%' order by nome ")

	if registros.eof Then
		response.write "encontrado=False"
	else
		response.write "encontrado=True&resultado="
		while not registros.eof
			resultado = resultado & "NOME : " & registros("nome")      & vbCr        	'pula linha
			resultado = resultado & "TEL. : " & registros("telefone")  & vbCr	     	'pula linha
			resultado = resultado & "RG   : " & registros("rg")        & vbCr        	'pula linha
			resultado = resultado & "CPF  : " & registros("cpf")       & vbCr & vbCr  	'pula linha
			registros.MoveNext
		wend
	End If

	response.write server.URLEncode(resultado)

	conexao.close
	set conexao = nothing
%>