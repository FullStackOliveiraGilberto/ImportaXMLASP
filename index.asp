<%
'Instancia o objeto XMLDOM.
'Set objXMLDoc = CreateObject("MSXML2.DOMDocument.4.0")
 starttime = Timer() 
 set objXMLDoc = server.createObject("Microsoft.XMLDOM")
 
'Indicamos que o download em segundo plano n�o � permitido
objXMLDoc.async = False
 
'Carrega o domcumento XML
objXMLDoc.load("C:\temp\ProjetoGilberto\Salomao_farmacias\farmadelivery.xml")
 
'Carrega o domcumento XML
'Para quem possui servi�o de REVENDA, utilize este caminho
'objXMLDoc.load("E:\vhosts\DOMINIO_COMPLETO\httpdocs\internet.xml")
 
'O m�todo parseError cont�m informa��es sobre o �ltimo erro ocorrido
if objXMLDoc.parseError <> 0 then
 
response.write "C�digo do erro: " & objXMLDoc.parseError.errorCode & "<br>"
response.write "Posi��o no arquivo: " & objXMLDoc.parseError.filepos & "<br>"
response.write "Linha: " & objXMLDoc.parseError.line & "<br>"
response.write "Posi��o na linha: " & objXMLDoc.parseError.linepos & "<br>"
response.write "Descri��o: " & objXMLDoc.parseError.reason & "<br>"
response.write "Texto que causa o erro: " & objXMLDoc.parseError.srcText & "<br>"
response.write "Url do arquivo com problemas: " & objXMLDoc.parseError.url
 
else
 
'A propriedade documentElement refere-se � raiz do documento
Set raiz = objXMLDoc.documentElement
 
'Looping para percorrer todos os elementos filhos
For i = 0 to raiz.childNodes.length -1
 
'A propriedade NodeName cont�m o nome do elemento
'A propriedade childNodes cont�m a lista de
'elementos filhos
Response.write raiz.NodeName & "<br>" & _
             "oferta_id: " & raiz.childNodes.item(i).childNodes.item(0).text & "<br>" &_  
			 "oferta_descricao: " &  raiz.childNodes.item(i).childNodes.item(1).text & "<br>" &_
			 "empresa_id " &  raiz.childNodes.item(i).childNodes.item(2).text & "<br>" &_
			 "oferta_valor " & raiz.childNodes.item(i).childNodes.item(3).text & "<br>" &_
			 "link_produto " &  raiz.childNodes.item(i).childNodes.item(4).text & "<br>"&_
			 "oferta_imgproduto: " &   raiz.childNodes.item(i).childNodes.item(5).text & "<br>"&_ 
			 "oferta_principio_ativo: " &   raiz.childNodes.item(i).childNodes.item(6).text & "<br>"&_
			 "oferta_codigo_ms: " &   raiz.childNodes.item(i).childNodes.item(7).text & "<br>"&_
			 "ean: " &   raiz.childNodes.item(i).childNodes.item(8).text & "<br><br>"
Next
 
end if
 
'Destruindo os objetos
Set objXMLDoc = Nothing
Set raiz = Nothing

endtime = Timer() 
'Mostramos os resultados obtidos. 
Response.Write "O carregamento se completou em " & endtime-starttime & " segundos = " 
Response.Write " (" & (endtime-starttime)*1000 & " mil�simos de segundos)." 

%>
