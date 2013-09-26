<%
' Ajuste o timeout conforme o tamanho 
' dos arquivos que você irá aceitar
server.scripttimeout = 5400

'Caminho virtual onde os arquivos devem ser salvos no servidor
filePathDefault = "displan@displansaude.com./"

const ForWriting = 2
const TristateTrue = -1
crlf = chr(13) & chr(10)

' Esta função recupera o nome do campo do formulário
function getFieldName( infoStr)
    sPos = inStr( infoStr, "name=")
    endPos = inStr( sPos + 6, infoStr, chr(34) & ";")
    if endPos = 0 then
 endPos = inStr( sPos + 6, infoStr, chr(34))
    end if
    getFieldName = mid( infoStr, sPos + 6, endPos - (sPos + 6))
end function

' Esta função recupera o nome do arquivo
function getFileName( infoStr)
    sPos = inStr( infoStr, "filename=")
    endPos = inStr( infoStr, chr(34) & crlf)
    getFileName = mid( infoStr, sPos + 10, endPos - (sPos + 10))
end function

' Esta função recupera o tipo mime do arquivo
function getFileType( infoStr)
    sPos = inStr( infoStr, "Content-Type: ")
    getFileType = mid( infoStr, sPos + 14)
end function

' Lê o arquivo o arquivo e tudo mais que foi enviado
postData = ""
Dim biData
biData = Request.BinaryRead(Request.TotalBytes)
' Como o arquivo é binário, iremos armazená-lo
' numa variável mais controlável
for nIndex = 1 to LenB( biData)
    postData = postData & Chr(AscB(MidB( biData, nIndex, 1)))
next

' Como os dados foram lidos de uma forma binária (BinaryRead), a instrução Request.Form 
' não está mais disponível. Portanto, será necessário processar as variáveis manualmente.
' Primeiro, será verificado o tipo de codificação.
contentType = Request.ServerVariables( "HTTP_CONTENT_TYPE")
ctArray = split( contentType, ";")
' O envio de arquivos somente funciona com a codificação
' "multipart/form-data", por isso vamos checar se esse foi a codificação usada.
if trim(ctArray(0)) = "multipart/form-data" then
    errMsg = ""
    ' Armazene o conteúdo do formulário ...
    bArray = split( trim( ctArray(1)), "=")
    boundry = trim( bArray(1))
    ' ... divida as variáveis
    formData = split( postData, boundry)
    ' agora, nós precisamos extrair a informação de cada variável.
    dim myRequest, myRequestFiles(9, 3) 
    Set myRequest = CreateObject("Scripting.Dictionary")
    fileCount = 0
    for x = 0 to ubound( formData)
 ' duas marcas "crlf" indicam o final da informação sobre campo
 ' tudo depois dele é o valor
 infoEnd = instr( formData(x), crlf & crlf)
 if infoEnd > 0 then
     ' lê as informações do campo...
     varInfo = mid( formData(x), 3, infoEnd - 3)
     ' lê o valor do campo
     varValue = mid( formData(x), infoEnd + 4, len(formData(x)) - infoEnd - 7)
     ' verifica se ele é um arquivo
     if (instr( varInfo, "filename=") > 0) then
   myRequestFiles( fileCount, 0) = getFieldName( varInfo)
   myRequestFiles( fileCount, 1) = varValue
   myRequestFiles( fileCount, 2) = getFileName( varInfo)
   myRequestFiles( fileCount, 3) = getFileType( varInfo)
   fileCount = fileCount + 1
     else
   ' outro tipo de campo
   myRequest.add getFieldName( varInfo), varValue
     end if
 end if
    next
else
    errMsg = "Tipo de codificação inválido!"
end if 

' Salva o arquivo enviado
set lf = server.createObject( "Scripting.FileSystemObject")
if myRequest("filename") = "original" then
    ' Usa o nome do arquivo original
    browserType = UCase( Request.ServerVariables( "HTTP_USER_AGENT"))
    if (inStr(browserType, "WIN") > 0) then
 ' Sistema Windows
 sPos = inStrRev( myRequestFiles( 0, 2), "\")
 fName = mid( myRequestFiles( 0, 2), sPos + 1)
    end if
 ' Sistema Mac
    if (inStr(browserType, "MAC") > 0) then
 fName = myRequestFiles(0, 2)
    end if
    ' Caminho onde os arquivos devem ser salvos
    filePath = filePathDefault & fName
else
    ' use o nome especificado pelo usuário
    ' Caminho do arquivo
    filePath = filePathDefault & myRequest("userSpecifiedName")
end if
savePath = server.mapPath( filePath)
set saveFile = lf.createtextfile(savePath, true)
saveFile.write( myRequestFiles(0, 1))
saveFile.close
%>

<html>
<body>
<% if errMsg = "" then %>
    Arquivo enviado com sucesso!
<%else %>
    Erro no envio do arquivo!<br>
    <%= errMsg %>
<% end if %>
</body>
</html>
&lt;/html&gt;
</body>
</html>
