<%
response.end
    Dim connString : connString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=testuser;Initial Catalog=myDatabase;Data Source=myDBServer;Password=myPassword"
    Response.Write("connString : " & connString & "<br /><br />")

    Dim connectionString 
    Set connectionString = Server.CreateObject("TenCrypto.ConnectionString")

    Dim encryptedConnString : encryptedConnString = connectionString.EncryptString(connString)
    Response.Write("encryptedConnString : " & encryptedConnString & "<br /><br />")

    Dim filePath : filePath = "D:\www\crypto\connString.conn"
    Response.Write("filePath : " & filePath & "<br /><br />")    

    Dim decryptedConnString : decryptedConnString = connectionString.DecryptString(filePath)
    Response.Write("decryptedConnString : " & decryptedConnString & "<br /><br />")   

    Set connectionString = Nothing

    Response.Write("global.asa -> Application(""defaultConnString"") : " & Application("defaultConnString"))
%>
<br/>
<br/>
-----------------------------------------------��ȣȭ-----------------------
<br/>
<br/>
<%
    Dim myString : myString = "�׽�Ʈ~1345�Ϻ�ȣ�ؽ�"
    Response.Write("���� : " &  myString & "<br />")
    Dim crypto
    Set crypto = Server.CreateObject("TenCrypto.Crypto")
    
    Response.Write("MD5Hashing : " &  crypto.MD5Hashing(myString) & "<br />")
    Response.Write("SHA1Hashing : " &  crypto.SHA1Hashing(myString) & "<br />")
    Response.Write("SHA256Hashing : " &  crypto.SHA256Hashing(myString) & "<br />")
    Response.Write("SHA512Hashing : " &  crypto.SHA512Hashing(myString) & "<br />")

    Dim key : key = "aa12331234523433455sss5235453334"

    '������Ʈ���� ���Ե� Ű ��� SymmetricAlgorithmType.Rijndael
    Dim encryptedString : encryptedString = crypto.EncryptString(myString) 
    Response.Write("EncryptString : " &  encryptedString & "<br />")
    Response.Write("DecryptString : " &  crypto.DecryptString(encryptedString) & "<br />")

    Dim encryptedAesString : encryptedAesString =  crypto.EncryptAesString(myString, key)
    Response.Write("EncryptAesString : " &  encryptedAesString & "<br />")
    Response.Write("DecryptAesString : " &  crypto.DecryptAesString(EncryptedAesString, key) & "<br />")

    Dim encryptDesString : encryptDesString =  crypto.EncryptDesString(myString, key)
    Response.Write("EncryptDesString : " &  encryptDesString & "<br />")
    Response.Write("DecryptDesString : " &  crypto.DecryptDesString(encryptDesString, key) & "<br />")

    Dim encryptTripleDesString : encryptTripleDesString =  crypto.EncryptTripleDesString(myString, key)
    Response.Write("EncryptTripleDesString : " &  encryptTripleDesString & "<br />")
    Response.Write("DecryptTripleDesString : " &  crypto.DecryptTripleDesString(encryptTripleDesString, key) & "<br />")
        

    Set crypto = Nothing    
%>
