<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->

<%
function sendmailbankokNew(mailto,userName,orderserial) ' �Ա�Ȯ�θ���
        dim sql,discountrate
        dim mailfrom, mailtitle, mailcontent
        dim fs,objFile,dirPath,fileName

        mailfrom = "�ٹ�����<customer@10x10.co.kr>"
        mailtitle = "������ �Ա��� ���������� ó�� �Ǿ����ϴ�!"

        ' ������ �ҷ��ͼ�
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")
        'fileName = dirPath&"\\email_bank2011.htm"
        fileName = dirPath&"\\email_new_bank.html"
        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
        mailcontent = replace(mailcontent,":USERNAME:",userName)
        mailcontent = replace(mailcontent,":ORDERSERIAL:",orderserial)

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
end function

'call sendmailbankokNew("kjy8517@10x10.co.kr","������","11041731871")
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->