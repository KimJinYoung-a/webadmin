<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/email/mailLib2.asp"-->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<%
' http://www.10x10.co.kr/lib/email/bank_ok_mail.asp?buyemail=lcseung@yahoo.com&buyname=��ö��
        dim sql,discountrate,subtotalprice
        dim mailfrom, mailto, mailtitle, mailcontent,buyemail,buyname
 
        buyemail = request.form("buyemail")
        buyname = request.form("buyname")
'        buyemail = request("buyemail")
'        buyname = request("buyname")

        if(buyemail="") then
            response.write("�ֹ��� �̸����� �Ѿ���� �ʾҽ��ϴ�.")
            dbget.close()	:	response.End
        end if
        if(buyname="") then
            response.write("�ֹ��� �̸��� �Ѿ���� �ʾҽ��ϴ�.")
            dbget.close()	:	response.End
        end if

        mailcontent = sendmailbankok(buyemail,buyname)

'        response.write mailcontent

        response.write "S_OK"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->