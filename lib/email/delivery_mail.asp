<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<%
' http://www.10x10.co.kr/lib/email/delivery_mail_test.asp?orderserial=02033000019&deliverno=1635753685
        dim sql,discountrate,subtotalprice,come
        dim mailfrom, mailto, mailtitle, mailcontent,orderserial,deliverno
 
        orderserial = request.form("orderserial")
        deliverno = request.form("deliverno")
        come = request.form("come")
		
		if(orderserial="") then
            response.write("�ֹ���ȣ�� �Ѿ�����ʾҽ��ϴ�.")
            dbget.close()	:	response.End
        end if
        
		if (come="comemail") then
			sendmailcome orderserial
			response.write "S_OK"
		end if
        
        if(deliverno="") then
            response.write("�ù� ������ȣ�� �Ѿ�����ʾҽ��ϴ�.")
            dbget.close()	:	response.End
        end if

        mailcontent = sendmailfinish(orderserial,deliverno)

'        response.write mailcontent

        response.write "S_OK"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->