<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��ü ���� ���� ���� �߼�
' History : 2019.05.20 ������ ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"--> 
<!-- #include virtual="/lib/email/maillib.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<%
dim sql
dim sCharge,suserMail , userid,refip ,authNo, sendMsg,sKey, mailCont
dim manager_email ,jungsan_email, deliver_email
userid 		= requestCheckVar(Request.Form("uid"),32)
sCharge 	= requestCheckVar(Request.Form("shp"),1)
refip			= request.ServerVariables("REMOTE_ADDR")
skey  		=	requestCheckVar(Request.Form("sKey"),32)

if userid ="" or sCharge ="" then
	 response.write("<script>alert('������ȣ ���� �Ķ���Ϳ� ������ �߻��߽��ϴ�. �����ڿ��� �������ּ���');</script>") 
response.end
end if
 
if (LCASE(userid)<>LCASE(session("reauthUID"))) then
    response.write("<script>alert(' ��ȣȭ�� ������ �߻��߽��ϴ�. Ȯ�� �� �ٽ� �õ����ּ���');</script>") 
response.end
end if

if skey <> md5(userid&"TPUSMS") then
	response.write("<script>alert('���̵� ��ȣȭ�� ������ �߻��߽��ϴ�. Ȯ�� �� �ٽ� �õ����ּ���');</script>") 
response.end
end if

sql =" select id, email as manager_email ,jungsan_email, deliver_email from db_partner.dbo.tbl_partner where id ='"&userid&"'"
	rsget.Open sql,dbget,1
  if  not rsget.EOF  then 
   		manager_email =rsget("manager_email")  
   		jungsan_email =rsget("jungsan_email")  
   		deliver_email =rsget("deliver_email")  
   end if
  rsget.close
 
 if sCharge = "J" then 
 	suserMail = jungsan_email
elseif  sCharge = "D" then 
 	suserMail = deliver_email
 else
 	suserMail = manager_email
end if	

if suserMail = "" then
	 response.write("<script>alert('����� �̸����� �������� �ʽ��ϴ�. Ȯ�� �� ��õ� ���ּ���');</script>") 
response.end
end if
 
'// ������ȣ �߻� �� DB���� �� SMS�߼�
Randomize()
authNo = int(Rnd()*1000000)		'6�ڸ� ����
authNo = Num2Str(authNo,6,"0","R")

'#�������� ����

mailCont = "<h1><b>��û�Ͻ� <span style=""color:#f22727"">������ȣ</span>�� �˷��帳�ϴ�.</b></h1><p></p>"
mailCont = mailCont & "<p>�Ʒ� ������ȣ 6�ڸ��� �Է�â�� �Է����ּ���.<br />������ȣ�� ������ �߼۵Ǵ� �������� <b>10�а��� ��ȿ</b>�մϴ�.</p><br />"
mailCont = mailCont & "<table style=""border-top:2px solid #A0A0A0;border-bottom:2px solid #A0A0A0;text-align:center;"">"
mailCont = mailCont & "<tr style=""border-bottom:1px solid #E0E0E0;"">"
mailCont = mailCont & "<td style=""background-color:#F0F0F0;padding:5px;"">������ȣ</td>"
mailCont = mailCont & "<td><b>" & authNo & "</b></td></tr>"
mailCont = mailCont & "<tr>"
mailCont = mailCont & "<td style=""background-color:#F0F0F0;padding:5px;"">�߼۽ð�</td>"
mailCont = mailCont & "<td style=""padding:5px;"">" & now() & "</td>"
mailCont = mailCont & "</tr>"
mailCont = mailCont & "</table>"

'#�������
sql = "insert into db_partner.dbo.tbl_partner_searchPWD_authno( userid, refip,authno)"&vbcrlf
sql = sql & " values ('"&userid&"','"&refip&"','"&authNo&"')"
dbget.Execute sql

'# �̸��� �߼�
call sendmailCS(suserMail, "��û�Ͻ� ������ȣ�� �˷��帳�ϴ�.", mailCont)
'response.Write mailCont

'//�߼� �ȳ� �� ī���� ����  
sendMsg= "<script type=""text/javascript"">" 
sendMsg = sendMsg &		" parent.document.all.dvAuth.style.display ='';" 
sendMsg = sendMsg &		"	parent.startLimitCounter('newMail');"  
sendMsg = sendMsg &		"	alert('���� �̸����� �߼��߽��ϴ�.\n�̸����� Ȯ�� �� �α������ּ���.');"  
		IF application("Svr_Info")="Dev" THEN 		'// TEST�����̸� �׳� Alertó�� 
sendMsg = sendMsg &	"	alert(' [�ٹ�����Partner] ����ȯ��IP���� ������ȣ�� [" & authNo & "]�Դϴ�.');"  
		end if
sendMsg = sendMsg & "</script>" 
 
response.write sendMsg
%>