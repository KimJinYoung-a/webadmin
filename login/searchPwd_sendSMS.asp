<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��й�ȣã�� �������� �߼�
' History : 2016.10.15 ������ ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<%
dim sql
dim sCharge,susercell , userid,refip ,authNo, sendMsg,sKey
dim manager_hp,jungsan_hp,deliver_hp
userid 		= requestCheckVar(Request.Form("uid"),32)
sCharge 	= requestCheckVar(Request.Form("shp"),1)
refip			= request.ServerVariables("REMOTE_ADDR")
skey  		=	requestCheckVar(Request.Form("sKey"),32)

if userid ="" or sCharge ="" then
	 response.write("<script>alert('������ȣ ���� �Ķ���Ϳ� ������ �߻��߽��ϴ�. �����ڿ��� �������ּ���');</script>")
	 dbget.close() : response.end
end if

if skey <> md5(userid&"TPUSMS") then
	response.write("<script>alert('���̵� ��ȣȭ�� ������ �߻��߽��ϴ�. Ȯ�� �� �ٽ� �õ����ּ���');</script>")
	dbget.close() : response.end
end if

sql =" select id, manager_name, manager_hp ,jungsan_name, jungsan_hp, deliver_name, deliver_hp from db_partner.dbo.tbl_partner where id ='"&userid&"'"
	rsget.Open sql,dbget,1
  if  not rsget.EOF  then
   		manager_hp =rsget("manager_hp")
   		jungsan_hp =rsget("jungsan_hp")
   		deliver_hp =rsget("deliver_hp")
   end if
  rsget.close

 if sCharge = "J" then
 	susercell = jungsan_hp
elseif  sCharge = "D" then
 	susercell = deliver_hp
 else
 	susercell = manager_hp
end if

if IsNull(susercell) then susercell = ""
if susercell = "" then
	 response.write("<script>alert('�ڵ��� ��ȣ�� �������� �ʽ��ϴ�. Ȯ�� �� ��õ� ���ּ���');</script>")
	 dbget.close() : response.end
end if

'// ������ȣ �߻� �� DB���� �� SMS�߼�
Randomize()
authNo = int(Rnd()*1000000)		'6�ڸ� ����
authNo = Num2Str(authNo,6,"0","R")

'#���� ����
''Call SendNormalSMS(susercell,"","[�ٹ�����SCM] " & susername & "�� ������ȣ�� ["&authNo&"]�Դϴ�.")
Call SendNormalSMS_LINK(susercell,"","[�ٹ�����Partner] ��й�ȣã�� ������ȣ�� ["&authNo&"]�Դϴ�.")

'#�������
sql = "insert into db_partner.dbo.tbl_partner_searchPWD_authno( userid, refip,authno)"&vbcrlf
sql = sql & " values ('"&userid&"','"&refip&"','"&authNo&"')"
dbget.Execute sql

'//�߼� �ȳ� �� ī���� ����
	sendMsg= "<script language=javascript>"
	sendMsg = sendMsg &		" parent.document.all.dvAuth.style.display ='';"
	sendMsg = sendMsg &		"	parent.startLimitCounter('new');"
	sendMsg = sendMsg &		"	alert('�޴������� ������ȣ�� �߼��߽��ϴ�.\nSMS�� Ȯ�� �� �α������ּ���.');"
			IF application("Svr_Info")="Dev" THEN 		'// TEST�����̸� �׳� Alertó��
	sendMsg = sendMsg &	"	alert(' [�ٹ�����Partner] ��й�ȣã�� ������ȣ�� [" & authNo & "]�Դϴ�.');"
			end if
 sendMsg = sendMsg & "</script>"

response.write sendMsg
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
