<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �޴���Ȯ���� ���� �޴�����ȣ ���� ó��
' History : 2013.02.18 ������ ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
dim cMember

dim empno, username, strSql
dim MobileNo1, MobileNo2, MobileNo3, MobileNo
dim manageUrl
dim NiceId, KeyString, ReturnURL, ConfirmMsg, strProcessType, strSendInfo, strOrderNo, SIKey

'// ���� �Ҵ�
empno = requestCheckVar(Request.form("empNo"),18)	' �����ȣ
MobileNo1 = requestCheckVar(Request.form("hpNum1"),3)	' �޴�����ȣ1
MobileNo2 = requestCheckVar(Request.form("hpNum2"),4)	' �޴�����ȣ2
MobileNo3 = requestCheckVar(Request.form("hpNum3"),4)	' �޴�����ȣ3

'// ���� �⺻���� ����
Set cMember = new CTenByTenMember
	cMember.Fempno = empno
	cMember.fnGetMemberData

	username      	= cMember.Fusername
Set cMember = Nothing

if username="" or isNull(username) then
    Call Alert_close("���������� �������� �ʽ��ϴ�.")
    response.end
end if

IF application("Svr_Info")="Dev" THEN
 	manageUrl 	    = "http://testwebadmin.10x10.co.kr"
 ELSE
 	manageUrl 	    = "http://webadmin.10x10.co.kr"
 END IF

	MobileNo = MobileNo1 + "-" + MobileNo2 + "-" + MobileNo3

	strSql = "Update db_partner.dbo.tbl_user_tenbyten " &_
			" Set usercell='" & MobileNo & "'" &_
			"	, isIdentify='Y' " &_
			" Where empno='" & CStr(empno) & "'"
	dbget.Execute(strSql)
%>
	<script language="javascript">
	alert('����Ȯ�� �� �Է��Ͻ� �޴�����ȣ�� ����Ǿ����ϴ�.');
	parent.opener.history.go(0);
	parent.close();
	</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->