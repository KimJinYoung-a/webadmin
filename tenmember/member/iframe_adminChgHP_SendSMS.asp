<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ���� �޴��� ���� �������� �߼�
' History : 2013.02.18 ������ ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/email/smslib.asp"-->
<%
dim cMember
dim empNo, chgHp
dim sempno, susername
dim strSql, chkWait
dim authNo

empNo=requestCheckVar(Request("eno"),15)
chgHp=requestCheckVar(Request("chp"),16)

if (empNo="") then
	Call Alert_Move("�����ȣ�� �����ϴ�.","about:blank")
	Response.End
end if

if (chgHp="" or chgHp="--") then
	Call Alert_Move("�߼��� �޴�����ȣ�� �����ϴ�.","about:blank")
	Response.End
end if

'// ���� �⺻���� ����
Set cMember = new CTenByTenMember
	cMember.Fempno = empNo
	cMember.fnGetMemberData

	sempno   		= cMember.Fempno
	susername      	= cMember.Fusername

Set cMember = Nothing

if (sempno="" or isNull(sempno)) then
	Call Alert_Move("�߸��� �����ȣ�Դϴ�.\n�ý������� �������ּ���.","about:blank")
	Response.End
end if

'// ������ȣ �߼ۿ��� Ȯ��
strSql = "select count(idx) " &_
		" from db_log.dbo.tbl_partner_login_log " &_
		" where userid='" & sempno & "' " &_
		" 	and loginSuccess='S' " &_
		" 	and datediff(ss,regdate,getdate()) between 0 and 180"
rsget.Open strSql,dbget,1
	chkWait = rsget(0)>0
rsget.Close

if chkWait then
	Call Alert_Move("�̹� ������ȣ�� �߼��Ͽ����ϴ�.\n�޴����� SMS�� Ȯ�����ּ���.","about:blank")
	Response.End
end if

'// ������ȣ �߻� �� DB���� �� SMS�߼�
Randomize()
authNo = int(Rnd()*1000000)		'6�ڸ� ����
authNo = Num2Str(authNo,6,"0","R")

'#���� ����
'Call SendNormalSMS(chgHp,"","[�ٹ����پ���] " & susername & "�� ������ȣ�� ["&authNo&"]�Դϴ�.")
Call SendNormalSMS_LINK(chgHp,"","[�ٹ����پ���] " & susername & "�� ������ȣ�� ["&authNo&"]�Դϴ�.")
'#�α� ����
Call AddLoginLog (sempno,"S",authNo)

'//�߼� �ȳ� �� ī���� ����
IF application("Svr_Info")="Dev" THEN
	'// TEST�����̸� �׳� Alertó��
	Response.Write "<script language=javascript>" &_
			"	parent.startLimitCounter('new');" &_
			"	alert('" & susername & "�� ������ȣ�� [" & authNo & "]�Դϴ�.');" &_
			"</script>"
else
	Response.Write "<script language=javascript>" &_
			"	parent.startLimitCounter('new');" &_
			"	alert('�޴������� ������ȣ�� �߼��߽��ϴ�.\nSMS�� Ȯ�� �� �α������ּ���.');" &_
			"</script>"
end if

'-----------------------------------------------------------
'// ������ ���� �α� ���� �Լ�
Sub AddLoginLog(param1,param2,param3)
    dim sqlStr, reFAddr
    reFAddr = request.ServerVariables("REMOTE_ADDR")

    sqlStr = " insert into [db_log].[dbo].tbl_partner_login_log" + VbCrlf
	sqlStr = sqlStr + " (userid,refip,loginSuccess,USBTokenSn)" + VbCrlf
	sqlStr = sqlStr + " values(" + VbCrlf
	sqlStr = sqlStr + " '" + param1 + "'," + VbCrlf
	sqlStr = sqlStr + " '" + Left(reFAddr,16) + "'," + VbCrlf
	sqlStr = sqlStr + " '" + param2 + "'," + VbCrlf
	sqlStr = sqlStr + " '" + param3 + "'" + VbCrlf
	sqlStr = sqlStr + " )" + VbCrlf

    dbget.Execute sqlStr
end Sub
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->