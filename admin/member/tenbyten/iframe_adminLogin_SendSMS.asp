<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �α��� �������� �߼�
' History : 2011.06.13 ������ ����
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
dim userid, lstp
dim sempno, susername, susercell, isIdentify, statediv, partnerusing
dim strSql, chkWait
dim authNo
dim retireday
userid=requestCheckVar(Request("uid"),32)
lstp  =requestCheckVar(Request("lstp"),10)  ''C �ΰ�� mobileApp ���� /2013/06/19 �߰�, W�ΰ�� ����Ʈ �ֹ���
if (lstp<>"C") and (lstp<>"W") then lstp="S"                '' �⺻�� S (���� SMS ���� �α���)
 
'// ���� �⺻���� ����
Set cMember = new CTenByTenMember
	cMember.Fuserid = userid
	cMember.fnGetScmMyInfo

	sempno   		= cMember.Fempno
	susername      	= cMember.Fusername
	susercell      	= cMember.Fusercell
	isIdentify		= cMember.FisIdentify
    
    statediv        = cMember.Fstatediv  ''2017/06/17 �߰�
    partnerusing    = cMember.Fpartnerusing
    retireday         = cMember.Fretireday
Set cMember = Nothing
 
'// ����Ȯ�������� �޾Ҵ��� Ȯ�� (�ȹ޾����� ����Ȯ�� �޴�����ȣ ���� �˾� ����) 
'if (isIdentify<>"Y") or (statediv<>"Y" and datediff("d",date(),retireday)<0 ) or (partnerusing<>"Y") then
if (isIdentify<>"Y") or (statediv<>"Y" and datediff("d",date(),retireday)<0 )   then
	Response.Write "<script language=javascript>parent.PopChgHPNum('"&retireday&"'); self.location ='about:blank';</script>"
	Response.End
end if



'// �޴�����ȣ ���� Ȯ��
if susercell="" or isNull(susercell) then
	Call Alert_Move("ȸ�� ������ �޴��� ��ȣ�� �����ϴ�.\nUSBŰ�� ����Ͽ� �α��� �� �޴��������� �Է����ּ���.","about:blank")
	Response.End
end if

'// ������ȣ �߼ۿ��� Ȯ��
strSql = "select count(idx) " &_
		" from db_log.dbo.tbl_partner_login_log " &_
		" where userid='" & userid & "' " &_
		" 	and loginSuccess='"&lstp&"' " &_
		" 	and datediff(ss,regdate,getdate()) between 0 and 180"
		
rsget.Open strSql,dbget,1
	chkWait = rsget(0)>0
rsget.Close

if chkWait then
    if (lstp="W") then
        response.write "<script>parent.jsSetStep(2);alert('�̹� ������ȣ�� �߼��Ͽ����ϴ�.\n�޴����� SMS�� Ȯ�����ּ���.'); self.location ='about:blank';</script>"
    else
    	response.write "<script>parent.jsSetStep(2);parent.startLimitCounter();</script>"
	    Call Alert_Move("�̹� ������ȣ�� �߼��Ͽ����ϴ�.\n�޴����� SMS�� Ȯ�����ּ���.","about:blank")
    end if
	Response.End
end if

'// ������ȣ �߻� �� DB���� �� SMS�߼�
Randomize()
authNo = int(Rnd()*1000000)		'6�ڸ� ����
authNo = Num2Str(authNo,6,"0","R")

'#���� ����
'Call SendNormalSMS(susercell,"","[�ٹ����پ���] " & susername & "�� ������ȣ�� ["&authNo&"]�Դϴ�.")
'Call SendNormalSMS_LINK(susercell,"","[�ٹ����پ���] " & susername & "�� ������ȣ�� ["&authNo&"]�Դϴ�.")
Call SendKakaoMsg_LINK(susercell,"","S0001","[����] " & susername & "���� ������ȣ�� ["&authNo&"]�Դϴ�.","SMS","","","")

'#�α� ����
Call AddLoginLog (userid,lstp,authNo)

'//�߼� �ȳ� �� ī���� ����
IF application("Svr_Info")="Dev" THEN
	'// TEST�����̸� �׳� Alertó��
	Response.Write "<script language=javascript>" &_
			"	parent.startLimitCounter('new');" &_
			" parent.jsSetStep(2); "&_
			"	alert('" & susername & "�� ������ȣ�� [" & authNo & "]�Դϴ�.');" &_
			 "self.location ='about:blank';"&_
			"</script>"
else
    if (lstp="W") then
    Response.Write "<script language=javascript>" &_
    " parent.jsSetStep(2);	  "&_
			"	alert('�޴������� ������ȣ�� �߼��߽��ϴ�.\n������ȣ �Է��� ���� �� �ּ���...');" &_
			  "self.location ='about:blank';"&_
			"</script>"
    else
	Response.Write "<script language=javascript>" &_
			"	parent.startLimitCounter('new');" &_
			" parent.jsSetStep(2);	  "&_
			"	alert('�޴������� ������ȣ�� �߼��߽��ϴ�.\nSMS�� Ȯ�� �� �α������ּ���.');" &_
			  "self.location ='about:blank';"&_
			"</script>"
	end if
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