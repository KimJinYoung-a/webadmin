<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  휴대폰확인을 통한 휴대폰번호 변경 처리
' History : 2013.02.18 허진원 생성
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

'// 변수 할당
empno = requestCheckVar(Request.form("empNo"),18)	' 사원번호
MobileNo1 = requestCheckVar(Request.form("hpNum1"),3)	' 휴대폰번호1
MobileNo2 = requestCheckVar(Request.form("hpNum2"),4)	' 휴대폰번호2
MobileNo3 = requestCheckVar(Request.form("hpNum3"),4)	' 휴대폰번호3

'// 직원 기본정보 접수
Set cMember = new CTenByTenMember
	cMember.Fempno = empno
	cMember.fnGetMemberData

	username      	= cMember.Fusername
Set cMember = Nothing

if username="" or isNull(username) then
    Call Alert_close("직원정보가 존재하지 않습니다.")
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
	alert('본인확인 및 입력하신 휴대폰번호로 적용되었습니다.');
	parent.opener.history.go(0);
	parent.close();
	</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->