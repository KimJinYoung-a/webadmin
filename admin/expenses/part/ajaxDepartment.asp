<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %>
<%
'###########################################################
' Description : 운영비관리 팀 리스트  
' History : 2011.05.30 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpPartCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
Dim iPartTypeIdx, department_id, iOpExpPartIdx,sadminid,ipartsn
Dim clsPart, arrList
Dim blnAuth

	blnAuth = fnChkAdminAuth(session("ssAdminLsn"),session("ssAdminPsn"))
	IF not blnAuth THEN  '리스트 권한을 가진 사람을 제외하고 담당자와 담당부서  view 가능
		department_id = GetUserDepartmentID("",session("ssBctID"))
 		sadminid = 	session("ssBctId")
 	END IF

 	iPartTypeIdx	= requestCheckvar(Request("iPTIdx"),10)
 	iOpExpPartIdx	= requestCheckvar(Request("iPEPIdx"),10)

Set clsPart = new COpExpPart
	clsPart.FRectDepartmentID = department_id
	clsPart.FRectUserid = sadminid
	clsPart.FPartTypeidx = iPartTypeIdx
	arrList = clsPart.fnGetOpExppartAllListNew
Set clsPart = nothing
%>
<select name="selPtemp"	id="selP" class="select" onchange="chselPval(this.value);" >
	<option value="0">--선택--</option>
	<% sbOptPart arrList,iOpExpPartIdx%>
</select>
