		<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  사원 검색 및 내용 삽입
' History : 2011.03.10 허진원 생성
'						2017.04.06 정윤정 사번으로 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenAgitCls.asp" -->
<%
dim userid, empno, sdt
	userid = requestCheckvar(request("uid"),32)
   empno = requestCheckvar(Request("sEn"),32)
   sdt =  requestCheckvar(Request("sdt"),10)
dim oMember, arrList, iTotCnt
dim clsap
dim totap, useap, reqap,payap, ispenalty, psdate, pedate, pkind

if empno="" then Response.End

'// 사번으로 검색
Set oMember = new CTenByTenMember
oMember.FPagesize 		= 10
oMember.FCurrPage 		= 1
oMember.FSearchType 	= "3"	'검색구분(회원ID)
oMember.FSearchText 	= empno
oMember.Fstatediv 		= "Y"
oMember.Fextparttime 	= "0"	'0:전사원, 1:직원이상
	
arrList = oMember.fnGetMemberList
set oMember = nothing


'// 아이디 찾기로 나온경우
IF isArray(arrList) THEN
	
set clsap = new CMyAgit
		clsap.FRectEmpno = empno
		clsap.FRectChkStart = sDt
		clsap.fnGetMyAgit
		totap = clsap.FtotPoint
		useap = clsap.FusePoint 
		pkind = clsap.Fpenaltykind
		psdate = clsap.Fpenaltysdate
		pedate = clsap.Fpenaltyedate
set clsap = nothing
%>
<script language="javascript"> 
	parent.frm.chkCfm.value="Y";	
	parent.frm.sEn.value="<%=arrList(0,0)%>";
	parent.frm.userid.value="<%=arrList(2,0)%>";
	parent.frm.username.value="<%=arrList(1,0)%>";
	parent.frm.posit_nm.value="<%=arrList(13,0)%>";
	parent.frm.avPoint.value="<%=totap-useap%>";
	parent.frm.department_nm.value="<%=arrList(27,0)%>"; 
	parent.frm.userPhone.value="<%=arrList(18,0)%>";
	parent.frm.userHP.value="<%=arrList(17,0)%>";
</script>
<%
Else
	'// 이름으로 검색
	Set oMember = new CTenByTenMember
	oMember.FPagesize 		= 10
	oMember.FCurrPage 		= 1
	oMember.FSearchType 	= "2"	'검색구분(회원명)
	oMember.FSearchText 	= empno
	oMember.Fstatediv 		= "Y"
	oMember.Fextparttime 	= "0"	'0:전사원, 1:직원이상
		
	arrList = oMember.fnGetMemberList
	iTotCnt = oMember.FTotCnt
	set oMember = nothing

	IF isArray(arrList) THEN
		IF iTotCnt>1 then
%>
<script language="javascript">
	var psu = window.open("popSelectUser.asp?unm=<%=Server.URLEncode(empno)%>","popSelUsr","width=500,height=200,scrollbars=yes");
	psu.focus();
</script>
<%
		else 
set clsap = new CMyAgit
		clsap.FRectEmpno = empno
		clsap.FRectChkStart = sDt
		clsap.fnGetMyAgit
		totap = clsap.FtotPoint
		useap = clsap.FusePoint 
		pkind = clsap.Fpenaltykind
		psdate = clsap.Fpenaltysdate
		pedate = clsap.Fpenaltyedate
set clsap = nothing

%>
<script language="javascript">
	parent.frm.chkCfm.value="Y";
	parent.frm.sEn.value="<%=arrList(0,0)%>";
	parent.frm.userid.value="<%=arrList(2,0)%>";
	parent.frm.username.value="<%=arrList(1,0)%>";
	parent.frm.posit_nm.value="<%=arrList(13,0)%>";
	parent.frm.avPoint.value="<%=totap-useap%>";
	parent.frm.department_nm.value="<%=arrList(27,0)%>"; 
	parent.frm.userPhone.value="<%=arrList(18,0)%>";
	parent.frm.userHP.value="<%=arrList(17,0)%>";
</script>
<%
		end if
	else
%>
<script language="javascript">
alert("입력하신 [<%=empno%>]은(는) 텐바이텐 직원의 사번  또는 이름이 아닙니다.\n확인 후 다시 검사해주세요.");
</script>
<%
	End if
End If
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->