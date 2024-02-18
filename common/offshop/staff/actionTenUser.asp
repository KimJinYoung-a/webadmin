<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  오프라인 매장근무관리
' History : 2011.03.17 한용민 생성
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/staff/staff_cls.asp"-->
<%
dim SearchText ,oMember, arrList, iTotCnt , SearchType , shopid
	SearchText = request("SearchText")
	SearchType = request("SearchType")
	shopid = request("shopid")
	
'// 아이디로 검색
Set oMember = new CAgitCalendar
	oMember.Frectshopid = shopid
	oMember.FrectSearchType 	= SearchType	'검색구분
	oMember.FrectSearchText 	= SearchText
	oMember.Frectstatediv 		= "Y"			
	oMember.fnGetMemberList

IF oMember.ftotalcount > 0 THEN
%>

	<script language="javascript">
	
		parent.frm.chkCfm.value="Y";
		parent.frm.userid.value="<%= oMember.FOneItem.Fuserid %>";
		parent.frm.empno.value="<%= oMember.FOneItem.fempno %>";
		parent.frm.username.value="<%= oMember.FOneItem.fusername %>";
		parent.frm.posit_sn.value="<%= oMember.FOneItem.fposit_sn %>";
		parent.frm.part_sn.value="<%= oMember.FOneItem.fpart_sn %>";
		
	</script>

<% Else %>

	<script language="javascript">
		alert("해당 직원이 존재 하지 않습니다\n확인 후 다시 검사해주세요.");
	</script>

<% End If %>

<%
set oMember = nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->