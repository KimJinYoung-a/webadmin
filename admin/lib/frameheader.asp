<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/classes/partners/contractcls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<%
dim btcid, shopusercount
	btcid= session("ssBctID")

if (btcid="") then response.End

'//오프라인 관리매장수..파트번호 수기로 박아넣지 않고.. 쿼리한다	'/2013.03.04 한용민 추가
shopusercount = getshopusercount(btcid)

'### 전자결제-받은결제함-결제문서-결제대기 1이상인경우 ###
Dim clsLeapp, vIsReceiveEApp
vIsReceiveEApp = False
set clsLeapp = new CEApproval
clsLeapp.FadminId = btcid
clsLeapp.fnGetLeftMenu
If clsLeapp.FReportstate100 + clsLeapp.FReportstate101 > 0 Then
	vIsReceiveEApp = True
End If
set clsLeapp = nothing
'####################################################

function IsIppbxmngAvaile()
    IsIppbxmngAvaile = false
    IsIppbxmngAvaile = (session("ssAdminPsn")=10) or C_ADMIN_AUTH
end function

dim isISViewTopValid : isISViewTopValid = (LCASE(session("ssBctId"))<>"iiitester")
%>
<html>
<head>
<title>[10x10] Business Comunication</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<% if (isISViewTopValid) then %>
<script language="JavaScript" src="/js/common.js"></script>
<script language='javascript'>

function WindowMinSize(){
	parent.document.all('menuset').cols = "20,*";
	document.all.WINSIZE[0].style.display = "none";
	document.all.WINSIZE[1].style.display = "";
}

function WindowMaxSize(){
	parent.document.all('menuset').cols = "180,*";
	document.all.WINSIZE[0].style.display = "";
	document.all.WINSIZE[1].style.display = "none";
}

function PopBrandList(){
    var popwin = window.open("<%= getSCMSSLURL %>/admin/member/designerinfolist.asp","designerinfolist","width=1400 height=800 scrollbars=yes resizable=yes");
    popwin.focus();
}

function PopItemStock(){
    var popwin = window.open("<%=manageUrl%>/admin/stock/itemcurrentstock.asp?menupos=709","popitemstocklist","width=1400 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function popVcCalendar(part_sn){
    var popwin = window.open("<%=manageUrl%>/admin/member/tenbyten/pop_vacation_calendar.asp?part_sn=" + part_sn,"popVcCalendar","width=900,height=700,scrollbars=yes,resizable=yes");
	popwin.focus();
}

// refere 때문에 동적으로 생성
// 2016-02-23, skyer9, http-https 문제로 오류 발생 및 처리
function Shiftshop(comp){
	/*
	var targetFrm = top.contents;

	var o  = targetFrm.document.createElement("form");
    var oi1 = targetFrm.document.createElement("input");

	oi1.type = "hidden";
	oi1.name = "shiftid";
	oi1.value = comp.value;

	o.appendChild(oi1);
	targetFrm.document.body.appendChild(o);

	o.method = "get";
	o.action = "/admin/lib/shiftshop.asp";
	o.submit();
	document.location.reload();
	*/

	var frm = document.createElement("form");
	var obj = document.createElement("input");

	frm.method = "post";
	frm.action = "/admin/lib/shiftshop.asp";
	obj.type = "hidden";
	obj.name = "shiftid";
	obj.value = comp.value;

	frm.appendChild(obj);
	document.body.appendChild(frm);

	frm.submit();
}

//전자결재
function popEapp(){
	 var winEapp = window.open("<%=manageUrl%>/admin/approval/eapp/popIndex.asp","popEapp","width="+(screen.availWidth-100)+", height="+ (screen.availHeight-100) +",resizable=yes, scrollbars=yes");
	 winEapp.focus();
}

//업무협조
function popCooperate(){
	 var winCooperate = window.open("<%=manageUrl%>/admin/cooperate/popIndex.asp","popCooperate","width="+(screen.availWidth-100)+", height="+ (screen.availHeight-100) +",resizable=yes, scrollbars=yes");
	 winCooperate.focus();
}

//업무협조
function popPartCooperate(){
	var winCooperate = window.open("<%=manageUrl%>/admin/breakdown/?menupos=1378","popPartCooperate","width="+(screen.availWidth-100)+", height="+ (screen.availHeight-100) +",resizable=yes, scrollbars=yes");
	winCooperate.focus();
}

//이미지 Refresh (Purge)
function popImgPurge() {
	 var winImgPurge = window.open("/admin/apps/popImagePurge.asp","popImgPurge","width=500, height=400,resizable=yes, scrollbars=yes");
	 winImgPurge.focus();
}

//상품판매추이그래프
function popItemSellGraph() {
	 var popItemSellGraph = window.open("/admin/maechul/itemTrend.asp","popItemTrend","width=1400, height=800,resizable=yes, scrollbars=yes");
	 popItemSellGraph.focus();
}
</script>
<% end if %>
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="0" leftmargin="0">

<% if (application("Svr_Info")="Dev") then %>
<center><b><font color="red">This is 2011 Test Server...</font></b></center>
<% end if %>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr height="40" valign="top">
    <td valign="bottom">
    	&nbsp;
        <img src="/images/admin_logo_10x10.jpg" width="90" height="25" align="absbottom" border="0" title="메인으로 이동" style="cursor:pointer;" onclick="top.location.reload()" />
    	<b>Business Communication Tool</b>

    </td>
    <td align="right" valign="bottom">
    <% if (isISViewTopValid) then %>
        <% if (session("ssAdminPsn")="17") then %> <!-- 관계사 제낌-->
            <a href="/common/pop_organization_chart.asp" target="_blank" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">조직도</a>
	    	|
	    	<a href="javascript:PopBrandList();" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">브랜드리스트</a>
		    |
			<a href="javascript:popItemSellGraph();" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">상품판매추이</a>
			|
		    <a href="javascript:PopItemStock();" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">상품별재고현황</a>
		    <!-- moon 요청 -->
        <% else %>
			<a href="/admin/seminar/seminar_calendar.asp?menupos=1482" target="contents" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'" >세미나실 예약</a>
	        |
            <a href="javascript:popCooperate();"   onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'" >업무협조</a>
	        |
			<a href="javascript:popEapp();"   onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'" >전자결재</a>
			<% If vIsReceiveEApp Then %><img src="<%=manageUrl%>/images/ico_new.png"><% End If %>
	        |
	    	<a href="/common/pop_organization_chart.asp" target="_blank" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">조직도</a>
	    	|
	    	<a href="<%=manageUrl%>/admin/member/tenbyten/Agit/tenbyten_agit_calendar.asp" target="contents" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">텐바이텐아지트</a>
	    	|
        	<a href="javascript:popVcCalendar('');" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'" >휴가달력보기</a>
        	|
        	<a href="#" onclick="printbarcode_on_off_multi(); return false;" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'" >바코드출력</a>
        	|

	        <%
	        '/본사직원일경우
	        if (session("ssBctDiv")<=9) then
	        %>
			   <a href="" onclick="popImgPurge(); return false;" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'" >이미지Refresh</a>
	        	|
	        	<a href="javascript:PopBrandList();" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">브랜드리스트</a>
		        |
				<a href="javascript:popItemSellGraph();" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">상품판매추이</a>
				|
		        <a href="javascript:PopItemStock();" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">상품별재고현황</a>
		        |
		        <%
		        '//오프라인팀일경우		'/2012-03-02 용만 추가
		        if session("ssAdminPsn") = "13" then
		        %>
	        		<a href="<%= manageUrl %>/admin/offshop/board/offshop_board.asp" target="_blank" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">매장통합게시판</a>
		        	|
		    	<% end if %>

		    <% end if %>

			<%
			'//오프라인 관리매장이 있을경우		'/2011-01-07 용만 추가
			if shopusercount > 0 then
			%>
	        	<a href="javascript:PopBrandList();" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">브랜드리스트</a>
		        |
	        	<a href="<%= manageUrl %>/admin/offshop/board/offshop_board.asp" target="_blank" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">매장통합게시판</a>
		        |
				<% drawSelectshopuser "shopshift" ,session("ssBctBigo"),btcid,"onChange='Shiftshop(this);'" %>
			<% end if %>

			<% if (application("Svr_Info")="Dev") then %>
			<b>DEV</b> : <a href="/login/dologout.asp"><img src="/images/icon_logout.gif" width="64" height="17" border="0" align="absbottom"></a>
			<% end if %>
		<% end if %>
	<% end if %>
    </td>
</tr>
</table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#CCCCCC">
<tr height="20"  valign="top">
    <td width="175" align="right" valign="middle">
    <!--
		<div id=WINSIZE style="display:">창 확대하기
			<input type="button" class="button" value="☜" onClick="javascript:WindowMinSize()">
		</div>
		<div id=WINSIZE style="display:none">창 축소하기
			<input type="button" class="button" value="☜" onClick="javascript:WindowMaxSize()">
		</div>
	-->
    </td>
    <td >
        <% if (IsIppbxmngAvaile) and C_ADMIN_AUTH then %>
        <iframe id="i_ippbxmng" name="i_ippbxmng" src="/cscenter/ippbxmng/ippbxlogin_eicn2.asp" width="500" height="20" scrolling="no" frameborder="0"></iframe>
		<% elseif (IsIppbxmngAvaile) then %>
		<iframe id="i_ippbxmng" name="i_ippbxmng" src="/cscenter/ippbxmng/ippbxlogin_eicn2.asp" width="500" height="20" scrolling="no" frameborder="0"></iframe>
        <% else %>
        &nbsp;
        <% end if %>
    </td>
    <td width="500" align="right" valign="middle">
        <b><%=session("ssBctId")%>(<%= session("ssBctCname") %>)</b> 님이 로그인 하셨습니다.
    	&nbsp;
    	<a href="/login/dologout.asp"><img src="/images/icon_logout.gif" width="64" height="17" border="0" align="absbottom"></a>
    	&nbsp;&nbsp;
    </td>
</tr>
</table>


</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
