<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 클래스
' Hieditor : 2011.04.22 이상구 생성
'			 2012.08.24 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteBrandCls.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

Dim xSiteId, makerid, gubun, research, page, incnotuse
Dim i

	research = requestCheckvar(request("research"),10)
	page = requestCheckvar(request("page"),10)
	xSiteId = requestCheckvar(request("xSiteId"),32)
	makerid = requestCheckvar(request("makerid"),32)
	gubun = requestCheckvar(request("gubun"),32)
	incnotuse = requestCheckvar(request("incnotuse"),32)

''if (research="") then gubun="B001"
if (page="") then page=1

Dim oCxSiteBrand
set oCxSiteBrand = new CxSiteBrand
	oCxSiteBrand.FPageSize = 20
	oCxSiteBrand.FCurrPage = page

	oCxSiteBrand.FRectxSiteId   	= xSiteId
	oCxSiteBrand.FRectMakerid   	= makerid
	oCxSiteBrand.FRectGubun   		= gubun
	oCxSiteBrand.FRectIncNotUse   	= incnotuse

    oCxSiteBrand.getXSiteBrandList

%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

/*
function GetxSiteCSOrderList(sellsite) {
	var frm = document.frmAct;

	if (confirm("진행 하시겠습니까?") != true) {
		return;
	}

	frm.mode.value = "getxsitecslist";
	frm.sellsite.value = sellsite;
	frm.submit();
}

function jsSearchByOutMallOrderSerial(outmallorderserial) {
	var frm = document.frm;
	frm.outmallorderserial.value = outmallorderserial;
	frm.submit();
}

function jsSearchByOrderSerial(orderserial) {
	var frm = document.frm;
	frm.orderserial.value = orderserial;
	frm.submit();
}

function Cscenter_Action_List(orderserial) {
    var window_width = 1280;
    var window_height = 960;

    var popwin = window.open("/cscenter/action/cs_action.asp?orderserial=" + orderserial ,"Cscenter_Action_List","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function jsSetFinish(idx) {
	var frm = document.frmAct;

	if (confirm("완료처리 하시겠습니까?") != true) {
		return;
	}

	frm.mode.value = "setfinish";
	frm.idx.value = idx;
	frm.submit();
}

function jsDelFinish(idx) {
	var frm = document.frmAct;

	if (confirm("완료처리 취소 하시겠습니까?") != true) {
		return;
	}

	frm.mode.value = "delfinish";
	frm.idx.value = idx;
	frm.submit();
}
*/

function popInsXSiteBrandInfo(idx, xSiteId, gubun) {
    var window_width = 500;
    var window_height = 250;

    var popwin = window.open("popXSiteBrandMod.asp?xSiteId=" + xSiteId + "&gubun=" + gubun + "&idx=" + idx,"popInsXSiteBrandInfo","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		&nbsp;
	    제휴몰 :
	    <% call drawSelectBoxXSiteOrderInputPartner("xSiteId", xSiteId) %>
		&nbsp;
		브랜드 :
		<% drawSelectBoxDesignerwithName "makerid",makerid %>
	    &nbsp;
	    구분 :
		<select class="select" name="gubun"  >
			<option value="" <%= chkIIF(gubun="", "selected","") %> >전체</option>
	     	<option value="excoupon" <%= chkIIF(gubun="excoupon","selected","") %> >쿠폰제외브랜드</option>
     	</select>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		&nbsp;
     	삭제정보 포함 :<input type="checkbox" name="incnotuse"  value="Y" <% if (incnotuse = "Y") then %>checked<% end if %> >
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="등록하기" onClick="popInsXSiteBrandInfo('', '<%= xSiteId %>', '<%= gubun %>');">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		검색결과 : <b><%= oCxSiteBrand.FTotalcount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= oCxSiteBrand.FTotalPage %></b>
	</td>
</tr>
<form name="frmAct" method="post" action="xSiteCSOrder_Process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="xSiteId" value="">
<input type="hidden" name="idx" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="40">IDX</td>
	<td width="100">제휴몰</td>
	<td width="100">브랜드</td>
	<td width="100">구분</td>
	<!--
	<td width="80">시작일</td>
	<td width="80">종료일</td>
	-->
	<td width="80">상태</td>
	<td width="80">등록자</td>
	<td width="80">등록일</td>
	<td>비고</td>
</tr>

<% for i=0 to oCxSiteBrand.FresultCount -1 %>
<tr align="center" bgcolor="<% if (oCxSiteBrand.FItemList(i).Fuseyn = "Y") then %>FFFFFF<% else %>DDDDDD<% end if %>" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td><a href="javascript:popInsXSiteBrandInfo('<%= oCxSiteBrand.FItemList(i).Fidx %>', '', '')"><%= oCxSiteBrand.FItemList(i).Fidx %></a></td>
	<td><a href="javascript:popInsXSiteBrandInfo('<%= oCxSiteBrand.FItemList(i).Fidx %>', '', '')"><%= oCxSiteBrand.FItemList(i).FxSiteId %></a></td>
	<td><a href="javascript:popInsXSiteBrandInfo('<%= oCxSiteBrand.FItemList(i).Fidx %>', '', '')"><%= oCxSiteBrand.FItemList(i).Fmakerid %></a></td>
	<td><font color="<%= oCxSiteBrand.FItemList(i).GetGubunColor %>"><%= oCxSiteBrand.FItemList(i).GetGubunName %></font></td>
	<!--
	<td>
		<% if Not IsNull(oCxSiteBrand.FItemList(i).Fstartdate) then %>
			<%= Left(oCxSiteBrand.FItemList(i).Fstartdate, 10) %>
		<% end if %>
	</td>
	<td>
		<% if Not IsNull(oCxSiteBrand.FItemList(i).Fenddate) then %>
			<%= Left(oCxSiteBrand.FItemList(i).Fenddate, 10) %>
		<% end if %>
	</td>
	-->
	<td><%= oCxSiteBrand.FItemList(i).GetItemStatus %></td>
	<td><%= oCxSiteBrand.FItemList(i).Freguserid %></td>
	<td>
		<%= Left(oCxSiteBrand.FItemList(i).Fregdate, 10) %>
	</td>
	<td align="left">
		<acronym title="<%= oCxSiteBrand.FItemList(i).Fcomment %>"><%= Left(oCxSiteBrand.FItemList(i).Fcomment, 15) %></acronym>
	</td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16" align="center">
		<% if oCxSiteBrand.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCxSiteBrand.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oCxSiteBrand.StartScrollPage to oCxSiteBrand.FScrollCount + oCxSiteBrand.StartScrollPage - 1 %>
			<% if i>oCxSiteBrand.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oCxSiteBrand.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</form>
</table>

<%
set oCxSiteBrand = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
