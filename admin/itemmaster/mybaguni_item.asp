<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/classes/color/colortrend_cls.asp"-->
<%
dim menupos, imenuposStr, imenuposnotice, imenuposhelp
menupos = request("menupos")
if menupos ="" then menupos=1

imenuposStr = fnGetMenuPos(menupos, imenuposnotice, imenuposhelp)

'// 즐겨찾기
dim IsMenuFavoriteAdded

IsMenuFavoriteAdded = fnGetMenuFavoriteAdded(session("ssBctID"), menupos)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery_common.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<link rel="stylesheet" href="/css/scm.css" type="text/css" />
<script language='javascript'>
function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function fnMenuFavoriteAct(mode) {
	var frm = document.frmMenuFavorite;
	frm.mode.value = mode;

	var msg;
	var ret;
	if (mode == "delonefavorite") {
		msg = "즐겨찾기에서 제외하시겠습니까?";
	} else {
		msg = "즐겨찾기에 추가하시겠습니까?";
	}

	ret = confirm(msg);

	if (ret) {
		frm.submit();
	}
}
</script>
</head>
<body>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/mybaguni_item_cls.asp"-->
<%
	Dim i, cMyB, vPage, vSDate, vEDate
	vPage = NullFillWith(requestCheckVar(request("page"),10),1)
	vSDate = requestCheckvar(request("sdate"),10)
	vEDate = requestCheckvar(request("edate"),10)
	
	If vSDate = "" Then
		vSDate = FormatDate(DateAdd("d",-14,now()),"0000-00-00")
	End If
	
	If vEDate = "" Then
		vEDate = FormatDate(now(),"0000-00-00")
	End If

	SET cMyB = New CMyBaguni
	cMyB.FRectSDate = vSDate
	cMyB.FRectEDate = vEDate
	cMyB.FCurrPage = vPage
	cMyB.FPageSize = 20
	cMyB.fnGetMyBaguniItemList
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>

function searchFrm(p){
	frm1.page.value = p;
	frm1.submit();
}

function jsPopCal(sName){
	var winCal;

	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}
</script>

<div class="contSectFix scrl">
	<div class="contHead">
		<div class="locate"><h2>DATAMART &gt; <strong>장바구니 상품별 카운트</strong></h2></div>
		<div class="helpBox">
			<form name="frmMenuFavorite" method="post" action="/admin/menu/popEditFavorite_process.asp">
				<input type="hidden" name="mode" value="">
				<input type="hidden" name="menu_id" value="1836">
			</form>
			<a href="javascript:fnMenuFavoriteAct('addonefavorite')">즐겨찾기</a> l 
			<!-- 마스터이상 메뉴권한 설정 //-->
			<a href="Javascript:PopMenuEdit('3931');">권한변경</a> l 
			<!-- Help 설정 //-->
			<a href="Javascript:PopMenuHelp('3931');">HELP</a>
		</div>
	</div>

	<!-- 상단 검색폼 시작 -->
	<form name="frm1" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<div class="searchWrap">
		<div class="search rowSum1">
			<ul>
				<li>
					<label class="formTit" for="term1">장바구니 등록일 :</label>
					<input type="text" name="sdate" value="<%=vSDate%>" onClick="jsPopCal('sdate');" style="cursor:pointer;" size="10" maxlength="10" readonly>&nbsp;~&nbsp;
					<input type="text" name="edate" value="<%=vEDate%>" onClick="jsPopCal('edate');" style="cursor:pointer;" size="10" maxlength="10" readonly>
				</li>
			</ul>
		</div>
		<input type="submit" class="schBtn" value="검색" />
	</div>
	</form>
	
	<div class="pad20">
		<div class="overHidden">
			<div class="ftLt">
				<p class="cBk1 ftLt">* 총 <%=FormatNumber(cMyB.FTotalCount,0)%> 개</p>
			</div>
		</div>
		<div class="tPad15">
			<table class="tbType1 listTb">
				<thead>
				<tr>
					<th><div></div></th>
					<th><div>상품코드</div></th>
					<th><div>상품명</div></th>
					<th><div>현재판매가</div></th>
					<th><div>장바구니 담긴수</div></th>
					<th><div></div></th>
				</tr>
				</thead>
				<tbody>
				<%
					If cMyB.FResultCount > 0 Then
						For i=0 To cMyB.FResultCount-1
				%>
						<tr>
							<td><img src="<%=cMyB.FItemList(i).Flistimage%>" height="50"></td>
							<td><%=cMyB.FItemList(i).Fitemid%></td>
							<td><%=cMyB.FItemList(i).Fitemname%></td>
							<td><%=FormatNumber(cMyB.FItemList(i).Fsellcash,0)%></td>
							<td><%=FormatNumber(cMyB.FItemList(i).Fitemcount,0)%></td>
							<td><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=cMyB.FItemList(i).Fitemid%>" target="_blank">[10x10 상세보기]</a></td>
						</tr>
				<%
						Next
					End If
				%>
				</tbody>
			</table>
			<br />
			<div class="ct tPad20 cBk1">
				<% if cMyB.HasPreScroll then %>
				<a href="javascript:searchFrm('<%= cMyB.StartScrollPage-1 %>')">[pre]</a>
				<% else %>
	    			[pre]
	    		<% end if %>
	    		
	    		<% for i=0 + cMyB.StartScrollPage to cMyB.FScrollCount + cMyB.StartScrollPage - 1 %>
	    			<% if i>cMyB.FTotalpage then Exit for %>
	    			<% if CStr(vPage)=CStr(i) then %>
	    			<span class="cRd1">[<%= i %>]</span>
	    			<% else %>
	    			<a href="javascript:searchFrm('<%= i %>')">[<%= i %>]</a>
	    			<% end if %>
	    		<% next %>
				
				<% if cMyB.HasNextScroll then %>
	    			<a href="javascript:searchFrm('<%= i %>')">[next]</a>
	    		<% else %>
	    			[next]
	    		<% end if %>
			</div>
		</div>
	</div>
</div>

<% SET cMyB = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->