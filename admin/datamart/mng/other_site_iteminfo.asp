<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include file="./other_site_iteminfo_cls.asp" -->
<%

'// =============================================================================
'//
'// 데이타 생성 프로세스
'//
'// 1. AWS Lambda Service 에서 데이타 생성 (myGetRemoteItemInfo 참조)
'//
'//   - 생성된 데이타는 AWS S3 에 저장됨 (Amazon S3/tenbyten-weblog-seoul/remote_site_item_price 참조)
'//
'// 2. S3 에서 디비로 데이타입력
'//
'//   - 192.168.0.103 디비에서 파일 가져옴
'//
'//   - 가져온 csv 파일을 디비로 벌크 인서트함 (73번 디비 [db_analyze_etc].[dbo].[sp_Ten_Site_Price_BulkInsert] 참조)
'//
'// =============================================================================

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
<%
	Dim i, cOSIt, vPage, vRegDate, vTen, vOther, vIsMatch
	vPage = NullFillWith(requestCheckVar(request("page"),10),1)
	vRegDate = NullFillWith(requestCheckVar(request("regdate"),10),date())
	vIsMatch = NullFillWith(requestCheckVar(request("ismatching"),1),"a")
	vTen = "<strong class=""fontred"">[T]</strong>"
	vOther = "<strong>[타]</strong>"

	SET cOSIt = New COSItem
	cOSIt.FCurrPage = vPage
	cOSIt.FPageSize = 20
	cOSIt.FRectRegDate = vRegDate
	cOSIt.FRectIsMatch = CHKIIF(vIsMatch="a","",vIsMatch)
	cOSIt.fnOtherSiteItemlist
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<style type="text/css">
.fontred {color:#FF0000 !important;}
</style>
<script>
function searchFrm(p){
	frm1.page.value = p;
	frm1.submit();
}
function jsSiteURL(u){
	var siteitem;
	siteitem = window.open('/admin/datamart/mng/no_referer_go_url.asp?urll='+u+'','siteitem','width=1600,height=1000,toolbar=yes, location=yes, directories=yes, status=yes, menubar=yes, scrollbars=yes, copyhistory=yes, resizable=yes');
	//siteitem = window.open('http://testwebadmin.10x10.co.kr/admin/sitemaster/play2016/','siteitem','width=1500,height=1000,toolbar=yes, location=yes, directories=yes, status=yes, menubar=yes, scrollbars=yes, copyhistory=yes, resizable=yes');
	siteitem.focus();
}
function jsMatchEdit(c,i,d){
	var matchitem;
	matchitem = window.open('/admin/datamart/mng/other_site_item_search.asp?sitecode='+c+'&siteitemid='+i+'&regdate='+d+'','matchitem','width=1000,height=900, scrollbars=yes, resizable=yes');
	matchitem.focus();
}
</script>

<div class="contSectFix scrl">
	<div class="contHead">
		<div class="locate"><h2><%=imenuposStr%></h2></div>
		<div class="helpBox">
			<form name="frmMenuFavorite" method="post" action="/admin/menu/popEditFavorite_process.asp">
				<input type="hidden" name="mode" value="">
				<input type="hidden" name="menu_id" value="3942">
			</form>
			<% if (menupos > 1) then %>
				<% if (IsMenuFavoriteAdded) then %>
					<a href="javascript:fnMenuFavoriteAct('delonefavorite')">즐겨찾기</a> l
				<% else %>
					<a href="javascript:fnMenuFavoriteAct('addonefavorite')">즐겨찾기</a> l
				<% end if %>
			<% end if %>
			<!-- 마스터이상 메뉴권한 설정 //-->
			<% if C_ADMIN_AUTH then %>
			<a href="Javascript:PopMenuEdit('3942');">권한변경</a> l
			<% end if %>
			<!-- Help 설정 //-->
			<% if (imenuposhelp<>"") or (C_ADMIN_AUTH) then %>
			<a href="Javascript:PopMenuHelp('3942');">HELP</a>
			<% end if %>
		</div>
	</div>

	<!-- 상단 검색폼 시작 -->
	<form name="frm1" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<div class="searchWrap">
		<div class="search rowSum1">
			<ul>
				<li class="lMar10 rMar10">
					등록일 :
					<input type="text" name="regdate" id="regdate" value="<%=vRegDate%>" size="10" maxlength="10" readonly>
					<span><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="regdate_trigger" border="0" style="cursor:pointer" /></span>
				</li>
				<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "regdate", trigger    : "regdate_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
				</script>
				<li class="lMar10 rMar10">
					<span class="tPad10"><label class="formTit" for="term1">10x10 매칭 :</label></span>
					<input type="radio" name="ismatching" id="ismatching" value="a" style="width: 1.6em; height: 1.6em;" <%=CHKIIF(vIsMatch="a","checked","")%>> 전체&nbsp;&nbsp;
					<input type="radio" name="ismatching" id="ismatching" value="o" style="width: 1.6em; height: 1.6em;" <%=CHKIIF(vIsMatch="o","checked","")%>> 매칭된것&nbsp;&nbsp;
					<input type="radio" name="ismatching" id="ismatching" value="x" style="width: 1.6em; height: 1.6em;" <%=CHKIIF(vIsMatch="x","checked","")%>> 매칭안된것
				</li>
			</ul>
		</div>
		<input type="submit" class="schBtn" value="검색" />
	</div>
	</form>

	<div class="pad20">
		<div class="overHidden">
			<div class="ftLt">
				<p class="cBk1 ftLt">* 총 <%=cOSIt.FTotalCount%> 개</p>
			</div>
			<div class="ftRt">
				※ <strong class="fontred">[T]</strong> : 10x10, <strong>[타]</strong> : 타사이트
			</div>
		</div>
		<div class="tPad15">
			<table class="tbType1 listTb">
				<thead>
				<tr>
					<th><div>등록일</div></th>
					<th><div><%=vOther%>구분</div></th>
					<th><div><%=vOther%>상품코드</div></th>
					<th><div><%=vOther%>상품명</div></th>
					<th><div><%=vOther%>브랜드</div></th>
					<th><div><%=vOther%>가격</div></th>
					<th><div><%=vTen%>상품코드</div></th>
					<th><div><%=vTen%>상품명</div></th>
					<th><div><%=vTen%>브랜드</div></th>
					<th><div><%=vTen%>가격</div></th>
					<th><div></div></th>
				</tr>
				</thead>
				<tbody>
				<%
					If cOSIt.FResultCount > 0 Then
						For i=0 To cOSIt.FResultCount-1
				%>
						<tr>
							<td><%=cOSIt.FItemList(i).FregTime%></td>
							<td><%=cOSIt.FItemList(i).Fsitecode%></td>
							<td><%=cOSIt.FItemList(i).Fsiteitemcode%> [<a href="" onClick="jsSiteURL('<%=fnSiteURL(cOSIt.FItemList(i).Fsitecode,cOSIt.FItemList(i).Fsiteitemcode)%>'); return false;">링크</a>]</td>
							<td style="text-align:left;"><%=cOSIt.FItemList(i).Fsiteitemname%></td>
							<td><%=cOSIt.FItemList(i).Fobrandname%></td>
							<td>
								<%
									If cOSIt.FItemList(i).ForgsellCost <> cOSIt.FItemList(i).FrealsellCost Then
										Response.Write "<strike>" & FormatNumber(cOSIt.FItemList(i).ForgsellCost,0) & "</strike> -> " & FormatNumber(cOSIt.FItemList(i).FrealsellCost,0)
										Response.Write "[<span class='fontred'>" & fnPercentView(cOSIt.FItemList(i).ForgsellCost,cOSIt.FItemList(i).FrealsellCost) & "%</span>]"
									Else
										Response.Write FormatNumber(cOSIt.FItemList(i).FrealsellCost,0)
									End If
								%>
							</td>
							<td>
								<%
									If cOSIt.FItemList(i).Fitemid <> 0 Then
										Response.Write cOSIt.FItemList(i).Fitemid & " [<a href='http://www.10x10.co.kr/" & cOSIt.FItemList(i).Fitemid & "' target='_blank'>링크</a>]"
									End If
								%>
							</td>
							<td style="text-align:left;">
								<%
									If cOSIt.FItemList(i).Fitemid <> 0 Then
										Response.Write  cOSIt.FItemList(i).Fitemname
									End If
								%>
							</td>
							<td><%=cOSIt.FItemList(i).Fbrandname%></td>
							<td>
								<%
									If cOSIt.FItemList(i).Fitemid <> 0 Then
										Response.Write FormatNumber(cOSIt.FItemList(i).Fsellcash,0)
									End If
								%>
							</td>
							<td><input type="button" value="매칭" onClick="jsMatchEdit('<%=cOSIt.FItemList(i).Fsitecode%>','<%=cOSIt.FItemList(i).Fsiteitemcode%>','<%=cOSIt.FItemList(i).FregTime%>');"></td>
						</tr>
				<%
						Next
					End If
				%>
				</tbody>
			</table>
			<br />
			<div class="ct tPad20 cBk1">
				<% if cOSIt.HasPreScroll then %>
				<a href="javascript:searchFrm('<%= cOSIt.StartScrollPage-1 %>')">[pre]</a>
				<% else %>
	    			[pre]
	    		<% end if %>

	    		<% for i=0 + cOSIt.StartScrollPage to cOSIt.FScrollCount + cOSIt.StartScrollPage - 1 %>
	    			<% if i>cOSIt.FTotalpage then Exit for %>
	    			<% if CStr(vPage)=CStr(i) then %>
	    			<span class="cRd1">[<%= i %>]</span>
	    			<% else %>
	    			<a href="javascript:searchFrm('<%= i %>')">[<%= i %>]</a>
	    			<% end if %>
	    		<% next %>

				<% if cOSIt.HasNextScroll then %>
	    			<a href="javascript:searchFrm('<%= i %>')">[next]</a>
	    		<% else %>
	    			[next]
	    		<% end if %>
			</div>
		</div>
	</div>
</div>
<% SET cOSIt = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
