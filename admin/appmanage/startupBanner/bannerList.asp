<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
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
<script type="text/javascript" src="/js/xl.js"></script>
<script type="text/javascript" src="/js/common.js"></script>
<script type="text/javascript" src="/js/report.js"></script>
<script type="text/javascript" src="/cscenter/js/cscenter.js"></script>
<script type="text/javascript" src="/js/calendar.js"></script>
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery_common.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<link rel="stylesheet" href="/css/scm.css" type="text/css" />
<script type='text/javascript'>
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
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/appmanage/startupBannerCls.asp" -->
<%
	Dim i, cSBanner, vPage, vIsUsing, vSdt, vEdt, vOs, vType, vTitle, vLink, vStat
	vPage = NullFillWith(requestCheckVar(request("page"),10),1)
	vIsUsing = NullFillWith(requestCheckVar(request("isusing"),1),"Y")

	vSdt	= requestCheckVar(request("sdt"),10)
	vEdt	= requestCheckVar(request("edt"),10)
	vOs		= requestCheckVar(request("tos"),10)
	vType	= requestCheckVar(request("ttp"),2)
	vTitle	= requestCheckVar(request("btt"),30)
	vLink	= requestCheckVar(request("lnk"),40)
	vStat	= requestCheckVar(request("stat"),1)

'	if vSdt="" then vSdt=dateadd("d",-3,date())
'	if vEdt="" then vEdt=dateadd("d",3,date())

	SET cSBanner = New CStartupBanner
	cSBanner.FCurrPage = vPage
	cSBanner.FPageSize = 20
	cSBanner.FRectStartDate	= vSdt
	cSBanner.FRectEndDate	= vEdt
	cSBanner.FRectTgOS		= vOs
	cSBanner.FRectTgType	= vType
	cSBanner.FRectTitle		= vTitle
	cSBanner.FRectLink		= vLink
	cSBanner.FRectStatus	= vStat
	cSBanner.FRectIsUsing	= vIsUsing
	cSBanner.GetStartupBannerList
%>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>
function goPage(p){
	frm1.page.value = p;
	frm1.submit();
}

function popDetail(idx){	
	var popModi;
	popModi = window.open('bannerView.asp?idx='+idx+'','popBnrView','width=1000,height=524,scrollbars=yes,resizable=yes');
	popModi.focus();
}

$(function(){
	$(".tbType1 .tbListRow").hover(function() {
		$(this).toggleClass('hover');
	});
});
</script>

<div class="contSectFix scrl">
	<div class="contHead">
		<div class="locate"><h2><%=imenuposStr%></h2></div>
		<div class="helpBox">
			<form name="frmMenuFavorite" method="post" action="/admin/menu/popEditFavorite_process.asp">
				<input type="hidden" name="mode" value="">
				<input type="hidden" name="menu_id" value="<%=menupos%>">
			</form>
			<a href="javascript:fnMenuFavoriteAct('addonefavorite')">즐겨찾기</a> l 
			<!-- 마스터이상 메뉴권한 설정 //-->
			<a href="Javascript:PopMenuEdit('<%=menupos%>');">권한변경</a> l 
			<!-- Help 설정 //-->
			<a href="Javascript:PopMenuHelp('<%=menupos%>');">HELP</a>
		</div>
	</div>

	<!-- 상단 검색폼 시작 -->
	<form name="frm1" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<div class="searchWrap">
		<div class="search rowSum1">
			<ul>
				<li>
					<label class="formTit" for="termSdt">기간 :</label>
					<input type="text" name="sdt" value="<%=vSdt%>" class="formTxt" id="termSdt" style="width:100px" placeholder="시작일" />
					<input type="image" src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkStart_trigger" onclick="return false;" />
					~
					<input type="text" name="edt" value="<%=vEdt%>" class="formTxt" id="termEdt" style="width:100px" placeholder="종료일" />
					<input type="image" src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkEnd_trigger" onclick="return false;" />
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "termSdt", trigger    : "ChkStart_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
						var CAL_End = new Calendar({
							inputField : "termEdt", trigger    : "ChkEnd_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
				</li>
				<li>
					<label class="formTit" for="srcStat">상태 :</label>
					<select name="stat" class="formSlt" id="srcStat">
						<option value="" <%=chkIIF(vStat="","selected","")%>>전체</option>
						<option value="0" <%=chkIIF(vStat="0","selected","")%>>등록대기</option>
						<option value="5" <%=chkIIF(vStat="5","selected","")%>>오픈</option>
						<option value="9" <%=chkIIF(vStat="9","selected","")%>>종료</option>
					</select>
				</li>
				<li>
					<label class="formTit" for="srcOs">타겟 :</label>
					<select name="tos" class="formSlt" id="srcOs">
						<option value="" <%=chkIIF(vOs="","selected","")%>>전체</option>
						<option value="ios" <%=chkIIF(vOs="ios","selected","")%>>iOS</option>
						<option value="android" <%=chkIIF(vOs="android","selected","")%>>Android</option>
					</select>
					<select name="ttp" class="formSlt" id="srcTp">
						<option value="" <%=chkIIF(vType="","selected","")%>>전체</option>
						<option value="00" <%=chkIIF(vType="00","selected","")%>>모든고객</option>
						<option value="30" <%=chkIIF(vType="30","selected","")%>>비회원</option>
						<option value="15" <%=chkIIF(vType="15","selected","")%>>Orange</option>
						<option value="10" <%=chkIIF(vType="10","selected","")%>>Yellow</option>
						<option value="11" <%=chkIIF(vType="11","selected","")%>>Green</option>
						<option value="12" <%=chkIIF(vType="12","selected","")%>>Blue</option>
						<option value="13" <%=chkIIF(vType="13","selected","")%>>VIP Silver</option>
						<option value="14" <%=chkIIF(vType="14","selected","")%>>VIP Gold</option>
						<option value="16" <%=chkIIF(vType="16","selected","")%>>VVIP</option>
						<option value="20" <%=chkIIF(vType="20","selected","")%>>VIP전체</option>
					</select>
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn>
		<div class="search">
			<ul>
				<li>
					<label class="formTit" for="srcTt">제목 :</label>
					<input type="text" id="srcTt" class="formTxt" name="btt" value="<%=vTitle%>" style="width:200px" />
				</li>
				<li>
					<label class="formTit" for="srcLnk">링크 :</label>
					<input type="text" id="srcLnk" class="formTxt" name="lnk" value="<%=vLink%>" style="width:200px" />
				</li>
			</ul>
		</div>
		<input type="submit" class="schBtn" value="검색" />
	</div>
	</form>
	<div class="pad20">
		<div class="overHidden">
			<div class="ftLt">
				<p class="cBk1 ftLt">* 총 <%=cSBanner.FTotalCount%> 개</p>
			</div>
			<div class="ftRt">
				<p class="btn2 cBk1 ftLt"><a href="#" onclick="popDetail('');return false;"><span class="eIcon"><em class="fIcon">신규등록</em></span></a></p>
			</div>
		</div>
		<div class="tPad15">
			<table class="tbType1 listTb">
				<thead>
				<tr>
					<th><div>No.</div></th>
					<th><div>시작일</div></th>
					<th><div>종료일</div></th>
					<th><div>이미지</div></th>
					<th><div>제목</div></th>
					<th><div>링크</div></th>
					<th><div>우선순위</div></th>
					<th><div>상태</div></th>
					<th><div>운영체제</div></th>
					<th><div>타겟</div></th>
				</tr>
				</thead>
				<tbody>
				<%
					If cSBanner.FResultCount > 0 Then
						For i=0 To cSBanner.FResultCount-1
				%>
						<tr style="cursor:pointer;" class="tbListRow">
							<td onclick="popDetail(<%=cSBanner.FItemList(i).Fidx%>);"><%=cSBanner.FItemList(i).Fidx%></td>
							<td><%=cSBanner.FItemList(i).FstartDate%></td>
							<td><%=cSBanner.FItemList(i).FexpireDate%></td>
							<td onclick="popDetail(<%=cSBanner.FItemList(i).Fidx%>);"><img src="<%=cSBanner.FItemList(i).FbannerImg%>" height="50" /></td>
							<td onclick="popDetail(<%=cSBanner.FItemList(i).Fidx%>);"><%="[" & cSBanner.FItemList(i).getLinkTypeNm & "] " & cSBanner.FItemList(i).FbannerTitle%></td>
							<td><%=cSBanner.FItemList(i).FlinkURL%>
								<a href="<%=vmobileUrl & cSBanner.FItemList(i).FlinkURL%>" target="_blank" class="cBl1 tLine lMar10">확인하기</a>
							</td>
							<td><%=cSBanner.FItemList(i).getImportanceNm%></td>
							<td><%=cSBanner.FItemList(i).getStatusNm%></td>
							<td><%=cSBanner.FItemList(i).getTargetOSNm%></td>
							<td><%=cSBanner.FItemList(i).getTargetTypeNm%></td>
						</tr>
				<%
						Next
					End If
				%>
				</tbody>
			</table>
			<br />
			<div class="ct tPad20 cBk1">
				<% if cSBanner.HasPreScroll then %>
				<a href="#" onclick="goPage(<%= cSBanner.StartScrollPage-1 %>);return false;">[pre]</a>
				<% else %>
	    			[pre]
	    		<% end if %>
	    		
	    		<% for i=0 + cSBanner.StartScrollPage to cSBanner.FScrollCount + cSBanner.StartScrollPage - 1 %>
	    			<% if i>cSBanner.FTotalpage then Exit for %>
	    			<% if CStr(vPage)=CStr(i) then %>
	    			<span class="cRd1">[<%= i %>]</span>
	    			<% else %>
	    			<a href="#" onclick="goPage(<%= i %>);return false;">[<%= i %>]</a>
	    			<% end if %>
	    		<% next %>
				
				<% if cSBanner.HasNextScroll then %>
	    			<a href="#" onclick="goPage(<%= i %>);return false;">[next]</a>
	    		<% else %>
	    			[next]
	    		<% end if %>
			</div>
		</div>
	</div>
</div>

<% SET cSBanner = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
