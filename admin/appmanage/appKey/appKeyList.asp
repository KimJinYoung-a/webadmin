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
<!-- #include virtual="/lib/classes/appmanage/appKeyCls.asp" -->
<%
	Dim i, vPage, vIsUsing, vType, vOsType, vAppVersion, cAppKeys
	vPage = NullFillWith(requestCheckVar(request("page"),10),1)
	vIsUsing = requestCheckVar(request("isusing"),1)

	vType	= requestCheckVar(request("type"),50)
	vOsType	= requestCheckVar(request("ostype"),50)
	vAppVersion	= requestCheckVar(request("appversion"),50)

	SET cAppKeys = New CappKey
	cAppKeys.FCurrPage		= vPage
	cAppKeys.FPageSize 		= 10
	cAppKeys.FRectType		= vType
	cAppKeys.FRectOsType		= vOsType
	cAppKeys.FRectAppVersion	= vAppVersion
	cAppKeys.FRectIsUsing	= vIsUsing
	cAppKeys.GetAppKeyList
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
	popModi = window.open('appKeyView.asp?idx='+idx+'','popAppKeyView','width=1000,height=524,scrollbars=yes,resizable=yes');
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
					<label class="formTit" for="type">앱구분 :</label>
					<select name="type" class="formSlt" id="type">
						<option value="" <%=chkIIF(vType="","selected","")%>>전체</option>
						<option value="wishapp" <%=chkIIF(vType="wishapp","selected","")%>>wishapp</option>
						<option value="hitchhiker" <%=chkIIF(vType="hitchhiker","selected","")%>>hitchhiker</option>
					</select>
				</li>			
				<li>
					<label class="formTit" for="ostype">OS종류 :</label>
					<select name="ostype" class="formSlt" id="ostype">
						<option value="" <%=chkIIF(vOsType="","selected","")%>>전체</option>
						<option value="android" <%=chkIIF(vOsType="android","selected","")%>>Android</option>
						<option value="ios" <%=chkIIF(vOsType="ios","selected","")%>>iOS</option>
					</select>					
				</li>
				<li>
					<label class="formTit" for="isusing">사용여부 :</label>
					<select name="isusing" class="formSlt" id="isusing">
						<option value="" <%=chkIIF(vIsUsing="","selected","")%>>전체</option>
						<option value="Y" <%=chkIIF(vIsUsing="Y","selected","")%>>사용</option>
						<option value="N" <%=chkIIF(vIsUsing="N","selected","")%>>사용안함</option>
					</select>
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn>
		<input type="submit" class="schBtn" value="검색" />
	</div>
	</form>
	<div class="pad20">
		<div class="overHidden">
			<div class="ftLt">
				<p class="cBk1 ftLt">* 총 <%=cAppKeys.FTotalCount%> 개</p>
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
					<th><div>앱종류</div></th>
					<th><div>OS종류</div></th>
					<th><div>APP버전</div></th>
					<th><div>인증키</div></th>
					<th><div>등록일</div></th>
					<th><div>수정일</div></th>
					<th><div>등록자</div></th>
					<th><div>사용여부</div></th>
				</tr>
				</thead>
				<tbody>
				<%
					If cAppKeys.FResultCount > 0 Then
						For i=0 To cAppKeys.FResultCount-1
				%>
						<tr style="cursor:pointer;" class="tbListRow">
							<td onclick="popDetail(<%=cAppKeys.FKeyList(i).Fidx%>);"><%=cAppKeys.FKeyList(i).Fidx%></td>
							<td onclick="popDetail(<%=cAppKeys.FKeyList(i).Fidx%>);"><%=cAppKeys.FKeyList(i).Ftype%></td>
							<td onclick="popDetail(<%=cAppKeys.FKeyList(i).Fidx%>);"><%=cAppKeys.FKeyList(i).FosType%></td>
							<td onclick="popDetail(<%=cAppKeys.FKeyList(i).Fidx%>);"><%=cAppKeys.FKeyList(i).FappVersion%></td>
							<td onclick="popDetail(<%=cAppKeys.FKeyList(i).Fidx%>);"><%=cAppKeys.FKeyList(i).FvalidationKey%></td>
							<td onclick="popDetail(<%=cAppKeys.FKeyList(i).Fidx%>);"><%=cAppKeys.FKeyList(i).FregDate%></td>
							<td onclick="popDetail(<%=cAppKeys.FKeyList(i).Fidx%>);"><%=cAppKeys.FKeyList(i).FlastUpDate%></td>
							<td onclick="popDetail(<%=cAppKeys.FKeyList(i).Fidx%>);"><%=cAppKeys.FKeyList(i).FadminName%></td>
							<td onclick="popDetail(<%=cAppKeys.FKeyList(i).Fidx%>);"><%=cAppKeys.FKeyList(i).getIsUsingNm%></td>
						</tr>
				<%
						Next
					End If
				%>
				</tbody>
			</table>
			<br />
			<div class="ct tPad20 cBk1">
				<% if cAppKeys.HasPreScroll then %>
				<a href="#" onclick="goPage(<%= cAppKeys.StartScrollPage-1 %>);return false;">[pre]</a>
				<% else %>
	    			[pre]
	    		<% end if %>
	    		
	    		<% for i=0 + cAppKeys.StartScrollPage to cAppKeys.FScrollCount + cAppKeys.StartScrollPage - 1 %>
	    			<% if i>cAppKeys.FTotalpage then Exit for %>
	    			<% if CStr(vPage)=CStr(i) then %>
	    			<span class="cRd1">[<%= i %>]</span>
	    			<% else %>
	    			<a href="#" onclick="goPage(<%= i %>);return false;">[<%= i %>]</a>
	    			<% end if %>
	    		<% next %>
				
				<% if cAppKeys.HasNextScroll then %>
	    			<a href="#" onclick="goPage(<%= i %>);return false;">[next]</a>
	    		<% else %>
	    			[next]
	    		<% end if %>
			</div>
		</div>
	</div>
</div>

<% SET cAppKeys = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
