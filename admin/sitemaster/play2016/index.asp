<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : PLAYing
' Hieditor : 이종화 생성
'			 2022.07.07 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/classes/color/colortrend_cls.asp"-->
<%
dim menupos, imenuposStr, imenuposnotice, imenuposhelp
	menupos = requestCheckVar(getNumeric(request("menupos")),10)
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
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/play/play2016Cls.asp" -->
<%
	Dim i, cPl, vPage, vIsUsing, vType, vTitle, vPartMKID, vPartWDID, vPartPBID, vState, vVolnum
	vPage = NullFillWith(requestCheckVar(request("page"),10),1)
	vIsUsing = requestCheckVar(request("isusing"),1)
	vType = requestCheckVar(request("playtype"),2)
	vTitle = requestCheckVar(request("title"),200)
	vState = requestCheckVar(request("state"),2)
	vVolnum = requestCheckVar(request("volnum"),3)
	vPartMKID = NullFillWith(requestCheckVar(request("partmdid"),50),"")
	vPartWDID = NullFillWith(requestCheckVar(request("partwdid"),50),"")
	vPartPBID = NullFillWith(requestCheckVar(request("partpbid"),50),"")
	
	SET cPl = New CPlay
	cPl.FCurrPage = vPage
	cPl.FPageSize = 20
	cPl.FRectState = vState
	cPl.FRectVolnum = vVolnum
	cPl.FRectWDID = vPartWDID
	cPl.fnPlayMasterList
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>
function goNewReg(midx){
	location.href = "volwrite.asp?midx="+midx+"";
}
function searchFrm(p){
	frm1.page.value = p;
	frm1.submit();
}
function jsTagview(idx,type){	
	var poptagm;
	poptagm = window.open('pop_tagReg.asp?idx='+idx+'&playcate='+type+'','poptagm','width=500,height=400,scrollbars=yes,resizable=yes');
	poptagm.focus();
}
function jsItem(idx,type){
	var popPItem;
	popPItem = window.open('item.asp?playidx='+idx+'&playcate='+type+'','popPItem','width=1200,height=1000,scrollbars=yes,resizable=yes');
	popPItem.focus();
}
</script>

<div class="contSectFix scrl">
	<div class="contHead">
		<div class="locate"><h2>[ON] PLAY &gt; <strong>PLAYing</strong></h2></div>
		<div class="helpBox">
			<form name="frmMenuFavorite" method="post" action="/admin/menu/popEditFavorite_process.asp">
				<input type="hidden" name="mode" value="">
				<input type="hidden" name="menu_id" value="1836">
			</form>
			<a href="javascript:fnMenuFavoriteAct('addonefavorite')">즐겨찾기</a> l 
			<!-- 마스터이상 메뉴권한 설정 //-->
			<a href="Javascript:PopMenuEdit('1836');">권한변경</a> l 
			<!-- Help 설정 //-->
			<a href="Javascript:PopMenuHelp('1836');">HELP</a>
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
					<label class="formTit" for="term1">상 태 :</label>
					<select name="state" class="formSlt">
						<%=fnStateSelectBox("select",vState)%>
					</select>
				</li>
				<li>
					<label class="formTit" for="term1">Vol No. :</label>
					<input type="text" class="formTxt" name="volnum" value="<%=vVolnum%>" style="width:200px" />
					(숫자만)
				</li>
				<li>
					<label class="formTit" for="term1">WD :</label>
					<% sbGetpartid "partwdid",vPartWDID,"","12" %>
				</li>
			</ul>
		</div>
		<input type="submit" class="schBtn" value="검색" />
	</div>
	</form>
	
	<div class="pad20">
		<div class="overHidden">
			<div class="ftLt">
				<p class="cBk1 ftLt">* 총 <%=cPl.FTotalCount%> 개 / Vol No. , 오픈일 기준 정렬되어 있습니다.</p>
			</div>
			<div class="ftRt">
				<!--<p class="ftLt">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>//-->
				<p class="btn2 cBk1 ftLt"><a href="javascript:goNewReg('');"><span class="eIcon"><em class="fIcon">신규등록</em></span></a></p>
			</div>
		</div>
		<div class="tPad15">
			<table class="tbType1 listTb">
				<thead>
				<tr>
					<th><div>idx</div></th>
					<th><div>Vol No.</div></th>
					<th><div>타이틀</div></th>
					<th><div>상태</div></th>
					<th><div>오픈일</div></th>
					<th><div>BGcolor</div></th>
					<th><div>작업일자</div></th>
					<th><div>담당자</div></th>
					<th><div>작업자</div></th>
					<th><div></div></th>
				</tr>
				</thead>
				<tbody>
				<%
					If cPl.FResultCount > 0 Then
						For i=0 To cPl.FResultCount-1
				%>
						<tr>
							<td><%=cPl.FItemList(i).Fmidx%></td>
							<td><%=Format00(3,cPl.FItemList(i).Fvolnum)%></td>
							<td><%= ReplaceBracket(cPl.FItemList(i).Ftitle) %></td>
							<td><%=fnStateSelectBox("one",cPl.FItemList(i).Fstate)%></td>
								<td><%=cPl.FItemList(i).Fstartdate%></td>
							<td bgcolor="#<%=cPl.FItemList(i).Fmobgcolor%>">#<%=cPl.FItemList(i).Fmobgcolor%></td>
							<td>
								등록일 : <%=cPl.FItemList(i).Fregdate%><br /><br />
								수정 : <%=cPl.FItemList(i).Flastupdatename%><br /><%=cPl.FItemList(i).Flastupdate%></td>
							<td><%=cPl.FItemList(i).FpartMKname%></td>
							<td>
								WD:<%=cPl.FItemList(i).FpartWDname%><br />
								PB:<%=cPl.FItemList(i).FpartPBname%>
							</td>
							<td>
								<input type="button" onClick="goNewReg('<%=cPl.FItemList(i).Fmidx%>');" value="수 정">
							</td>
						</tr>
				<%
						Next
					End If
				%>
				</tbody>
			</table>
			<br />
			<div class="ct tPad20 cBk1">
				<% if cPl.HasPreScroll then %>
				<a href="javascript:searchFrm('<%= cPl.StartScrollPage-1 %>')">[pre]</a>
				<% else %>
	    			[pre]
	    		<% end if %>
	    		
	    		<% for i=0 + cPl.StartScrollPage to cPl.FScrollCount + cPl.StartScrollPage - 1 %>
	    			<% if i>cPl.FTotalpage then Exit for %>
	    			<% if CStr(vPage)=CStr(i) then %>
	    			<span class="cRd1">[<%= i %>]</span>
	    			<% else %>
	    			<a href="javascript:searchFrm('<%= i %>')">[<%= i %>]</a>
	    			<% end if %>
	    		<% next %>
				
				<% if cPl.HasNextScroll then %>
	    			<a href="javascript:searchFrm('<%= i %>')">[next]</a>
	    		<% else %>
	    			[next]
	    		<% end if %>
			</div>
		</div>
	</div>
</div>

<% SET cPl = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->