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
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/search/search_manageCls.asp"-->
<%
dim menupos, imenuposStr, imenuposnotice, imenuposhelp
menupos = request("menupos")
if menupos ="" then menupos=1

imenuposStr = fnGetMenuPos(menupos, imenuposnotice, imenuposhelp)

'// 즐겨찾기
dim IsMenuFavoriteAdded

IsMenuFavoriteAdded = fnGetMenuFavoriteAdded(session("ssBctID"), menupos)


Dim i, cQuick, vPage, vDateGubun, vSDate, vEDate, vQuickType, vEndType, vUseYN, vSearchGubun, vSearchTxt
vPage = NullFillWith(requestCheckVar(Request("page"),10),1)
vDateGubun = NullFillWith(requestCheckVar(Request("dategubun"),10),"write")
vSDate = requestCheckVar(Request("sdate"),10)
vEDate = requestCheckVar(Request("edate"),10)
vQuickType = requestCheckVar(Request("quicktype"),3)
vEndType = requestCheckVar(Request("endtype"),10)
vUseYN = NullFillWith(requestCheckVar(Request("useyn"),1),"")
vSearchGubun = requestCheckVar(Request("searchgubun"),10)
vSearchTxt = requestCheckVar(Request("searchtxt"),50)

Set cQuick = New CSearchMng
cQuick.FCurrPage = vPage
cQuick.FPageSize = 15
cQuick.FRectDateGubun = vDateGubun
cQuick.FRectSDate = vSDate
cQuick.FRectEDate = vEDate
cQuick.FRectQuickType = vQuickType
cQuick.FRectEndType = vEndType
cQuick.FRectUseYN = vUseYN
cQuick.FRectSearchGubun = vSearchGubun
cQuick.FRectSearchTxt = vSearchTxt
cQuick.fnQuickLinkList

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
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

function searchFrm(p){
	if($("#searchgubun").val() == "" && $("#searchtxt").val() != ""){
		alert("퀵링크명, 검색 키워드, 작성자 등을 선택 후 검색하세요.");
		return;
	}
	frm1.page.value = p;
	frm1.submit();
}

function jsQuickReg(idx){
	var popquickreg;
	popquickreg = window.open('quickLinkManage.asp?idx='+idx+'','popquickreg','width=800,height=530,scrollbars=yes,resizable=yes');
	popquickreg.focus();
}

function jsQuickDelete(idx){
	if(confirm("선택하신 퀵링크를 삭제하시겠습니까?\n삭제하고 나면 복구되지 않습니다.") == true) {
		frm2.idx.value = idx;
		frm2.submit();
	} else {
		return false;
	}
}
</script>
</head>
<body>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<div class="contSectFix scrl">
	<div class="contHead">
		<div class="locate"><h2>검색 &gt; <strong>퀵링크 관리</strong></h2></div>
		<div class="helpBox">
			<form name="frmMenuFavorite" method="post" action="/admin/menu/popEditFavorite_process.asp">
				<input type="hidden" name="mode" value="">
				<input type="hidden" name="menu_id" value="3959">
			</form>
			<a href="javascript:fnMenuFavoriteAct('addonefavorite')">즐겨찾기</a> l 
			<!-- 마스터이상 메뉴권한 설정 //-->
			<a href="Javascript:PopMenuEdit('3959');">권한변경</a> l 
			<!-- Help 설정 //-->
			<a href="Javascript:PopMenuHelp('3959');">HELP</a>
		</div>
	</div>

	<!-- 상단 검색폼 시작 -->
	<form name="frm1" method="get" action="">
	<input type="hidden" name="page" value="<%=vPage%>">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<!-- search -->
	<div class="searchWrap">
		<div class="search">
			<ul>
				<li>
					<label class="formTit">기간 :</label>
					<select class="formSlt" title="옵션 선택" id="dategubun" name="dategubun">
						<option value="write" <%=CHKIIF(vDateGubun="write","selected","")%>>작성일</option>
						<option value="sdate" <%=CHKIIF(vDateGubun="sdate","selected","")%>>시작일</option>
						<option value="edate" <%=CHKIIF(vDateGubun="edate","selected","")%>>종료일</option>
					</select>
					<input type="text" class="formTxt" id="sdate" name="sdate" value="<%=vSDate%>" style="width:100px" placeholder="시작일" maxlength="10" readonly />
					<img src="/images/admin_calendar.png" id="sdate_trigger" alt="달력으로 검색" />
					<script language="javascript">
						var CAL_Start = new Calendar({
							inputField : "sdate", trigger    : "sdate_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
					~
					<input type="text" class="formTxt" id="edate" name="edate" value="<%=vEDate%>" style="width:100px" placeholder="종료일" maxlength="10" readonly />
					<img src="/images/admin_calendar.png" id="edate_trigger" alt="달력으로 검색" />
					<script language="javascript">
						var CAL_End = new Calendar({
							inputField : "edate", trigger    : "edate_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn>
		<div class="search">
			<ul>
				<li>
					<p class="formTit">퀵링크 속성 :</p>
					<select class="formSlt" id="quicktype" name="quicktype" title="옵션 선택">
						<option value="" <%=CHKIIF(vQuickType="","selected","")%>>전체</option>
						<%=fnQuickLinkTypeSelect(vQuickType)%>
					</select>
				</li>
				<li>
					<p class="formTit">종료여부 :</p>
					<select class="formSlt" id="endtype" name="endtype" title="옵션 선택">
						<option value="" <%=CHKIIF(vEndType="","selected","")%>>전체</option>
						<option value="now" <%=CHKIIF(vEndType="now","selected","")%>>진행</option>
						<option value="end" <%=CHKIIF(vEndType="end","selected","")%>>종료</option>
					</select>
				</li>
				<li>
					<p class="formTit">사용여부 :</p>
					<select class="formSlt" id="useyn" name="useyn" title="옵션 선택">
						<option value="" <%=CHKIIF(vUseYN="","selected","")%>>전체</option>
						<option value="y" <%=CHKIIF(vUseYN="y","selected","")%>>사용</option>
						<option value="n" <%=CHKIIF(vUseYN="n","selected","")%>>사용안함</option>
					</select>
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn>
		<div class="search">
			<ul>
				<li>
					<label class="formTit" for="schWord">검색어 :</label>
					<select class="formSlt" id="searchgubun" name="searchgubun" title="옵션 선택">
						<option value="" <%=CHKIIF(vSearchGubun="","selected","")%>>-선택-</option>
						<option value="q.name" <%=CHKIIF(vSearchGubun="q.name","selected","")%>>퀵링크명</option>
						<option value="k.keyword" <%=CHKIIF(vSearchGubun="k.keyword","selected","")%>>검색키워드</option>
						<option value="t.username" <%=CHKIIF(vSearchGubun="t.username","selected","")%>>작성자</option>
					</select>
					<input type="text" class="formTxt" id="searchtxt" name="searchtxt" value="<%=vSearchTxt%>" style="width:500px" placeholder="퀵링크명, 검색 키워드, 작성자 등을 선택 후 검색하세요." />
				</li>
			</ul>
		</div>
		<input type="button" class="schBtn" value="검색" onClick="searchFrm(1);" />
	</div>
	<!-- //search -->
	</form>

	<div class="cont">
		<div class="pad20">
			<div class="overHidden">
				<div class="ftLt">
					<input type="button" class="btn" value="퀵링크 등록" onClick="jsQuickReg('');" />
				</div>
			</div>

			<div>
				<div class="rt pad10">
					<span>검색결과 : <strong><%=FormatNumber(cQuick.FTotalCount,0)%></strong></span> <span class="lMar10">페이지 : <strong><%=cQuick.FtotalPage%> / <%=FormatNumber(vPage,0)%></strong></span>
				</div>
				<table class="tbType1 listTb">
					<thead>
					<tr>
						<th><div>No.</div></th>
						<th><div>퀵링크 속성</div></th>
						<th><div>퀵링크명</div></th>
						<th><div>노출기간</div></th>
						<th><div>사용여부</div></th>
						<th><div>검색 키워드</div></th>
						<th><div>작성자</div></th>
						<th><div>작성일</div></th>
						<th><div>URL</div></th>
						<th><div></div></th>
					</tr>
					</thead>
					<tbody>
					<%
						If cQuick.FResultCount > 0 Then
							For i=0 To cQuick.FResultCount-1
					%>
							<tr>
								<td><%=cQuick.FItemList(i).Fidx%></td>
								<td><%=fnQuickLinkTypeName(cQuick.FItemList(i).Fquicktype)%></td>
								<td class="lt"><%=cQuick.FItemList(i).Fquickname%></td>
								<td>
									<%
										If cQuick.FItemList(i).Fviewgubun = "always" Then
											Response.Write "상시노출"
										ElseIf cQuick.FItemList(i).Fviewgubun = "period" Then
											If cQuick.FItemList(i).Fedate < date() Then
												Response.Write "종료"
											Else
												Response.Write Left(cQuick.FItemList(i).Fsdate,10) & " ~ " & Left(cQuick.FItemList(i).Fedate,10)
											End If
										End If
									%>
								</td>
								<td><%=CHKIIF(cQuick.FItemList(i).Fuseyn="y","사용","사용안함")%></td>
								<td class="lt"><%=cQuick.FItemList(i).Fkeyword%></td>
								<td><%=cQuick.FItemList(i).Fregusername%></td>
								<td><%=Left(cQuick.FItemList(i).Fregdate, 10)%></td>
								<td>
									<a href="http://www.10x10.co.kr<%=cQuick.FItemList(i).Furl_pc%>" class="cBl1 tLine" target="_blank">[PC바로가기]</a><br />
									<a href="http://m.10x10.co.kr<%=cQuick.FItemList(i).Furl_m%>" class="cBl1 tLine" target="_blank">[M바로가기]</a>
								</td>
								<td>
									<input type="button" class="btn" value="수정" onClick="jsQuickReg('<%=cQuick.FItemList(i).Fidx%>');" />&nbsp;
									<input type="button" class="btn" value="삭제" onClick="jsQuickDelete('<%=cQuick.FItemList(i).Fidx%>');" />
								</td>
							</tr>
					<%
							Next
						End If
					%>
					</tfoot>
				</table>
				<div class="ct tPad20 cBk1">
					<% if cQuick.HasPreScroll then %>
					<a href="javascript:searchFrm('<%= cQuick.StartScrollPage-1 %>')">[pre]</a>
					<% else %>
		    			[pre]
		    		<% end if %>
		    		
		    		<% for i=0 + cQuick.StartScrollPage to cQuick.FScrollCount + cQuick.StartScrollPage - 1 %>
		    			<% if i>cQuick.FTotalpage then Exit for %>
		    			<% if CStr(vPage)=CStr(i) then %>
		    			<span class="cRd1">[<%= i %>]</span>
		    			<% else %>
		    			<a href="javascript:searchFrm('<%= i %>')">[<%= i %>]</a>
		    			<% end if %>
		    		<% next %>
					
					<% if cQuick.HasNextScroll then %>
		    			<a href="javascript:searchFrm('<%= i %>')">[next]</a>
		    		<% else %>
		    			[next]
		    		<% end if %>
				</div>
			</div>
		</div>
	</div>
</div>
<form name="frm2" action="quickLinkproc.asp" method="post" target="iframeproc" style="margin:0px;">
<input type="hidden" id="action" name="action" value="delete">
<input type="hidden" id="idx" name="idx" value="">
</form>
<iframe src="about:blank" name="iframeproc" width="0" height="0" frameborder="0"></iframe>
<% Set cQuick = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->