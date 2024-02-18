<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 메일진
' History : 2018.04.27 이상구 생성(메일러 연동 생성 메일러로 발송 내역 전송. 메일 가져오기 생성.)
'			2019.06.24 정태훈 수정(템플릿 기능 신규 추가)
'			2020.05.28 한용민 수정(TMS 메일러 추가)
'###########################################################
%>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbTMSopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/mailzinenewcls.asp"-->

<%
Dim omail,ix,page,sDt, eDt, area, isusing, SearchKey, mailergubun
	page = requestcheckvar(getNumeric(request("page")),10)
	sDt = requestcheckvar(request("sDt"),10)
	eDt = requestcheckvar(request("eDt"),10)
	area = requestcheckvar(request("area"),32)
	isusing = requestcheckvar(request("isusing"),1)
	SearchKey = requestcheckvar(request("SearchKey"),256)
	mailergubun = requestcheckvar(request("mailergubun"),16)

if page = "" then page = 1
if sDt = "" then sDt = dateadd("d",-30,date)
if eDt = "" then eDt = dateadd("d",7,date)
'if mailergubun = "" then mailergubun = "EMS"

if mailergubun="" or isnull(mailergubun) then
	response.write "메일러 구분이 없습니다."
	dbget.close() : response.end
end if

set omail = new CMailzineList
	omail.FPageSize = 20
	omail.FCurrPage = page
	omail.FrectSDate = sDt
	omail.FrectEDate = eDt
	omail.FrectSearchKey = SearchKey
	omail.FrectUsing = isusing
	omail.FrectArea = area
	omail.frectmailergubun = mailergubun

	if mailergubun<>"" then
		omail.MailzineList()
	end if
%>

<script type="text/javascript">

// 등록(수기메일)
function editreg(idx){
	var editreg = window.open('/admin/mailzine/mailzine_detail.asp?idx='+idx+'&mailergubun=<%= mailergubun %>&menupos=<%= menupos %>','editreg','width=1400,height=800,scrollbars=yes,resizable=yes');
	editreg.focus();
}

// 등록(자동)
function jsModifyMailzine(idx) {
	var popwin = window.open('/admin/mailzine/mailzine_detail_new.asp?idx='+idx+'&mailergubun=<%= mailergubun %>&menupos=<%= menupos %>','jsModifyMailzine','width=1400,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// 등록(자동,템플릿)
function jsModifyNewMailzine(idx) {
	var popwin = window.open('/admin/mailzine/template/mailzine_detail_setting.asp?idx='+idx+'&mailergubun=<%= mailergubun %>&menupos=<%= menupos %>','jsModifyMailzine','width=1400,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function siteyn(idx, isusing, gubun){

	var chisusing;

	if(isusing == 'Y'){
		chisusing = 'N';
	}else{
		chisusing = 'Y';
		if (gubun != '5') {
			alert('\n\n디자인이 완성상태가 아닙니다.\n\n');
			return false;
		}
	}

	if (confirm('현재상태는'+isusing+'입니다.\n'+chisusing+'로 변경하시겠습니까?') == true) {
		FrameCKP.location.href='/admin/mailzine/mailzine_siteyn.asp?idx='+idx+'&isusing='+isusing+'&menupos=<%= menupos %>';
	}else{
		return false;
	}
}

function blackListReg(){
	var popBlackList = window.open('/admin/mailzine/mailzine_blacklist_pop.asp?menupos=<%= menupos %>','popBlackListReg','width=600,height=200,scrollbars=yes,resizable=yes');
	popBlackList.focus();
}

function displayManual(idx,member){
	var popDisplayManual = "";

	if(member=='member'){
		popDisplayManual = window.open('/admin/mailzine/mailzine_display.asp?idx='+idx+'&menupos=<%= menupos %>','displayManual','width=1400,height=800,scrollbars=yes,resizable=yes');
	}else{
		popDisplayManual = window.open('/admin/mailzine/mailzine_display_not.asp?idx='+idx+'&menupos=<%= menupos %>','displayManual','width=1400,height=800,scrollbars=yes,resizable=yes');
	}

	popDisplayManual.focus();
}

function displayNew(idx, member, type) {
	var display = window.open('/admin/mailzine/mailzine_display_new.asp?idx='+idx + '&member=' + member + '&type=' + type+'&menupos=<%= menupos %>','displayNew','width=1400,height=800,scrollbars=yes,resizable=yes');
	display.focus();
}

// 신형템플릿 메일 발송
function displayTemplates(idx, member, type) {
	var popDisplayTemplates = window.open('/admin/mailzine/template/mailzine_display.asp?idx='+idx + '&member=' + member + '&type=' + type+'&mailergubun=<%= mailergubun %>&menupos=<%= menupos %>','displayTemplates','width=1400,height=800,scrollbars=yes,resizable=yes');
	popDisplayTemplates.focus();
}

function mailCodeView(idx,member){
	var popmailCodeView = "";

	if(member=='member'){
		popmailCodeView = window.open('/admin/mailzine/mailzine_code_view.asp?idx='+idx+'&menupos=<%= menupos %>','code','width=1400,height=800,scrollbars=yes,resizable=yes');
	}else if(member=='basicMailFormCopy'){
		popmailCodeView = window.open('/admin/mailzine/mailzine_target_templet.asp?menupos=<%= menupos %>','code','width=1400,height=800,scrollbars=yes,resizable=yes');
	}else{
		popmailCodeView = window.open('/admin/mailzine/mailzine_code_view_not.asp?idx='+idx+'&menupos=<%= menupos %>','code','width=1400,height=800,scrollbars=yes,resizable=yes');
	}

	popmailCodeView.focus();
}

function goPage(pg) {
	document.frm.page.value=pg;
	document.frm.submit();
}
function reservationOK(idx, saveHTML){
	//alert(saveHTML);
	if(confirm("메일 확실히 예약 하셨습니까??")){
		FrameCKP.location.href='/admin/mailzine/mailzine_siteyn.asp?idx='+idx+'&reservationOK=OK&saveHtml=' + saveHTML+'&menupos=<%= menupos %>';
	}
}

function jsMailzineCode() {
	var winCodeView = window.open('/admin/mailzine/code/PopManageCode.asp','codeview','width=1400,height=800,scrollbars=yes,resizable=yes');
	winCodeView.focus();
}

function jsMailzineTemplate() {
	var winTemplateView = window.open('/admin/mailzine/code/PopManageTemplate.asp','templateview','width=1400,height=800,scrollbars=yes,resizable=yes');
	winTemplateView.focus();
}

// 전체 선택,취소
function chgSel_on_off(){
	var frm = document.monthly;
	if (frm.lineSel.length){
		for(var i=0;i<frm.lineSel.length;i++)
		{
			frm.lineSel[i].checked=frm.tt_sel.checked;
		}
	}else{
		frm.lineSel.checked=frm.tt_sel.checked;
	}
}

// 선택된 항목 삭제/복구
function fnDeleteMail(){
	var i, chk=0;
	var frm = document.monthly;


	if (frm.lineSel.length){
		for(i=0;i<frm.lineSel.length;i++)
		{
			if(frm.lineSel[i].checked)
			{
				chk++;
			}
		}
	}else{
			if(frm.lineSel.checked)
			{
				chk++;
			}
	}

	if(chk==0){
		alert("한 개 이상 선택해주십시오.");
		return;
	}else{
		if(confirm("선택하신 " + chk + "개의  항목을 삭제 하시겠습니까?")){
			frm.mode.value="delete";
			frm.target="FrameCKP";
			frm.action="mailzine_siteyn.asp";
			frm.submit();
		}else{
			return;
		}
	}
}

</script>

<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mailergubun" value="<%= mailergubun %>">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="#EEEEEE">검색<br>조건</td>
	<td align="left">
		* 메일러구분 : <%= mailergubun %>
		&nbsp;&nbsp;
		* 발송기간 :
        <input id="sDt" name="sDt" value="<%=sDt%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
        <input id="eDt" name="eDt" value="<%=eDt%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		&nbsp;&nbsp;
		* 제목 :
		<input type="text" class="text" name="SearchKey" value="<%=SearchKey%>" size="20">
		&nbsp;&nbsp;
		* 노출여부 :
		<% Drawisusing "isusing" , isusing , "class='select'" %>
		&nbsp;&nbsp;
		* 발송지역 :
		<% Drawareagubun "area" , area , "class='select'" %>
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "sDt", trigger    : "sDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "eDt", trigger    : "eDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
	<td width="50" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="검색" onClick="goPage('');">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<% mailzine_member_count %><Br><br><% mailzine_notmember_count %>
	</td>
	<td align="right"></td>
</tr>
<tr>
	<td align="left">
		<input type="button" value="선택삭제" onclick="fnDeleteMail();" class="button">
	</td>
	<td align="right">
		<input type="button" value="신규등록(템플릿)" onclick="jsModifyNewMailzine(-1);" class="button">
		<input type="button" value="신규등록(수기)" onclick="editreg('');" class="button">
		<!--<input type="button" value="신규등록(자동)" onclick="jsModifyMailzine(-1);" class="button">-->
		&nbsp;
		<input type="button" value="코드관리" onclick="jsMailzineCode();" class="button">
		<input type="button" value="템플릿관리" onclick="jsMailzineTemplate();" class="button">

		<% if C_ADMIN_AUTH then %>
			<Br>관리자권한:
			(<input type="button" class="button" value="이메일기본폼복사" onclick="mailCodeView('','basicMailFormCopy');">
			<input type="button" value="블랙리스트재작성(관리자)" onclick="blackListReg();" class="button">)
		<% end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<form method="post" name="monthly" style="margin:0px;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="mode" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= omail.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= omail.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=20><input type="checkbox" name="tt_sel" onclick="chgSel_on_off()"></td>
	<td width=60>No</td>
	<td width=60>발송일</td>
	<td>Title</td>
	<td width=90>작성구분</td>
	<td width=50>디자인<Br>완성여부</td>
	<td width=80>최종<Br>수정일</td>
	<td width=80>예약완료시간</td>
	<td width=40>사이트<Br>노출</td>
	<td width=100>발송지역</td>
	<td width=80>발송회원등급</td>
	<td width=40>메일러</td>
	<td width=40>비고</td>
	<td width=220>코드추출</td>
</tr>
<% if omail.FresultCount>0 then %>
<% for ix=0 to omail.FresultCount-1 %>
<tr align="center" <% if omail.FItemList(ix).farea="ten_china" then %>bgcolor="<%= adminColor("dgray") %>"<% elseif (isnull(omail.FItemList(ix).FreservationDATE) <> False ) AND (omail.FItemList(ix).farea="finger_all") Then %>bgcolor="<%= adminColor("pink") %>"<% elseif (isnull(omail.FItemList(ix).FreservationDATE) <> False ) AND (omail.FItemList(ix).farea="ten_all") Then %>bgcolor="<%= adminColor("green") %>"<% else %>bgcolor="#FFFFFF"  onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";<% end if %> >
	<td><input type="checkbox" name="lineSel" value="<% = omail.FItemList(ix).Fidx %>"></td>
	<td><% = omail.FItemList(ix).Fidx %></td>
	<td><% = omail.FItemList(ix).Fregdate %></td>
	<td align="left">(광고) <% = omail.FItemList(ix).Ftitle %></td>
	<td>
		<% if omail.FItemList(ix).Fregtype2="0" then %>
			<%= omail.FItemList(ix).GetRegTypeName() %>
		<% else %>
			<%= GetRegNewTypeName(omail.FItemList(ix).Fregtype2) %>
		<% end if %>
	</td>
	<td>
		<% if omail.FItemList(ix).Fgubun = "5" then %>
			완성
		<% else %>
			미완성
		<% end if %>
	</td>
	<td>
		<%= left(omail.FItemList(ix).Flastupdate,10) %>
		<br><%= mid(omail.FItemList(ix).Flastupdate,11,12) %>
	</td>
	<td>
		<% If (C_ADMIN_AUTH or C_SYSTEM_Part or C_MD or C_MKT_Part) AND (omail.FItemList(ix).Fgubun = "5") AND (isnull(omail.FItemList(ix).FreservationDATE) <> False ) Then %>
			<input type="button" value="발송" onclick="javascript:reservationOK('<% = omail.FItemList(ix).Fidx %>', '<%= CHKIIF(omail.FItemList(ix).Fregtype <> "1", "Y", "N")%>');" class="button"><br>
		<% End If %>

		<%= Chkiif(isnull(omail.FItemList(ix).FreservationDATE), "예약 전", left(omail.FItemList(ix).FreservationDATE,10)&"<br>"&mid(omail.FItemList(ix).FreservationDATE,11,12)) %>
	</td>
	<td>
		<% = omail.FItemList(ix).Fisusing %>
		<Br>
		<input type="button" value="변경" class="button" onclick="siteyn('<% = omail.FItemList(ix).Fidx %>','<% = omail.FItemList(ix).Fisusing %>', '<%= omail.FItemList(ix).Fgubun %>');">
	</td>
	<td>
		<%= getareagubun(omail.FItemList(ix).farea) %>
	</td>
	<td>
		<% = omail.FItemList(ix).fmemgubun %>
	</td>
	<td><% = omail.FItemList(ix).fmailergubun %></td>
	<td>
		<% if omail.FItemList(ix).Fregtype2<>"0" then %>
			<input type="button" value="수정" class="button" onclick="jsModifyNewMailzine(<% = omail.FItemList(ix).Fidx %>);">
		<% elseif omail.FItemList(ix).Fregtype2="0" and omail.FItemList(ix).Fregtype <> "1" then %>
			<input type="button" value="수정" class="button" onclick="jsModifyMailzine(<% = omail.FItemList(ix).Fidx %>);">
		<% else %>
			<input type="button" value="수정" class="button" onclick="editreg(<% = omail.FItemList(ix).Fidx %>);">
		<% end if %>
	</td>
	<td>
		<%
		' 신형 템플릿 메일
		if omail.FItemList(ix).Fregtype2<>"0" then
		%>
			<input type="button" value="미리보기(회원)" class="button" onclick="displayTemplates(<% = omail.FItemList(ix).Fidx %>,'member', 'view');">
			<input type="button" value="미리보기(비회원)" class="button" onclick="displayTemplates(<% = omail.FItemList(ix).Fidx %>,'notmember', 'view');">
			<Br>
			<input type="button" value="코드(회원/비회원)" class="button" onclick="displayTemplates(<% = omail.FItemList(ix).Fidx %>,'member', 'code');">
			<input type="button" value="코드(테스트발송)" class="button" onclick="displayTemplates(<% = omail.FItemList(ix).Fidx %>,'test', 'code');" <% if mailergubun<>"TMS" then response.write " disabled" %>>
		<%
		' 자동메일
		elseif ((omail.FItemList(ix).Fregtype2="0") or omail.FItemList(ix).Fregtype2="") and omail.FItemList(ix).Fregtype <> "1" then
		%>
			<input type="button" value="미리보기(회원)" class="button" onclick="displayNew(<% = omail.FItemList(ix).Fidx %>,'member', 'view');">
			<input type="button" value="미리보기(비회원)" class="button" onclick="displayNew(<% = omail.FItemList(ix).Fidx %>,'notmember', 'view');">
			<Br>
			<input type="button" value="코드(회원)" class="button" onclick="displayNew(<% = omail.FItemList(ix).Fidx %>,'member', 'code');">
			<input type="button" value="코드(비회원)" class="button" onclick="displayNew(<% = omail.FItemList(ix).Fidx %>,'notmember', 'code');">
		<%
		' 수기등록메일
		else
		%>
			<input type="button" value="미리보기(회원)" class="button" onclick="displayManual(<% = omail.FItemList(ix).Fidx %>,'member');">
			<input type="button" value="미리보기(비회원)" class="button" onclick="displayManual(<% = omail.FItemList(ix).Fidx %>,'notmember');">
			<Br>
			<input type="button" value="코드(회원)" class="button" onclick="mailCodeView(<% = omail.FItemList(ix).Fidx %>,'member');">
			<input type="button" value="코드(비회원)" class="button" onclick="mailCodeView(<% = omail.FItemList(ix).Fidx %>,'notmember');">
		<% end if %>
	</td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
       	<% if omail.HasPreScroll then %>
			<span class="list_link"><a href="javascript:goPage(<%= omail.StarScrollPage-1 %>)">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for ix = 0 + omail.StarScrollPage to omail.StarScrollPage + omail.FScrollCount - 1 %>
			<% if (ix > omail.FTotalpage) then Exit for %>
			<% if CStr(ix) = CStr(omail.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= ix %></b></font></span>
			<% else %>
			<a href="javascript:goPage(<%=ix%>)" class="list_link"><font color="#000000"><%= ix %></font></a>
			<% end if %>
		<% next %>
		<% if omail.HasNextScroll then %>
			<span class="list_link"><a href="javascript:goPage(<%=ix%>)">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
</form>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe name="FrameCKP" src="" frameborder="0" width="100%" height="400" ></iframe>
<% else %>
	<iframe name="FrameCKP" src="" frameborder="0" width="0" height="0" ></iframe>
<% end if %>

<%
set omail = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbTMSclose.asp" -->