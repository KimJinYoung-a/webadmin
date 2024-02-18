<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이벤트관리
' History : 2016.08.08 김진영 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/event/eventCls.asp"-->
<%
Dim idx, oEvent
Dim gubun, actid, company_name, evt_startdate, evt_enddate, contentsCode, evt_name, isusing

idx = RequestCheckvar(request("idx"),10)

If idx <> "" Then
	Set oEvent = new CEvent
		oEvent.FRectIdx = idx
		oEvent.getEventOneItem

		gubun			= oEvent.FOneItem.FGubun
		actid			= oEvent.FOneItem.FActid
		company_name	= oEvent.FOneItem.FCompany_name
		evt_startdate	= oEvent.FOneItem.FEvt_startdate
		evt_enddate		= oEvent.FOneItem.FEvt_enddate
		contentsCode	= oEvent.FOneItem.FContentsCode
		evt_name		= oEvent.FOneItem.FEvt_name
		isusing			=  oEvent.FOneItem.FIsusing
	Set oEvent = nothing
End If

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script>
function fngubunChg(v, i){
	if(v != ''){
		$.ajax({
			url: "/academy/events/ajax_gubunTR.asp?gubun="+v+"&idx="+i,
			dataType : "html",
			type : "get",
			success : function(result){
				$("#cghGubun").empty().html(result);
				$("#cghGubun").show();
			}
		});
	}else{
		$("#cghGubun").empty()
		$("#cghGubun").hide()
	}
}
//강사 전체보기 팝업
function pop_lecture(v){
    var popwin = window.open("/academy/events/pop_lecturerList.asp?gubun="+v+"","pop_lecture","width=500,height=700,scrollbars=yes,resizable=yes");
	popwin.focus();
}
//진행중인 강좌보기 팝업
function pop_lec(){
	var lecturer_id = $("#lecid").val();
	if(lecturer_id == ''){
		alert('강사를 먼저 선택하세요');
		return false;
	}
    var popwin2 = window.open("/academy/events/pop_lecList.asp?lecturer_id="+lecturer_id+"","pop_lec","width=900,height=700,scrollbars=yes,resizable=yes");
	popwin2.focus();
}
//판매중인 작품보기 팝업
function pop_art(){
	var makerid = $("#tecid").val();
	if(makerid == ''){
		alert('작가를 먼저 선택하세요');
		return false;
	}
    var popwin3 = window.open("/academy/events/pop_artList.asp?makerid="+makerid+"","pop_art","width=900,height=700,scrollbars=yes,resizable=yes");
	popwin3.focus();
}
function frm_check(f){
	if($("#evt_startdate").val() == '' ){
		alert('이벤트 시작일을 입력하세요');
		return false;
	}

	if($("#evt_enddate").val() == '' ){
		alert('이벤트 종료일을 입력하세요');
		return false;
	}

	if( $("#evt_startdate").val() > $("#evt_enddate").val() ) {
		alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
		return false;
	}

	if($("#gubun").val() == '' ){
		alert('등록위치를 선택하세요');
		return false;
	}

	if($("#gubun").val() == 'D'){
		if($("#tecid").val() == '' ){
			alert('작가를 입력하세요');
			return false;
		}
		if($("#diycode").val() == ''){
			alert('작품코드를 입력하세요');
			return false;
		}
	}else if($("#gubun").val() == 'L'){
		if($("#lecid").val() == '' ){
			alert('강사를 입력하세요');
			return false;
		}
		if($("#lecidx").val() == ''){
			alert('강좌코드를 입력하세요');
			return false;
		}
	}

	if($("#evt_startdate").val() == $("#evt_enddate").val()){
		alert('동일 기간은 선택 불가능 합니다');
		$("#evt_enddate").val('');
		$("#evt_enddate").focus();
		return false;
	}
	
	if($("#evt_name").val() == '' ){
		alert('이벤트명을 입력하세요');
		return false;
	}

	if(confirm('저장 하시겠습니까?')){
		frm.action = '/academy/events/event_process.asp';
	<% If idx <> "" Then %>
		frm.mode.value = 'U';
	<% Else %>
		frm.mode.value = 'I';
	<% End If %>
		frm.submit();
	}
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="80" bgcolor="FFFFFF">
	<td><font size="4"><strong><%= Chkiif(idx <> "", "이벤트 수정", "신규 이벤트 등록") %></strong></font></td>
</tr>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="idx" value="<%= idx %>">
<col width="20%" />
<col width="" />
<tr height="30" bgcolor="FFFFFF">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">기간</td>
	<td>
		<input id="evt_startdate" readonly name="evt_startdate" class="text" value="<%= evt_startdate %>" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="evt_startdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> 00:00:00 ~
		<input id="evt_enddate" readonly name="evt_enddate" class="text" value="<%= evt_enddate %>" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="evt_enddate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> 00:00:00
		<script language="javascript">
		var CAL_Start = new Calendar({
			inputField : "evt_startdate", trigger    : "evt_startdate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_End.args.min = date;
				CAL_End.redraw();
				this.hide();
			}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
		var CAL_End = new Calendar({
			inputField : "evt_enddate", trigger    : "evt_enddate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_Start.args.max = date;
				CAL_Start.redraw();
				this.hide();
			}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
		</script>
	</td>
</tr>
<tr height="30" bgcolor="FFFFFF">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">등록위치</td>
	<td>
		<select name="gubun" id="gubun" class="select" onchange="fngubunChg(this.value, '');">
			<option value="">선택</option>
			<option value="D" <%= chkiif(gubun="D", "selected", "") %> >작가 프로필(작품)</option>
			<option value="L" <%= chkiif(gubun="L", "selected", "") %> >강사 프로필(강좌)</option>
		</select>
	</td>
</tr>
<tbody id="cghGubun" style="display:none;"></tbody>
<tr height="30" bgcolor="FFFFFF">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">이벤트명</td>
	<td><input type="text" class="text" name="evt_name" id="evt_name" value="<%= evt_name %>" size="50"></td>
</tr>
<tr height="30" bgcolor="FFFFFF">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">사용여부</td>
	<td>
		<label><input type="radio" name="isusing" value="Y" class="radio" <% If isusing="" OR isusing="Y" Then response.write "checked" %>>Y</label>
		<label><input type="radio" name="isusing" value="N" class="radio" <%= chkiif(isusing="N", "checked", "") %> >N</label>
	</td>
</tr>
<tr height="30" bgcolor="FFFFFF" align="center">
	<td colspan="2">
		<input type="button" class="button" value="저장" onclick="frm_check(this.frm);" style="color:red;font-weight:bold">
		<input type="button" class="button" value="취소" onclick="history.back(-1);">
	</td>
</tr>
</form>
</table>
<% If idx <> "" Then %>
<script>
	fngubunChg('<%= gubun %>', '<%= idx %>');
</script>
<% End If %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->