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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mailzinecls.asp"-->
<%
CONST MAXHeightPX = 1400    '''이 수치에 대해서는 확실하지 않음.. (2,000px 보다 작은 사이즈 2개를 넣었을때 도 깨진경우가 있음)

dim idx, mode, mailergubun
	idx = requestCheckVar(request("idx"),32)
	mailergubun = requestcheckvar(request("mailergubun"),16)

if (idx = "") then
	idx = -1
end if

if (idx > 0) then
	mode = "modi"
else
	mode = "ins"
end if

if mailergubun="" or isnull(mailergubun) then
	response.write "메일러 구분이 없습니다."
	dbget.close() : response.end
end if

dim omail
set omail = new CMailzineList
	omail.frectidx = idx
	''omail.FRectRegType = "2"
	omail.frectmailergubun = "EMS"
	omail.MailzineDetail()

if (omail.FOneItem.Fregtype = "") then
	omail.FOneItem.Fregtype = "2"
end if

%>
<style>
#mask {
	position:absolute;
	z-index:9000;
	background-color:#000;
	display:none;
	left:0;
	top:0;
}
.window{
	display: none;
	position:absolute;
	left:100px;
	bottom:10px;
	z-index:10000;
}
</style>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/JavaScript">

function jsSubmit(frm) {
	var regtype = jsGetRegType();

	if (frm.title.value == "") {
		alert('제목을 입력하세요.');
		return;
	}

	if (frm.regdate.value == "") {
		alert('메일발송 예정일을 입력하세요.');
		return;
	}

	if (frm.area.value == "") {
		alert('발송지역을 입력하세요.');
		return;
	}

	if (frm.memgubun.value == "") {
		alert('회원등급을 입력하세요.');
		return;
	}

	if (frm.isusing.value == "") {
		alert('사이트 노출여부를 입력하세요.');
		return;
	}

	if (frm.secretGubun.value == "") {
		alert('시크릿구분을 입력하세요.');
		return;
	}

	if (regtype == '2') {
		if (frm.evt_code.value == "") {
			alert('주말특가 이벤트코드를 입력하세요.');
			return;
		}

		if (frm.evt_code.value*0 != 0) {
			alert('잘못된 주말특가 이벤트코드입니다.');
			return;
		}

		if (frm.img1editname.value == "") {
			alert('가져오기 버튼을 누르세요.');
			return;
		}
	}

	if (frm.mode.value == "ins") {
		// 신규 등록시
		if (frm.isusing.value == "Y") {
			alert('신규 등록시에는 사이트 노출을 선택할 수 없습니다.');
			return;
		}
		if (frm.gubun.value == "5") {
			alert('신규 등록시에는 메일진 작성상태를 완료로 할 수 없습니다.');
			return;
		}
	}

	if (confirm('저장하시겠습니까?') == true) {
		frm.submit();
	}

}

function jsGetRegType() {
	return $('input:radio[name="regtype"]:checked').val();
}

function jsSetDisabledObj(obj, disabled) {
	obj.disabled = disabled;
	if (obj.type != 'textarea') {
		obj.style.background = disabled ? '#DDDDDD' : '#FFFFFF';
	}
}

function jsSetItemState() {
	var frm = document.frm;
	var regtype = jsGetRegType();
	if (regtype == undefined) { return; }

	jsSetDisabledObj(frm.evt_code, false);

	if (regtype == '2') {
		jsSetDisabledObj(frm.img2editname, false);
	} else if (regtype == '3') {
		jsSetDisabledObj(frm.img2editname, true);
	} else if (regtype == '4') {
		jsSetDisabledObj(frm.img2editname, false);

	// 다이어리스토리
	} else if (regtype == '5') {
		jsSetDisabledObj(frm.evt_code, true);
		jsSetDisabledObj(frm.img1editname, false);
		jsSetDisabledObj(frm.img2editname, false);
	}

	jsSetDisabledObj(frm.img3editname, false);
	jsSetDisabledObj(frm.img4editname, false);
}

function jsGetList() {
	var frm = document.frm;
	var regtype = jsGetRegType();

	if (frm.regdate.value == "") {
		alert('메일발송 예정일을 입력하세요.');
		return;
	}

	if (regtype!='5'){
		if (frm.evt_code.value == "") {
			alert('메인 이벤트코드를 입력하세요.');
			return;
		}
		if (frm.evt_code.value*0 != 0) {
			alert('잘못된 메인 이벤트코드입니다.');
			return;
		}
	}

	document.iframe_proc.location.href = '/admin/mailzine/mailzine_process.asp?mode=getlist&regdate=' + frm.regdate.value + '&regtype=' + regtype + '&evt_code=' + frm.evt_code.value;
}

$(document).ready(function(){
	jsSetItemState();
});

</script>

<form name="frm" method="post" action="/admin/mailzine/mailzine_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="mailergubun" value="<%= mailergubun %>">
<table width="95%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150">메일진 제목</td>
	<td><input type="text" name="title" class="input" size="55" value="<%= omail.FOneItem.Ftitle %>" /> * 고객명 : ${EMS_M_NAME}</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150"><b>메일진 발송예정일</b></td>
	<td>
		<input id="regdate" name="regdate" value="<%= omail.FOneItem.Fregdate %>" class="text_ro" size="10" maxlength="10" readonly /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="regdate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
		var regdate = new Calendar({
			inputField : "regdate", trigger    : "regdate_trigger",
			onSelect: function() {
				this.hide();
			}, bottomBar: true, dateFormat: "%Y.%m.%d", fdow: 0
		});
		</script>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150"><b>메일진 종류</b></td>
	<td>
		<input type="radio" name="regtype" value="2" <% if omail.FOneItem.Fregtype = "2" then response.write "checked"%> onClick="jsSetItemState()" <%= CHKIIF(mode="modi" and omail.FOneItem.Fregtype <> "2", "disabled", "") %>> 주말특가
		&nbsp;
		<input type="radio" name="regtype" value="3" <% if omail.FOneItem.Fregtype = "3" then response.write "checked"%> onClick="jsSetItemState()" <%= CHKIIF((mode="modi" and omail.FOneItem.Fregtype <> "3"), "disabled", "") %>> 기획전
		&nbsp;
		<input type="radio" name="regtype" value="4" <% if omail.FOneItem.Fregtype = "4" then response.write "checked"%> onClick="jsSetItemState()" <%= CHKIIF((mode="modi" and omail.FOneItem.Fregtype <> "4"), "disabled", "") %>> 기획전+MD's Pick
		&nbsp;
		<input type="radio" name="regtype" value="5" <% if omail.FOneItem.Fregtype = "5" then response.write "checked"%> onClick="jsSetItemState()" <%= CHKIIF((mode="modi" and omail.FOneItem.Fregtype <> "5"), "disabled", "") %>> 다이어리스토리
		&nbsp;
		<input type="button" class="button" value="가져오기" onClick="jsGetList()">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">발송지역</td>
	<td>
		<% Drawareagubun "area" , omail.FOneItem.Farea , "class='select'" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">발송회원등급</td>
	<td>
		<% DrawMemberGubun "memgubun" , omail.FOneItem.Fmemgubun , "class='select'" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">사이트노출</td>
	<td>
		<% Drawisusing "isusing" , omail.FOneItem.Fisusing , "class='select'" %> * 즉시 고객에게 노출됩니다.
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">시크릿 적용</td>
	<td>
		<% DrawsecretGubun "secretGubun" , omail.FOneItem.FsecretGubun , "class='select'" %> * 사이트노출시, 시크릿 적용을 Y로 두면 타이틀만 노출되고 클릭이 되지 않습니다.
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150"><b>메일진 작성상태</b></td>
	<td>
		<select name="gubun" class="select">
			<option value="1" <% if omail.FOneItem.Fgubun = "1" then response.write "selected"%>>미완성</option>
			<option value="5" <% if omail.FOneItem.Fgubun = "5" then response.write "selected"%>>완성</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="1">
	<td align="center" width="150"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150">메인 이벤트코드</td>
	<td>
		<input type="text" name="evt_code" class="input" size="12" value="<%= omail.FOneItem.Fevt_code %>"> * 주말특가 또는 메인 이벤트코드
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150">기획전 이벤트코드목록</td>
	<td>
		<textarea class="textarea" cols="20" rows="6" name="img1editname"><%= omail.FOneItem.Fimgmap1 %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150">MD Pick 상품목록</td>
	<td>
		<textarea class="textarea" cols="20" rows="6" name="img2editname"><%= omail.FOneItem.Fimgmap2 %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150">JUST 1DAY</td>
	<td>
		<input type="text" class="input" name="img3editname" size="20" value="<%= omail.FOneItem.Fimgmap3 %>"  readonly />
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150">텐바이텐 클래스</td>
	<td>
		<textarea class="textarea" cols="20" rows="3" name="img4editname" readonly><%= omail.FOneItem.Fimgmap4 %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="1">
	<td align="center" width="150"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150">발송예약 완료일</td>
	<td>
		<%= omail.FOneItem.FreservationDATE %> * 예약완료 이후에는 [상품운영팀]이슬비에게 연락해야만 수정내용이 반영됩니다.
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150">작성자(최종수정)</td>
	<td>
		<%= CHKIIF(IsNull(omail.FOneItem.Fmodiuserid), omail.FOneItem.Freguserid, omail.FOneItem.Fmodiuserid) %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="50">
	<td align="center" colspan="2">
		<input type="button" class="button" value="저장" onClick="jsSubmit(document.frm);">
	</td>
</tr>
</table>
</form>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe name="iframe_proc" width="100%" height="400" frameborder="0"></iframe>
<% else %>
	<iframe name="iframe_proc" width="0" height="0" frameborder="0"></iframe>
<% end if %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
