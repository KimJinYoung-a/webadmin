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
<!-- #include virtual="/lib/db/dbTMSopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mailzinenewcls.asp"-->
<%
CONST MAXHeightPX = 1400    '''이 수치에 대해서는 확실하지 않음.. (2,000px 보다 작은 사이즈 2개를 넣었을때 도 깨진경우가 있음)

dim idx, mode, eregdate, mailergubun
	idx = requestCheckVar(request("idx"),32)
	mailergubun = requestcheckvar(request("mailergubun"),16)

eregdate=now()

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
	omail.frectmailergubun = mailergubun
	omail.MailzineDetail()

if (omail.FOneItem.Fregtype2 = "") then
	omail.FOneItem.Fregtype2 = "101"
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
	if (frm.regtype.value == "") {
		alert('메일진 종류를 선택 하세요.');
		return;
	}
	if (confirm('저장하시겠습니까?') == true) {
		frm.submit();
	}

}

function jsGetRegType() {
	return $("select[name=regtype]").val();
}

function jsSetDisabledObj(obj, disabled) {
	obj.disabled = disabled;
	if (obj.type != 'textarea') {
		obj.style.background = disabled ? '#DDDDDD' : '#FFFFFF';
	}
}

function jsSetItemState(objvalue) {
	var frm = document.frm;
	var regtype = objvalue;

	if (regtype == undefined) { return; }

    if (frm.regdate.value == "") {
        alert('메일발송 예정일을 입력하세요..');
        return;
    }
    var url="ajaxTemplateLoad.asp";
    var params='mode=getlist&regdate=' + frm.regdate.value + '&regtype=' + regtype + '&idx=' + frm.idx.value;
    $.ajax({
        type:"POST",
        url:url,
        data:params,
        success:function(args){ 
            $("#tempMail").html(args);
        },
        error:function(e){
            alert("데이터로딩에 문제가 생겼습니다. 시스템팀에 문의해주세요");
        }
    });
}

function jsGetList() {
	var frm = document.frm;
	var regtype = jsGetRegType();
	var codecheck = fnEvtCodeCheck();

	if (frm.regdate.value == "") {
		alert('메일발송 예정일을 입력하세요...');
		return;
	}
	if(!codecheck){
		return;
	}

	document.iframe_proc.location.href = '/admin/mailzine/template/mailzine_process.asp?mode=getlist&regdate=' + frm.regdate.value + '&regtype=' + regtype + '&evt_code=' + frm.arrevtcode.value;
}

function delimg(imgnumber){
	frm_mail.target="iframe_proc";
	frm_mail.action = '/admin/mailzine/template/mailzine_process.asp';
	frm_mail.imgnumber.value=imgnumber;
	frm_mail.mode.value='imgdel';
	frm_mail.submit();
}

function jsSetImg(sImg, sName){ 
	var winImg;
	winImg = window.open('/admin/mailzine/template/pop_uploadimg.asp?yr=<%=Year(eregdate)%>&sImg='+sImg+'&sName='+sName,'popImg','width=370,height=150');
	winImg.focus();
}
<% if (idx > 0) then %>
$(function(){
	jsSetItemState(<%=omail.FOneItem.Fregtype2%>);
});
<% end if %>
</script>

<form name="frm" method="post" action="/admin/mailzine/template/mailzine_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="mailergubun" value="<%= mailergubun %>">
<input type="hidden" name="arrevtcode">
<table width="95%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150">메일진 제목</td>
	<td><input type="text" name="title" class="input" size="55" value="<%= omail.FOneItem.Ftitle %>" /> * 고객명 : ${EMS_M_NAME}</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150"><b>메일진 발송예정일</b></td>
	<td>
		<input id="regdate" name="regdate" value="<%= omail.FOneItem.Fregdate %>" class="text_ro" size="10" maxlength="10" readonly /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="regdate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
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
		<% DrawMailzineKind "regtype", omail.FOneItem.Fregtype2, "onChange='jsSetItemState(this.value);'" %>
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
</table>

<div id="tempMail"></div>


<table width="95%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150">작성자(최종수정)</td>
	<td>
		<%= CHKIIF(IsNull(omail.FOneItem.Fmodiuserid), omail.FOneItem.Freguserid, omail.FOneItem.Fmodiuserid) %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="50">
	<td align="center" colspan="2">
		<input type="button" class="button" value=" 저 장 하 기 " onClick="jsSubmit(document.frm)">
	</td>
</tr>
</table>
</form>
<form name="frm_mail" method="post">
	<input type="hidden" name="idx" value="<% = idx %>">
	<input type="hidden" name="imgnumber">
	<input type="hidden" name="mode">
</form>
<% IF application("Svr_Info")="Dev" THEN %>
	<iframe name="iframe_proc" width="100%" height="400" frameborder="0"></iframe>
<% else %>
	<iframe name="iframe_proc" width="0" height="0" frameborder="0"></iframe>
<% end if %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbTMSclose.asp" -->