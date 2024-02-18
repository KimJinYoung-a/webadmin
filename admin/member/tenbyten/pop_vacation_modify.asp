<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 휴가관리
' History : 2011.01.19 정윤정 생성
'			2022.09.21 한용민 수정(오류수정)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%

dim userid
dim isyearvacation
dim oMember
dim divcd, startday, endday, totalvacationday
dim joinday

dim i

userid = requestCheckvar(request("userid"),32)

'// 로그인정보(등급)에 따라 기본 부서 설정(파트선임 이상:3 ,개발팀 or 운영개발팀 및 인사총무팀:20 제외)
if Not((session("ssAdminLsn")<=3 and C_SYSTEM_Part) or C_PSMngPart) then
	response.write "파트선임 이상 및 인사총무팀만 휴가를 등록할 수 있습니다."
	response.end
end if



'==============================================================================
dim yearvacation_startday, yearvacation_endday

yearvacation_startday = Cstr(Year(now())) & "-01-01"
yearvacation_endday = Cstr(Year(now()) + 1) & "-03-31"

%>
<html>
<head>
<title>연차(휴가) 등록</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/bct.css" type="text/css">
<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">

var STARTDAY, ENDDAY;

function SaveVacation() {
	var frm = document.frm;

	if ((frm.userid.value.length < 1) && (frm.empno.value.length < 1)) {
		alert("WEBADMIN 아이디 또는 사번을 입력해주십시요.");
		frm.userid.focus();
		return false;
	}

	if (frm.divcd.value == "") {
		alert("구분을 입력해주십시요.");
		return false;
	}

	if ((frm.startday.value == "") || (frm.endday.value == "")) {
		alert("사용가능기간을 입력해주십시요.");
		return false;
	}

	if (frm.totalvacationday.value.length < 1) {
		alert("휴가일수를 입력해주십시요.");
		frm.totalvacationday.focus();
		return false;
	}

	if (frm.totalvacationday.value*0 != 0) {
		alert("휴가일수는 숫자만 입력가능합니다.");
		frm.totalvacationday.focus();
		return false;
	}

	if (frm.totalvacationday.value <= 0) {
		alert("휴가일수는 1 이상이어야 합니다.");
		frm.totalvacationday.focus();
		return false;
	}

	if (checkDate() == false) { return false; }

	if(confirm("등록 하시겠습니까?")) {
		frm.submit();
	}
}

function checkDate() {
	var frm = document.frm;

	var startday = frm.startday.value;
	var endday = frm.endday.value;
	var totalvacationday = frm.totalvacationday.value;

	var startdate = toDate(startday);
	var enddate = toDate(endday);

	var tmp;

	if (startdate > enddate) {
		alert("종료일이 시작일보다 과거날짜입니다.");
		return false;
	}

	return true;
}

function SetYearVacation() {
	var frm = document.frm;

	frm.startday.value = "";
	frm.endday.value = "";

	if (frm.employtype.value == "") {
		alert("먼저 아이디 또는 사번을 확인하세요.");
		frm.divcd.value = "";
		return;
	}

	if (frm.divcd.value == "1") {
		frm.startday.value = STARTDAY;
		frm.endday.value = ENDDAY;
	}
}

function SubmitSearchEmployType()
{
	var frm = document.frm;

	ResetEmployType();

	if ((frm.userid.value.length < 1) && (frm.empno.value.length < 1)) {
		alert("WEBADMIN 아이디 또는 사번을 입력해주십시요.");
		frm.userid.focus();
		return false;
	}

	if ((frm.userid.value.length >= 1) && (frm.empno.value.length >= 1)) {
		if (confirm("아이디와 사번이 모두 입력되었습니다.\n아이디를 기준으로 계약구분을 확인합니다.\n\n진행하시겠습니까?") != true) {
			return;
		}
		frm.empno.value = "";
	}
 
	var ifr = document.getElementById("ifremploytype");
	ifr.src = "domodifyvacation.asp?mode=chkemploytype&userid=" + frm.userid.value + "&empno=" + frm.empno.value;
}


function ResetEmployType() {
	var frm = document.frm;

	frm.employtype.value = "";
	frm.userid1.value = "";
	frm.empno1.value = "";

	STARTDAY = "";
	ENDDAY = "";
}


function ReActEmployType(resultval, empno, userid, posit_sn)
{
	var frm = document.frm;

	frm.empno.value = empno;
	frm.userid.value = userid;
	frm.posit_sn.value = posit_sn;
	
	// 시작일은 등록하는 달의 1일
	var s = new Date();
	s.setDate(1);
	STARTDAY = toDateString(s);

	var e = new Date();

	switch (resultval) {
		case 1:
			// 정규직
			frm.employtype.value = "정규직";
			frm.userid1.value = frm.userid.value;
			frm.empno1.value = frm.empno.value;

			// 다음해 3월 말일
			e.setYear(s.getFullYear() + 1);
			e.setMonth(3 - 1);
			e.setDate(31);
			break;
		case 2:
			// 계약직
			frm.employtype.value = "계약직";
			frm.userid1.value = frm.userid.value;
			frm.empno1.value = frm.empno.value;

			e.setYear(s.getFullYear() + 2);
			e.setMonth(empno.substring(6,8) - 1);
			e.setDate(empno.substring(8,10));
			e = new Date(e.getTime() - 1 * 24 * 60 * 60 * 1000); // 전날
			break;
		default:
			//
	}
  
	ENDDAY = toDateString(e);
	
	if (posit_sn ==13){
		document.all.divyv.style.display= "";
	}else{
		document.all.divyv.style.display= "none;";
	}
	
}

function chkCalVacation(){
	frmCal.empno.value = frm.empno.value;
	frmCal.target = "ifremploytype";
	frmCal.submit(); 
}

</script>
</head>
<body leftmargin="5" topmargin="5">
<form name="frmCal" method="post" action="domodifyvacation.asp">
	<input type="hidden" name="mode" value="calYV">
	<input type="hidden" name="empno" value="">
</form>
<form name="frm" method="post" action="domodifyvacation.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="add">
<input type="hidden" name="posit_sn" value="">
<table width="470" border="0" cellpadding="2" cellspacing="1" align="center" class="a" bgcolor=#BABABA> 
	<tr height="25">
		<td valign="bottom" colspan=2  bgcolor="F4F4F4">
			<font color="red"><strong>연차(휴가) 등록</strong></font>
		</td>
	</tr>
	<tr align="left" height="25">
		<td width=120 bgcolor="<%= adminColor("tabletop") %>">WEBADMIN 아이디</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="userid" class="text" size="16" value="<%= userid %>" onChange="ResetEmployType()"> <input type="button" class="button" value="확인" onclick="SubmitSearchEmployType()">
		</td>
	</tr>
	<tr align="left" height="25">
		<td width=120 bgcolor="<%= adminColor("tabletop") %>">사번</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="empno" class="text" size="20" value="" onChange="ResetEmployType()"> <input type="button" class="button" value="확인" onclick="SubmitSearchEmployType()">
		</td>
	</tr>
	</tr>
	<tr align="left" height="25">
		<td width=120 bgcolor="<%= adminColor("tabletop") %>">계약구분</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="employtype" class="text_ro" size="6" readonly>
			<input type="hidden" name="empno1">
			<input type="hidden" name="userid1">
		</td>
	</tr>
	<tr align="left" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">구분</td>
		<td bgcolor="#FFFFFF">
			<select class="select" name=divcd onchange="SetYearVacation();">
				<option value="">====</option>
				<option value="1" <% if (divcd = "1") then %>selected<% end if %>>연차</option>
				<!--
				<option value="2">월차</option>
				-->
				<option value="3">포상</option>
				<option value="4">위로</option>
				<option value="6">경조사</option>
				<option value="5">장기</option>
				<option value="7">휴일대체</option>
				<option value="8">기타휴가</option>
				<option value="9">보상휴가</option>
				<option value="A">생일휴가</option>
			</select>
		</td>
	</tr>
	<tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">사용가능기간</td>
    	<td bgcolor="#FFFFFF">
    		<input id="sDt" name="startday" value="<%=startday%>" class="text" size="10" maxlength="10" />
			-
			<input id="eDt" name="endday" value="<%=endday%>" class="text" size="10" maxlength="10" />
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "sDt", trigger    : "sDt",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_End.args.min = date;
						CAL_End.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
				var CAL_End = new Calendar({
					inputField : "eDt", trigger    : "eDt",
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
	<tr align="left" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">휴가일수</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="totalvacationday" class="text" size="4" maxlength="4" value="<%= totalvacationday %>">
			<span style="display:none;" id="divyv">시간 <input type="button" class="button" value="계산" onClick="chkCalVacation();"></span>
		</td>
	</tr>
	<tr align="center" height="25">
		<td colspan="2" bgcolor="#FFFFFF">
			<input type="button" class="button" value="확인" onclick="SaveVacation()">
			<input type="button" class="button" value="취소" onClick="self.close()">
		</td>
	</tr>
</table><br>
</form>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe src="" id="ifremploytype" name="ifremploytype" frameborder="0" width="100%" height="300">
<% else %>
	<iframe src="" id="ifremploytype" name="ifremploytype" frameborder="0" width="0" height="0">
<% end if %>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
