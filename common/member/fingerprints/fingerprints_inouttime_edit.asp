<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : 지문인식 근태관리
' Hieditor : 2011.03.22 한용민 생성
'            2012.02.15 허진원 - 미니달력 교체
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/member/fingerprints/fingerprints_cls.asp" -->

<%
dim idx , placeid , empno ,YYYYMMDD ,inoutType ,inoutTime ,posIdx ,posDate ,regdate
dim lasteditupdate, isusing ,ofingerprints ,inoutTimes, dispadmin, job_sn, loginjob_sn
	idx = requestCheckVar(request("idx"),10)
	dispadmin = false

set ofingerprints = new cfingerprints_list
	ofingerprints.frectidx = idx
	ofingerprints.ffingerprints_item()
	
	if ofingerprints.ftotalcount >0 then
		idx = ofingerprints.FOneItem.fidx
		placeid = ofingerprints.FOneItem.fplaceid
		empno = ofingerprints.FOneItem.fempno
		YYYYMMDD = ofingerprints.FOneItem.fYYYYMMDD
		inoutType = ofingerprints.FOneItem.finoutType
		inoutTime = ofingerprints.FOneItem.finoutTime
		inoutTime = left(inoutTime,10)
		inoutTimes = Num2Str(Hour(ofingerprints.FOneItem.finoutTime),2,"0","R") & ":" & Num2Str(Minute(ofingerprints.FOneItem.finoutTime),2,"0","R")& ":" & Num2Str(second(ofingerprints.FOneItem.finoutTime),2,"0","R")	
		posIdx = ofingerprints.FOneItem.fposIdx
		posDate = ofingerprints.FOneItem.fposDate
		regdate = ofingerprints.FOneItem.fregdate
		lasteditupdate = ofingerprints.FOneItem.flasteditupdate
		isusing	 = ofingerprints.FOneItem.fisusing
	end if

if inoutTime="" then inoutTime=date
if inoutTimes="" then inoutTimes="00:00:00"	

' 글쓴이의 직책.
job_sn=replace(replace(getjob_sn(empno, ""),"","0"),"0","9999")		' 선택안함은 9999
job_sn=replace(job_sn,"13","1")		' CFO -> CEO
' 로그인한사람의 직책.
loginjob_sn=replace(replace(getjob_sn(session("ssBctSn"), ""),"","0"),"0","9999")		' 선택안함은 9999
loginjob_sn=replace(loginjob_sn,"13","1")		' CFO -> CEO

' 관리자
if C_ADMIN_AUTH then
	dispadmin = true

' 직책 : CEO, CFO, 본부장
elseif loginjob_sn="1" or loginjob_sn="2" or loginjob_sn="3" then
	dispadmin = true
elseif job_sn>loginjob_sn then
	dispadmin = true
end if

%>

<script type='text/javascript'>

	function editfingerprints(){
	
		if (frmfingerprints.empno.value==''){
			alert('사원번호를 입력해 주세요');
			frmfingerprints.empno.focus();
			return;
		}
		
		if (frmfingerprints.inoutType.value==''){
			alert('상태를 선택해 주세요');
			frmfingerprints.inoutType.focus();
			return;
		}
		
		if (frmfingerprints.inoutTime.value==''){
			alert('날짜를 입력해 주세요');
			frmfingerprints.inoutTime.focus();
			return;
		}

		if (frmfingerprints.inoutTimes.value==''){
			alert('날짜를 시간 단위로 입력해 주세요');
			frmfingerprints.inoutTimes.focus();
			return;
		}
				
		if (frmfingerprints.isusing.value==''){
			alert('사용여부를 선택해 주세요');
			frmfingerprints.isusing.focus();
			return;
		}
		
		frmfingerprints.action='/common/member/fingerprints/fingerprints_inouttime_process.asp';
		frmfingerprints.mode.value='fingerprintsedit';
		frmfingerprints.submit();
	}
	
</script>

<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form method="post" name="frmfingerprints">
<input type="hidden" name="mode">
<tr bgcolor="FFFFFF" align="center">
	<td>
		번호
	</td>
	<td align="left">
		<%= idx %>
		<input type="hidden" name="idx" value="<%=idx%>">
	</td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td>
		사원번호
	</td>
	<td align="left">
		<input type="text" name="empno" value="<%=empno%>" size=32 maxlength=32>
	</td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td>
		상태
	</td>
	<td align="left">
		<% DrawinoutType "inoutType",inoutType,"" %>
	</td>
</tr>
<tr bgcolor="FFFFFF" align="center">	
	<td>
		출근일
	</td>
	<td align="left">
		<input id="inoutTime" name="inoutTime" value="<%=inoutTime%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="inoutTime_trigger" border="0" style="cursor:pointer" align="absmiddle" />
    	<input type="text" name="inoutTimes" size="8" maxlength="8" class="text" value="<%=inoutTimes%>">
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "inoutTime", trigger    : "inoutTime_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>	
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td>
		사용여부
	</td>
	<td align="left">
		<% Drawisusing "isusing",isusing,"" %>
	</td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td colspan="2">
		<% if dispadmin then %>
			<input type="button" onclick="editfingerprints();" class="button" value="저장">
		<% else %>
			본인의 근태는 본인이 수정하실수 없으며, 본인보다 상급자만 수정이 가능 합니다.
		<% end if %>
	</td>
</tr>
</form>
</table>

<%
set ofingerprints = nothing
%>	
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->