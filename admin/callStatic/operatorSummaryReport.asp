<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim tenUserID, vSessionID, use201502, research
dim i

Dim sDate	: sDate	= req("sDate", Date() - 30)
Dim eDate	: eDate	= req("eDate", Date())

research = req("research", "")
tenUserID = req("tenUserID", "")
vSessionID = session("ssBctId")
use201502 = req("use201502", "")


Dim strSql

if (use201502 = "Y") then
	strSql = " db_datamart.dbo.sp_Ten_Call_Report ('" & sDate & "', '" & eDate & "', '" & tenUserID & "')"
else
	strSql = " db_datamart.dbo.sp_Ten_Call_Report_NEW ('" & sDate & "', '" & eDate & "', '" & tenUserID & "')"
end if


db3_rsget.CursorLocation = adUseClient
db3_rsget.Open strSql,db3_dbget,adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

Dim rs
If Not db3_rsget.EOF Then
	rs = db3_rsget.getRows()
End If
db3_rsget.close


%>

<script language='javascript'>


function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);

	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

function showDays(yyyymmdd, days)
{
	var f = document.frmWrite;

	f.yyyymmdd.value = yyyymmdd;
	f.workDays.value = days;

	document.getElementById("today").innerHTML = yyyymmdd;
	document.getElementById("days").style.display = "block";
}

function saveDays()
{
	var f = document.frmWrite;

	if (f.tenUserID.value)
		f.submit();

}

function searchTenUser(tenUserID)
{
	var arr = tenUserID.split(" - ");
	var f = document.frm;

	f.tenUserID.value = arr[1];
	f.submit();

}

function viewDetail(tenUserID,inout)
{
	<% if (use201502 = "") then %>
	alert("콜서버에서 조회가능합니다.");
	return;
	<% end if %>

	var detailpop;
	detailpop = window.open('operatorSummaryReport_pop.asp?tenUserID='+tenUserID+'&inout='+inout+'&sdate=<%=sdate%>&edate=<%=edate%>','detail','width=900,height=672,scrollbars=yes, resizable=yes');
	detailpop.focus();
}

function jsRecvData(frm) {
	if (confirm(" 데이타를 가져오시겠습니까? [" + frm.sDate.value + ", " + frm.tenUserID.value + "]") == true) {
		window.open("operatorSummaryReport_process.asp?mode=recvdata&yyyymmdd=" + frm.sDate.value + "&tenUserID=" + frm.tenUserID.value, 'jsRecvData', 'width=300,height=150,scrollbars=yes, resizable=yes');
	}
}

function jsSetMonth(diff) {
	var frm = document.frm;
	var sDate = frm.sDate;
	var eDate = frm.eDate;
	var newDate;

	var date = new Date(sDate.value);
	if (date.getDate() == 1) {
		newDate = new Date(date.setMonth(date.getMonth() + diff));
		frm.sDate.value = newDate.yyyymmdd();
		date.setMonth(date.getMonth()+1)
		newDate = new Date(date.addDays(-1));
		frm.eDate.value = newDate.yyyymmdd();
	} else {
		newDate = date.addDays(-1 * (date.getDate() - 1));
		frm.sDate.value = newDate.yyyymmdd();
		newDate = new Date(newDate.setMonth(newDate.getMonth()+1));
		newDate = newDate.addDays(-1);
		frm.eDate.value = newDate.yyyymmdd();
	}
}

Date.prototype.yyyymmdd = function() {
  var mm = this.getMonth() + 1; // getMonth() is zero-based
  var dd = this.getDate();

  return [this.getFullYear(),
          (mm>9 ? '' : '0') + mm,
          (dd>9 ? '' : '0') + dd
         ].join('-');
};

Date.prototype.addDays = function(days) {
    var date = new Date(this.valueOf());
    date.setDate(date.getDate() + days);
    return date;
}

</script>

<div id="days" style="display:none; position:absolute; z-index: 1; left: 500px; top: 180px; background-color:#FFFFFF; border:solid 1px #000000; padding:20px; width:200px; " align="center">
	<form name="frmWrite" action="operatorWorkDayProc.asp">
	<input type="hidden" name="sDate" value="<%=sDate%>">
	<input type="hidden" name="eDate" value="<%=eDate%>">
	<input type="hidden" name="tenUserID" value="<%=tenUserID%>">
	<input type="hidden" name="yyyymmdd">
		<span id="today"></span>
		&nbsp;
		근무
		<br><br>
		<select name="workDays">
			<option value="0.0">0</option>
			<option value="0.5">0.5</option>
			<option value="1.0">1</option>
		</select>
		일
		<br><br>
		<input type="button" class="button" value="수정" onclick="saveDays();">
		&nbsp;
		<input type="button" class="button" value="취소" onclick="document.getElementById('days').style.display = 'none';">
	</form>
</div>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
            날짜 : <input type="text" size="10" name="sDate" value="<%=sDate%>" onClick="jsPopCal('frm','sDate');" style="cursor:hand;">
            ~<input type="text" size="10" name="eDate" value="<%=eDate%>" onClick="jsPopCal('frm','eDate');" style="cursor:hand;">
			<input type="button" class="button" value="이전달" onClick="jsSetMonth(-1)">
			<input type="button" class="button" value="다음달" onClick="jsSetMonth(1)">
			&nbsp;
			상담원ID
			<input type="text" class="text" name="tenUserID" value="<%= tenUserID %>">
			&nbsp;
			<input type="checkbox" name="use201502" value="Y" <% if (use201502 = "Y") then %>checked<% end if %> > 2015-03 이전 자료보기
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<div align="right"><input type="button" class="button_s" value="자료 가져오기" onClick="jsRecvData(frm);"></div>

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td rowspan="2" width="100">
		<%If tenUserID <> "" Then %>
			날짜
		<%Else %>
			상담원ID
		<%End If%>
		</td>
		<td rowspan="2" width="40">내선</td>
        <td colspan="3">수신합계</td>
        <td colspan="3">발신합계</td>
        <td colspan="3">합계</td>
		<td rowspan="2">근무일수</td>
		<td rowspan="2">게시판처리수</td>
		<td rowspan="2">일평균통화량</td>
		<td rowspan="2">일평균게시판</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">

        <td>건수</td>
        <td>시간</td>
        <td>평균</td>

        <td>건수</td>
        <td>시간</td>
        <td>평균</td>

        <td>건수</td>
        <td>시간</td>
        <td>평균</td>

</tr>
<%
Dim servicePoint, servicePointText, pnt1, pnt2, pnt3, pnt4
Dim rowCnt
Dim sRs(20)

' 초데이터를 시분초 형식으로 변환
Function sec2time(ByVal sec)
	sec2time = Int(sec / 3600) & ":" & Right("0"&(Int(sec/60) Mod 60),2) & ":" & Right("0"&(sec Mod 60),2)
End Function

If IsArray(rs) Then
	rowCnt = UBound(rs,2) + 1
%>

	<%For i=0 To UBound(rs,2)%>
    <tr align="right" bgcolor="#FFFFFF" height="30">
	<%
		' Row 합산
		sRs(1) = sRs(1) + CDbl(rs(1,i))
		sRs(2) = sRs(2) + CDbl(rs(2,i))
		sRs(3) = sRs(3) + CDbl(rs(3,i))
		sRs(4) = sRs(4) + CDbl(rs(4,i))
		sRs(5) = sRs(5) + CDbl(rs(5,i))
		sRs(6) = sRs(6) + CDbl(rs(6,i))
		sRs(7) = sRs(7) + CDbl(rs(7,i))
		sRs(8) = sRs(8) + CDbl(rs(8,i))
		sRs(9) = sRs(9) + CDbl(rs(9,i))
		sRs(10) = sRs(10) + CDbl(rs(10,i))
		sRs(11) = sRs(11) + CDbl(rs(11,i))
		sRs(12) = sRs(12) + CDbl(rs(12,i))
		sRs(13) = sRs(13) + CDbl(rs(13,i))
		sRs(14) = sRs(14) + CDbl(rs(14,i))

		rs(7,i) = rs(1,i) + rs(4,i)
		rs(8,i) = rs(2,i) + rs(5,i)

		' 평균 항목 재계산
		If CDbl(rs(1,i)) > 0 Then
			rs(3,i) = CLng( CDbl(rs(2,i)) / CDbl(rs(1,i)) )
		End If
		If CDbl(rs(4,i)) > 0 Then
			rs(6,i) = CLng( CDbl(rs(5,i)) / CDbl(rs(4,i)) )
		End If
		If CDbl(rs(7,i)) > 0 Then
			rs(9,i) = CLng( CDbl(rs(8,i)) / CDbl(rs(7,i)) )
		End If

		' 일평균통화량, 일평균게시판처리수
		If CDbl(rs(11,i)) > 0 Then
			rs(12,i) = CLng( CDbl(rs(8,i)) / CDbl(rs(11,i)) )	' 합계시간 / 근무일수
			rs(14,i) = CDbl( CDbl(rs(13,i)) / CDbl(rs(11,i)) )	' 게시판처리수 / 근무일수
		End If
	%>
		<td align="center">
			<%If tenUserID <> "" Then %>
				<%=rs(0,i)%>
			<%Else %>
				<%=rs(16,i)%>
			<%End If%>
		</td>
		<td align="center"><%=rs(17,i)%></td>
		<td><a href="javascript:viewDetail('<%=rs(16,i)%>','in');"><%=FormatNumber(rs(1,i),0)%></a></td>
		<td><%=sec2time(rs(2,i))%></td>
		<td><%=sec2time(rs(3,i))%></td>
		<td><a href="javascript:viewDetail('<%=rs(16,i)%>','out');"><%=FormatNumber(rs(4,i),0)%></a></td>
		<td><%=sec2time(rs(5,i))%></td>
		<td><%=sec2time(rs(6,i))%></td>
		<td><a href="javascript:viewDetail('<%=rs(16,i)%>','all');"><%=FormatNumber(rs(7,i),0)%></a></td>
		<td><%=sec2time(rs(8,i))%></td>
		<td><%=sec2time(rs(9,i))%></td>

		<td>
		<%If tenUserID <> "" Then %>
			<a href="javascript:showDays('<%=rs(0,i)%>','<%=FormatNumber(rs(11,i),1)%>');"><%=FormatNumber(rs(11,i),1)%></a>
		<%Else %>
			<a href="javascript:searchTenUser('<%=rs(0,i)%>');">
				<%=FormatNumber(rs(11,i),1)%>
			</a>
		<%End If%>
		</td>

		<td><%=rs(13,i)%></td>
		<td><%=sec2time(rs(12,i))%></td>
		<td><%=FormatNumber(rs(14,i),1)%></td>
	</tr>
	<%Next%>
    <tr align="right" bgcolor="#FFFFFF">
	<%
		' 합계 평균 항목 재계산
		If CDbl(sRs(1)) > 0 Then
			sRs(3) = CLng( CDbl(sRs(2)) / CDbl(sRs(1)) )
		End If
		If CDbl(sRs(4)) > 0 Then
			sRs(6) = CLng( CDbl(sRs(5)) / CDbl(sRs(4)) )
		End If

		sRs(7) = sRs(1) + sRs(4)
		sRs(8) = sRs(2) + sRs(5)

		If CDbl(sRs(7)) > 0 Then
			sRs(9) = CLng( CDbl(sRs(8)) / CDbl(sRs(7)) )
		End If

		' 일평균통화량, 일평균게시판처리수
		If CDbl(sRs(11)) > 0 Then
			sRs(12) = CLng( CDbl(sRs(8)) / CDbl(sRs(11)) )	' 합계시간 / 근무일수
			sRs(14) = CDbl( CDbl(sRs(13)) / CDbl(sRs(11)) )	' 게시판처리수 / 근무일수
		End If
	%>
    	<td align="center" colspan="2"><b>합계 or 평균</b></td>
		<td><b><%=FormatNumber(sRs(1),0)%></b></td>
		<td><b><%=sec2time(sRs(2))%></b></td>
		<td><b><%=sec2time(sRs(3))%></b></td>
		<td><b><%=FormatNumber(sRs(4),0)%></b></td>
		<td><b><%=sec2time(sRs(5))%></b></td>
		<td><b><%=sec2time(sRs(6))%></b></td>
		<td><b><%=FormatNumber(sRs(7),0)%></b></td>
		<td><b><%=sec2time(sRs(8))%></b></td>
		<td><b><%=sec2time(sRs(9))%></b></td>

		<td><b><%=FormatNumber(sRs(11),1)%></b></td>
		<td><b><%=sRs(13)%></b></td>
		<td><b><%=sec2time(sRs(12))%></b></td>
		<td><b><%=FormatNumber(sRs(14),1)%></b></td>
    </tr>
<%
End If
%>
</table>



<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
