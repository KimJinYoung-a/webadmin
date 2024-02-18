<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/db3Helper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/PagingCls.asp"-->
<!-- #include virtual="/admin/callStatic/libFunction.asp"-->
<%

dim yyyymmdd_from, yyyymmdd_to, extension, calldate, calldate_from, calldate_to, hour_from, hour_to, phoneno, customerphoneno
dim dcontext, disposition, lastappsql, pagesize, currpage, mode
dim i, buf

yyyymmdd_from	= requestCheckVar(trim(request.Form("yyyymmdd_from")),10)
yyyymmdd_to		= requestCheckVar(trim(request.Form("yyyymmdd_to")),10)
extension 		= requestCheckVar(trim(request.Form("extension")),3)
hour_from	 	= requestCheckVar(trim(request.Form("hour_from")),2)
hour_to		 	= requestCheckVar(trim(request.Form("hour_to")),2)
dcontext		= requestCheckVar(trim(request.Form("dcontext")),40)
disposition		= requestCheckVar(trim(request.Form("disposition")),12)
phoneno			= requestCheckVar(trim(request.Form("phoneno")),12)
customerphoneno	= requestCheckVar(trim(request.Form("customerphoneno")),12)
mode			= requestCheckVar(trim(request.Form("mode")),32)

currpage		= requestCheckVar(trim(request.Form("currpage")),8)
pagesize		= 100



if (yyyymmdd_from = "") then
	yyyymmdd_from = Left((Date - 1), 10)					'임시 - 테스트
	yyyymmdd_to = Left((Date - 1), 10)
end if

if (yyyymmdd_from = yyyymmdd_to) then
	if (hour_from <> "") then
		calldate_from 	= yyyymmdd_from & " " & hour_from & ":00:00"
	end if
	if (hour_to <> "") then
		calldate_to 	= yyyymmdd_from & " " & hour_to & ":00:00"
	end if
else
	calldate_from = ""
	calldate_to = ""
end if




if (mode = "") then
	mode = "all"
end if

if (currpage = "") then
	currpage = 1
end if



Dim strSql
Dim rs

Dim paramInfo
paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
	,Array("@PageSize"		, adInteger	, adParamInput	,		, 50)	_
	,Array("@CurrPage"		, adInteger	, adParamInput	,		, currpage) _
	,Array("@yyyymmdd_from"	, adVarchar	, adParamInput	, 10    , yyyymmdd_from) _
	,Array("@yyyymmdd_to"	, adVarchar	, adParamInput	, 10    , yyyymmdd_to) _
	,Array("@extension" 	, adVarchar	, adParamInput	, 3     , extension) _
	,Array("@calldate_from"	, adVarchar	, adParamInput	, 20    , calldate_from) _
	,Array("@calldate_to"	, adVarchar	, adParamInput	, 20    , calldate_to) _
	,Array("@dcontext"  	, adVarchar	, adParamInput	, 40    , dcontext) _
	,Array("@disposition"	, adVarchar	, adParamInput	, 12    , disposition) _
	,Array("@phoneno"   	, adVarchar	, adParamInput	, 12    , phoneno) _
	,Array("@customerphoneno"	, adVarchar	, adParamInput	, 12    , customerphoneno) _
	,Array("@TotalCount"	, adBigInt	, adParamOutput	,		, 0) _
	,Array("@mode"   		, adVarchar	, adParamInput	, 32    , mode) _
)

strSql = "db_datamart.dbo.sp_Ten_Call_Search"

Call db3_fnExecSPReturnRSOutput(strSql, paramInfo)

If Not db3_rsget.EOF Then
	rs = db3_rsget.getRows()
End If
db3_rsget.close




Dim cPaging
Set cPaging = new PagingCls

cPaging.FTotalCount = GetValue(paramInfo, "@TotalCount")
cPaging.FTotalCount = CInt(cPaging.FTotalCount)
cPaging.FPageSize = 50
cPaging.FCurrPage = currpage
cPaging.Calc

'response.write "----------------" & cPaging.FTotalCount

'response.write "----------------" & dcontext & "----------------"

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

function gotoPage(page)
{
	/*
	if (document.frm.mode.selectedIndex > 0) {
		if ((document.frm.dcontext.selectedIndex != 1) || (document.frm.phoneno.selectedIndex != 1)) {
			alert("전화번호가 콜센터헌트이고, 수발신이 수신전화일때만\n\n실패전화 또는 포기전화 리스트를 볼 수 있습니다.");
			return;
		}
	}
	*/

	document.frm.currpage.value = page;
	document.frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="">
	<input type="hidden" name="currpage" value="<%= currpage %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
	       	날짜 : <input type="text" size="10" name="yyyymmdd_from" value="<%=yyyymmdd_from%>" onClick="jsPopCal('frm','yyyymmdd_from');" style="cursor:hand;"> - <input type="text" size="10" name="yyyymmdd_to" value="<%=yyyymmdd_to%>" onClick="jsPopCal('frm','yyyymmdd_to');" style="cursor:hand;"> (오늘의 통화내역은 검색되지 않습니다.)<br>
	       	내선번호 : <% DrawInlinePhoneBox extension %><br>
	       	시간 : <% DrawCallcenterHourBox hour_from, hour_to %> (1일 검색시에만 시간대별 검색이 가능합니다.)<br>
	       	수발신 : <% DrawCallcenterInOutStateBox dcontext %><br>
            전화번호 : <% DrawCallcenterPhoneNameBox phoneno %><br>
            고객번호 : <input type="text" size="12" name="customerphoneno" value="<%= customerphoneno %>"><br>
            상태 : <% DrawCallcenterModeBox mode %>
			&nbsp;
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:gotoPage(1);">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
	       	안내멘트(Playback) : 근무외 시간<br>
	       	통화종료(Hangup) : <br>
	       	통화연결(Dial) :
		</td>
		<td align="left">
	       	대기멘트(BackGround) : 모든 상담원 통화중<br>
	       	내선대기(WaitExten) : 내선 수동연결 대기중<br>
	       	연결대기(Busy) : 내선 연결 후 전화 안받음
		</td>
	</tr>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
	       	콜센터헌트(07075490429)<br>
            사무실헌트(07075490556)
		</td>
		<td align="left">
	       	대표번호1(07075490448)<br>
	       	대표번호2(07075490449)<br>
	       	유아러걸(07075490559,0216440560)
		</td>
	</tr>
</table>
<br>
총 <%= cPaging.FTotalCount %> 건
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">

        <td>no</td>
        <td>날짜</td>
        <td>내선번호</td>
        <td>아이디</td>
        <td>수발신</td>
        <td>발신</td>
        <td>수신</td>
        <td>통화시간</td>
        <td>상태</td>
        <!--
        <td>답변</td>
        <td>최종상태</td>
        -->

</tr>
<%
Dim rowCnt
Dim sRs(20)

'select top 100 yyyymmdd, extension, tenUserID, calldate, src, dst, dcontext, lastapp, duration, disposition, userfield

If IsArray(rs) Then
	rowCnt = UBound(rs,2) + 1
%>

	<%For i=0 To UBound(rs,2)%>
    <tr align="center" bgcolor="#FFFFFF">
	<%
		' Row 합산
		sRs(1) = sRs(1) + 1
		sRs(2) = sRs(2) + CDbl(rs(8,i))

	%>
		<td><%= sRs(1) %></td>
		<td><%= rs(3,i) %></td>
		<td><%= rs(1,i) %></td>
		<td><%= rs(2,i) %></td>
		<td><% PrintCallcenterInOutState rs(6,i) %></td>
		<td><% PrintCallcenterPhoneNumberString rs(4,i) %></td>
		<td><% PrintCallcenterPhoneNumberString rs(5,i) %></td>
		<td><%= SectoTime(rs(8,i)) %></td>
		<td>
<%

if ("inbound"=CStr(rs(6,i))) then
	'==========================================================================
	'수신전화

	if ("Playback"=CStr(rs(7,i))) then
		'======================================================================
		'안내멘트

		buf = "<a href='javascript:alert(""근무외시간 안내멘트중 통화종료"")'>정상(A)</a>"

	elseif ("Hangup"=CStr(rs(7,i))) then
		'======================================================================
		'통화종료

		if ("07075490429"=CStr(rs(5,i))) then
			'==================================================================
			'콜센터헌트

			if (rs(8,i) >= 12) then
				buf = "<a href='javascript:alert(""상담원 통화후 통화종료"")'>정상(B)</a>"
			else
				buf = "<a href='javascript:alert(""연결실패 - 기계오류??"")'>실패(A)</a>"
			end if

		elseif ((""<>CStr(rs(1,i))) and (("07075490448"=CStr(rs(5,i))) or ("07075490556"=CStr(rs(5,i))) or ("07075490557"=CStr(rs(5,i))) or ("07075490558"=CStr(rs(5,i))) or ("07075490449"=CStr(rs(5,i))))) then
			'==================================================================
			'업무전화

			if (rs(8,i) >= 12) then
				buf = "<a href='javascript:alert(""업무전화 통화후 통화종료"")'>정상(C)</a>"
			else
				buf = "<a href='javascript:alert(""연결실패 - 기계오류??"")'>실패(B)</a>"
			end if

		elseif ((""=CStr(rs(1,i))) and (("07075490556"=CStr(rs(5,i))) or ("07075490557"=CStr(rs(5,i))) or ("07075490558"=CStr(rs(5,i))))) then
			'==================================================================
			'업무전화 - 연결실패

			buf = "<a href='javascript:alert(""업무전화 - 연결실패"")'>실패(C)</a>"

		elseif ((""=CStr(rs(1,i))) and ("07075490448"=CStr(rs(5,i))) or ("07075490449"=CStr(rs(5,i)))) then
			'==================================================================
			'근무시간외 고객전화

			if (rs(8,i) >= 12) then
				buf = "<a href='javascript:alert(""안내멘트 청취후 통화종료"")'>정상(D)</a>"
			else
				buf = "<a href='javascript:alert(""연결실패 - 기계오류??"")'>실패(D)</a>"
			end if

		else
			'==================================================================
			'입력 오류

			buf = "<a href='javascript:alert(""입력실패후 통화포기"")'>포기(A)</a>"

		end if

	elseif ("Dial"=CStr(rs(7,i))) then
		'======================================================================
		'통화연결

		if ("07075490429"=CStr(rs(5,i))) then
			'==================================================================
			'콜센터헌트

			if (rs(8,i) >= 12) then
				buf = "<a href='javascript:alert(""상담원 통화후 통화종료"")'>정상(E)</a>"
			else
				buf = "<a href='javascript:alert(""연결실패 - 기계오류??"")'>실패(E)</a>"
			end if

		elseif ((""<>CStr(rs(1,i))) and (("07075490448"=CStr(rs(5,i))) or ("07075490556"=CStr(rs(5,i))) or ("07075490449"=CStr(rs(5,i))) or ("801"=CStr(rs(5,i))) or ("802"=CStr(rs(5,i))) or ("803"=CStr(rs(5,i))) or ("804"=CStr(rs(5,i))) or ("805"=CStr(rs(5,i))) or ("806"=CStr(rs(5,i))) or ("807"=CStr(rs(5,i))))) then
			'==================================================================
			'업무전화

			if (rs(8,i) >= 12) then
				buf = "<a href='javascript:alert(""업무전화 통화후 통화종료"")'>정상(F)</a>"
			else
				buf = "<a href='javascript:alert(""연결실패 - 기계오류??"")'>실패(F)</a>"
			end if

		else
			'==================================================================
			'입력 오류

			buf = "<a href='javascript:alert(""입력실패후 통화포기"")'>포기(B)</a>"

		end if

	elseif ("BackGround"=CStr(rs(7,i))) then
		'======================================================================
		'대기멘트 : 모든 상담원 통화중

		buf = "<a href='javascript:alert(""모든 상담원 통화중 멘트 청취중 통화포기"")'>포기(C)</a>" & CStr(rs(7,i))

	elseif ("WaitExten"=CStr(rs(7,i))) then
		'내선수동연결대기

		buf = "<a href='javascript:alert(""내선 수동연결 대기중 통화종료"")'>정상(G)</a>" & CStr(rs(7,i))

	elseif ("Busy"=CStr(rs(7,i))) then
		'내선연결시도(toexten) 하였으나 전화 안받음

		buf = CStr(rs(7,i))

	else
		'에러
		buf = CStr(rs(7,i))
	end if

elseif ("outbound"=CStr(rs(6,i))) then
	'==========================================================================
	'발신전화
	buf = CStr(rs(7,i))
elseif ("toexten"=CStr(rs(6,i))) then
	'==========================================================================
	'내선전환
	buf = CStr(rs(7,i))
else
	'==========================================================================
	'에러
	buf = CStr(rs(7,i))
end if



'수신전화
if ("inbound"=CStr(rs(6,i))) then

	if ("ResetCDR"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "1") then
		buf = "<span title='콜센터에 통화시도'><font color=red>콜 시도</font></span>"
	end if

	if ("Playback"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490449") then
		buf = "<span title='근무외시간 안내멘트중 통화종료'><font color=gray>안내멘트1</font></span>"
	end if

	if ("Hangup"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490449") then
		buf = "<span title='근무외시간 안내멘트 청취 후 통화종료'><font color=gray>안내멘트2</font></span>"
	end if

	if ("BackGround"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490449") then
		buf = "<span title='콜센터에 통화시도하였으나 1 번 미입력 또는 잘못 입력후 중단'><font color=gray>시도중단1</font></span>"
	end if

	if ("WaitExten"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490449") then
		buf = "<span title='콜센터에 통화시도하였으나 1 번 미입력 또는 잘못 입력후 중단'><font color=gray>시도중단2</font></span>"
	end if

	if ("ResetCDR"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "0") then
		buf = "<span title='사무실과 연결시도'><font color=black>전화시도</font></span>"
	end if

	if ("BackGround"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490448") then
		buf = "<span title='사무실과 연결시도 중 내선번호 미입력 상태로 통화중단'><font color=gray>시도중단3</font></span>"
	end if

	if ("WaitExten"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490448") then
		buf = "<span title='사무실과 연결시도 중 내선번호 미입력 또는 잘못입력 상태로 통화중단'><font color=gray>시도중단4</font></span>"
	end if

	if ("BackGround"=CStr(rs(7,i))) and (CStr(rs(5,i)) <> "07075490448") and (CStr(rs(5,i)) <> "07075490449") and (CStr(rs(5,i)) <> "1") then
		buf = "<span title='콜센터에 통화시도하였으나 1 번 미입력 또는 잘못 입력후 중단'><font color=gray>시도중단5</font></span>"
	end if

end if



'모지모지
if ("hunt_context"=CStr(rs(6,i))) then

	if ("Queue"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490429") then
		buf = "<span title='콜센터와 통화성공'><font color=green>콜 성공</font></span>"
	end if

	if ("Playback"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490556") then
		buf = "<span title='사무실 내선번호 입력 후 전화 안받음'><font color=gray>통화실패</font></span>"
	end if

	if ("Playback"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490559") then
		buf = "<span title='유아러걸 콜센터 전환'><font color=gray>유아러걸</font></span>"
	end if

	'if ("Queue"=CStr(rs(7,i))) and (CStr(rs(5,i)) = "07075490556") then
	'	buf = "<span title='내선전환 후 사무실과 연결성공'><font color=black>전화성공</font></span>"
	'end if

	'if ("Dial"=CStr(rs(7,i))) then
	'	''''''''''''''''buf = "<span title='내선전환 후 사무실과 통화성공'><font color=gray>통화성공</font></span>"
	'end if

end if



'모지모지(직접 내부 전화번호 연결??)
if ("pers_context"=CStr(rs(6,i))) then

	if ("Dial"=CStr(rs(7,i))) then
		if (rs(1,i) = rs(5,i)) then
			buf = "<span title='개인 전화번호 연결 시도 후 통화성공'><font color=gray>통화성공</font></span>"
		else
			buf = "<span title='개인 전화번호 연결 시도 후 통화성공(땡겨받음:" & rs(1,i) & " > " & rs(5,i) & ")'><font color=gray>통화성공</font></span>"
		end if
	end if

	if ("Hangup"=CStr(rs(7,i))) and (rs(5,i) = "908") and (rs(8,i) = 0) then
		buf = "<span title='개인 전화번호 연결 시도 하였으나 통화실패'><font color=red>통화실패</font></span>"
	end if

	if ("Hangup"=CStr(rs(7,i))) and (rs(8,i) > 0) then
		buf = "<span title='개인 전화번호 연결 시도 통화성공'><font color=red>통화성공</font></span>"
	end if

end if



'발신전화
if ("outbound"=CStr(rs(6,i))) then

	if (CStr(rs(4,i)) = "0216446030") then
		buf = "<span title='콜센터에서 외부에 전화 시도'><font color=black>콜 전화</font></span>"
	end if

	if (CStr(rs(4,i)) = "0216440560") then
		buf = "<span title='유아러걸 콜센터에서 외부에 전화 시도'><font color=gray>유아러걸</font></span>"
	end if

	if (CStr(rs(4,i)) = "0216441851") then
		buf = "<span title='물류센터에서 외부에 전화 시도'><font color=gray>물류 전화</font></span>"
	end if



end if

response.write buf

%>
		</td>
		<!--
		<td><%= rs(9,i) %></td>
		<td><% PrintCallcenterLastState rs(7,i) %></td>
		-->
	</tr>
	<%Next%>
<!--
    <tr align="center" bgcolor="#FFFFFF">
    	<td><b>합계</b></td>
    	<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td><<b><%=FormatNumber(sRs(2),0)%></b></td>
		<td></td>-->
		<!--
		<td></td>
		<td></td>
		-->
    <!--</tr>-->
<%
End If
%>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
	   	<% if cPaging.HasPrevScroll then %>
			<span class="list_link"><a href="javascript:gotoPage(<%= cPaging.StartScrollPage-1 %>)">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + cPaging.StartScrollPage to cPaging.StartScrollPage + cPaging.FScrollCount - 1 %>
			<% if (i > cPaging.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(cPaging.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:gotoPage(<%= i %>)" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if cPaging.HasNextScroll then %>
			<span class="list_link"><a href="javascript:gotoPage(<%= i %>)">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
