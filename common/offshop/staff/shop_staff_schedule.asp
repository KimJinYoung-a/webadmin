<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  오프라인 매장근무관리
' History : 2011.03.17 한용민 생성
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/staff/staff_cls.asp"-->
<%
dim part_sn, nowYear, nowMonth , loginidshopormaker ,showshopselect
dim dispwrite , empno , userid , username , schedulearr , i ,ostaff, lp, weekno
	part_sn		= Request("part_sn")
	empno		= Request("empno")
	nowYear		= Request("nowYear")
	nowMonth	= Request("nowMonth")
	userid		= Request("userid")
	username		= Request("username")
	
	if nowYear="" then nowYear=year(date)
	if nowMonth="" then nowMonth=month(date)
	
	part_sn = "18"
	dispwrite = false
	showshopselect = false
	loginidshopormaker = ""

if (C_IS_SHOP) then
	'직영/가맹점
	loginidshopormaker = C_STREETSHOPID
else
	if (C_IS_Maker_Upche) then
		loginidshopormaker = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
			loginidshopormaker = "--"		'표시안한다. 에러.
		else
			showshopselect = true
			loginidshopormaker = request("shopid")
		end if
	end if
end if

if loginidshopormaker <> "" then
	
	'//본사 직원 일경우	
	if C_ADMIN_USER then
		dispwrite = true
	
	'//매장 점장일 경우
	elseif getjob_sn("",session("ssBctID"))	= "6" then
		dispwrite = true
	end if
	
	dispwrite = true
end if

'// 달력 접수
Set ostaff = new CAgitCalendar		
	ostaff.FRectYear = nowYear
	ostaff.FRectMonth = Num2Str(nowMonth,2,"0","R")	
	ostaff.CalendarList
		
	if loginidshopormaker = "" then
		response.write "<script language='javascript'>"
		response.write "	alert('매장을 선택해주세요');"
		response.write "</script>"				
	end if

'//디비 서버 부하를 줄이기 위해 직원 스케줄 내역 한큐에 배열로 받아옴
schedulearr = fnPrintBookingCont(nowYear&"-"&Num2Str(nowMonth,2,"0","R"),part_sn,empno,loginidshopormaker,userid,username)
%>

<style type="text/css">
	.calTT { font-family:malgun gothic; color:#606060; }
	.calNoB { font-family:Arial; color:#000; font-weight:bold; }
	.calNoR { font-family:Arial; color:#FF7065; font-weight:bold; }
	.calHoly { font-family:malgun gothic; color:#FFFFFF; width:100%; background-color:#AABBFF; font-size:11px; line-height:14px; padding-left:5px; border-radius: 7px; -moz-border-radius: 7px; -webkit-border-radius: 7px; margin:2px;}
	.calToday { font-family:malgun gothic; color:#000; background-color:#F2F6FF;}
	.calFill1 { font-family:malgun gothic; color:#000; background-color:#FFF6F2; font-size:11px; line-height:13px; width:100%; padding:5px; border-radius: 7px; -moz-border-radius: 7px; -webkit-border-radius: 7px; margin:2px;}
	.calFill2 { font-family:malgun gothic; color:#000; background-color:#F6FFF2; font-size:11px; line-height:13px; width:100%; padding:5px; border-radius: 7px; -moz-border-radius: 7px; -webkit-border-radius: 7px; margin:2px;}
	.calFill3 { font-family:malgun gothic; color:#000; background-color:#F6F2FF; font-size:11px; line-height:13px; width:100%; padding:5px; border-radius: 7px; -moz-border-radius: 7px; -webkit-border-radius: 7px; margin:2px;}
	.calNull { background-color:#F0F0F0 }
</style>

<script language="javascript">

	function goPage(yyyy,mm) {
		frm.nowYear.value=yyyy;
		frm.nowMonth.value=mm;
		frm.submit();
	}

	function modiBook(idx,shopid) {
		var modiBook = window.open("/common/offshop/staff/shop_staff_schedule_Edit.asp?idx="+idx+"&shopid="+shopid,"modiBook","width=700,height=400,scrollbars=yes,resize=yes");
		modiBook.focus();
	}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="#EEEEEE">검색<br>조건</td>
	<td align="left">
		매장 : 
		<% if (showshopselect = true) then %>
			<% 'drawSelectBoxOffShop "shopid",loginidshopormaker %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",loginidshopormaker, "21") %>
		<% else %>
			<%= loginidshopormaker %><input type="hidden" name="shopid" value="<%= loginidshopormaker %>">
		<% end if %>			
		년/월 : <% call DrawYMSelBox("nowYear","nowMonth",nowYear,nowMonth) %>		
		부서 : <%=printPartOption("part_sn", part_sn)%>		
		<br>사원번호 : <input type="text" class="text" name="empno" value="<%=empno%>" size="16" maxlength="16">
		아이디 : <input type="text" class="text" name="userid" value="<%=userid%>" size="16" maxlength="16">
		이름 : <input type="text" class="text" name="username" value="<%=username%>" size="16" maxlength="16">
	</td>
	<td width="50" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 5px 0;">
<tr>
	<td width="40">&nbsp;</td>
	<td align="center" style="font-size:26px; font-family:malgun gothic;">
	    <input type="button" value="◀" onclick="goPage('<%=nowYear-1%>','<%=nowMonth%>')">
	    <b><%=nowYear%></b>
	    <input type="button" value="▶" onclick="goPage('<%=nowYear+1%>','<%=nowMonth%>')">
	    &nbsp;/&nbsp;
	    <input type="button" value="◀" onclick="goPage('<%=chkIIF(nowMonth-1<1,nowYear-1,nowYear)%>','<%=chkIIF(nowMonth-1<1,"12",nowMonth-1)%>')">
	    <b><%=nowMonth%></b>
	    <input type="button" value="▶" onclick="goPage('<%=chkIIF(nowMonth+1>12,nowYear+1,nowYear)%>','<%=chkIIF(nowMonth+1>12,"1",nowMonth+1)%>')">
	</td>
	<td width="40" align="right">
		<% if dispwrite then %>
			<input type="button" class="button" value="신규등록" onClick="modiBook('','<%= loginidshopormaker %>')">
		<% end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 예약달력 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#D0D0D0" style="padding:5px 0 5px 0;">
<tr height="30" align="center" bgcolor="#FFFFFF">
	<td width="14%" class="calTT">일</td>
	<td width="14%" class="calTT">월</td>
	<td width="14%" class="calTT">화</td>
	<td width="14%" class="calTT">수</td>
	<td width="14%" class="calTT">목</td>
	<td width="14%" class="calTT">금</td>
	<td width="14%" class="calTT">토</td>
</tr>
<% if ostaff.FResultCount>0 then %>

<tr align="center" height="180" valign="top" bgcolor="#FFFFFF">
<%
	'// 해당월 1일의 요일
	weekno = DatePart("w", DateSerial(nowYear,nowMonth,"01"))

	'// 달력시작 빈칸 표시
	if weekno>1 then
		for lp=1 to (weekno-1)
			Response.Write "<td class='calNull'>&nbsp;</td>"
		next
	end if

	for lp=0 to (ostaff.FResultCount-1)
	
	weekno = DatePart("w", ostaff.FItemList(lp).FDate)
%>
	<td <%=chkIIF(ostaff.FItemList(lp).FDate=cstr(date),"class='calToday'","")%>>
		<table width="98%" cellpadding="0" cellspacing="0" class="a">
		<tr><td align="right" class="<%=chkIIF(weekno=1 or ostaff.FItemList(lp).Fholiday>1,"calNoR","calNoB")%>"><%=lp+1%></td></tr>
		<tr>
			<td>
				<% if Not(ostaff.FItemList(lp).Fholiday_name="" or isNull(ostaff.FItemList(lp).Fholiday_name)) then %><div class="calHoly"><%=ostaff.FItemList(lp).Fholiday_name%></div><% end if %>
				<%
				'//직원 스케줄 내역 출력
				if isarray(schedulearr) then
					
				for i = 0 to ubound(schedulearr,2)
				
				if ostaff.FItemList(lp).FDate >= left(schedulearr(1,i),10) and ostaff.FItemList(lp).FDate <= left(schedulearr(2,i),10) then
				%>
				<div class='calFill<%=(schedulearr(0,i) mod 3)+1 %>' <% if dispwrite then %>onClick='modiBook(<%= schedulearr(0,i) %>)'<% end if %> style='cursor:pointer'>					
					
					<% if left(schedulearr(1,i),10)= ostaff.FItemList(lp).FDate then %>
						<b>[휴무시작 <%= hour(schedulearr(1,i)) %>시]</b>
					<% end if %>
					
					<% if left(schedulearr(2,i),10)= ostaff.FItemList(lp).FDate then %>
						<b>[휴무끝 <%= hour(schedulearr(2,i)) %>시]</b>
					<% end if %>
					
					<br><%= schedulearr(3,i) %>			
				</div>
				<%
				end if
				
				next
				
				end if
				%>
			</td>
		</tr>
		</table>
	</td>
<%
		'행구분
	if weekno=7 and day(dateAdd("d",1,ostaff.FItemList(lp).FDate))>1 then Response.Write "</tr><tr align='center' height='180' valign='top' bgcolor='#FFFFFF'>"
	
	next

	'// 달력끝 여백 표시
	if weekno < 7 then
		for lp=(weekno+1) to 7
			Response.Write "<td class='calNull'>&nbsp;</td>"
		next
	end if
%>
</tr>
</table>
<% end if %>
<!-- 예약달력 끝 -->

<%
	Set ostaff = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->