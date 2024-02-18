<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member/tenAgitCalendarCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
	dim part_sn, userid, nowYear, nowMonth, AreaDiv, strRst
	dim department_id, aDepNm, sDepNm
  	dim nowDateShort, availSdate, availEdate, nowYearShort

	'// 신청 가능 기간 지정
	nowDateShort = mid(date(),6,10)
	nowYearShort = year(date)
'	nowDateShort = "12-18"
'	nowYearShort = 2020
	if nowDateShort>="02-15" and nowDateShort<="04-14" then		'1차 신청
		availSdate = dateserial(nowYearShort,2,17)
		availEdate = dateserial(nowYearShort,4,30)
	elseif nowDateShort>="04-15" and nowDateShort<="06-14" then	'2차 신청
		availSdate = dateserial(nowYearShort,4,17)
		availEdate = dateserial(nowYearShort,6,30)
	elseif nowDateShort>="06-15" and nowDateShort<="08-14" then	'3차 신청
		availSdate = dateserial(nowYearShort,6,15)
		availEdate = dateserial(nowYearShort,8,31)
	elseif nowDateShort>="08-15" and nowDateShort<="10-14" then	'4차 신청
		availSdate = dateserial(nowYearShort,8,17)
		availEdate = dateserial(nowYearShort,10,31)
	elseif nowDateShort>="10-15" and nowDateShort<="12-14" then	'5차 신청
		availSdate = dateserial(nowYearShort,10,17)
		availEdate = dateserial(nowYearShort,12,31)
	elseif nowDateShort>="12-15" or nowDateShort<="02-14" then	'6차 신청(익년)
		availSdate = dateserial(chkIIF(nowDateShort<="02-14",nowYearShort-1,nowYearShort),12,12)
		availEdate = dateserial(chkIIF(nowDateShort<="02-14",nowYearShort,nowYearShort+1),2,28)
	end if

	part_sn		= requestCheckvar(Request("part_sn"),4)
	userid		= requestCheckvar(Request("userid"),32)
	nowYear		= requestCheckvar(Request("nowYear"),4)
	nowMonth	= requestCheckvar(Request("nowMonth"),2)
	department_id = requestCheckvar(Request("department_id"),10)
	AreaDiv		= requestCheckvar(Request("AreaDiv"),6)			'1:제주, 2:양평, 3:속초
	
	if nowYear="" then nowYear=year(date)
	if nowMonth="" then nowMonth=month(date)

	'// 달력 접수
	dim oAgitCal, lp, weekno
	Set oAgitCal = new CAgitCalendar

	oAgitCal.FRectYear = nowYear
	oAgitCal.FRectMonth = Num2Str(nowMonth,2,"0","R")
	oAgitCal.CalendarList

	'// 예약 내용 접수
	dim oAgitBook, bx
	Set oAgitBook = new CAgitCalendar

	oAgitBook.FRectArea = AreaDiv
	oAgitBook.FRectYear = nowYear
	oAgitBook.FRectMonth = Num2Str(nowMonth,2,"0","R")
	oAgitBook.FRectpart_sn = part_sn
	oAgitBook.FRectdepartment_id  = department_id
	oAgitBook.FRectUserid = userid
	oAgitBook.BookingList
%>
<style type="text/css">
<!--
.calTT { font-family:malgun gothic; color:#606060; }
.calNoB { font-family:Arial; color:#000; font-weight:bold; }
.calNoR { font-family:Arial; color:#FF7065; font-weight:bold; }
.calHoly { font-family:malgun gothic; color:#FFFFFF; background-color:#AABBFF; font-size:11px; padding-left:5px; border-radius: 7px; -moz-border-radius: 7px; -webkit-border-radius: 7px; margin:2px;}
.calHolyB { font-family:malgun gothic; color:#FFFFFF; background-color:#BABFCF; font-size:11px; padding-left:5px; border-radius: 7px; -moz-border-radius: 7px; -webkit-border-radius: 7px; margin:2px;}
.calHolyO { font-family:malgun gothic; color:#FFFFFF; background-color:#F2CB61; font-size:11px; padding-left:5px; border-radius: 7px; -moz-border-radius: 7px; -webkit-border-radius: 7px; margin:2px;}
.calToday { font-family:malgun gothic; color:#000; background-color:#F2F6FF;}
 
.calItem {font-family:malgun gothic; color:#000; font-size:11px; padding:5px; border-radius: 7px; -moz-border-radius: 7px; -webkit-border-radius: 7px; margin:2px;}
.calItem em {float:right; width:14px; height:14px; border:2px solid #FFF; border-radius: 50%; font-size:10px; line-height:1.3; font-weight:bold; font-style: normal; text-align:center;}

.infoBdg {float:right; width:14px; height:14px; border:2px solid #FFF; border-radius: 50%; font-size:10px; line-height:1.3; font-weight:bold; font-style: normal; text-align:center;}

.calFill1 {background-color:#F6FFE8;}
.calFill2 {background-color:#E8FFF6;}
.calFill3 {background-color:#E8F6FF;}
.calFill4 {background-color:#FFF6E8;}
.calFill5 {background-color:#F6E8FF;}
.calFill6 {background-color:#FFE8F6;}

.calNull { background-color:#F0F0F0 }

.btnCircle {width: 26px; height:26px; font-size: 10px; border-radius: 50%; padding:3px; vertical-align:5px;}
/*
.tabArea {display:table;}
.tabArea .tabRow {display:table-row;}
.tabArea .btnTabTop {display:table-cell; border-bottom: 10px solid #DADADA; border-right: 10px solid transparent; height: 0; width: 100px;}
.tabArea .btnTabBody {display:table-cell; background: #E0E0E0; height: 18px; width: 100px; text-align:center; cursor: pointer;}
*/
.tabArea .btnTabTop {border-bottom: 10px solid #DADADA; border-right: 10px solid transparent; height: 0; width: 100px; font-size:1px; line-height:1;}
.tabArea .btnTabBody {background: #E0E0E0; height: 18px; width: 100px; text-align:center; cursor: pointer;}

.tabArea .currentTop {border-bottom: 10px solid #F88A8A;}
.tabArea .currentBody {background: #F89898;}
.tabArea .currentTop1 {border-bottom: 10px solid #B0FFB0;}
.tabArea .currentBody1 {background: #C8FFC8;}
.tabArea .currentTop2 {border-bottom: 10px solid #FFBABA;}
.tabArea .currentBody2 {background: #FFC8C8;}
.tabArea .currentTop3 {border-bottom: 10px solid #BABAFF;}
.tabArea .currentBody3 {background: #C8C8FF;}

.bgBdgGr {background: #473; color:#FFF;}
.bgBdgRd {background: #743; color:#FFF;}
.bgBdgBl {background: #437; color:#FFF;}

-->
</style>
<script type="text/javascript">
<!--
	function goPage(yyyy,mm) {
		frm.nowYear.value=yyyy;
		frm.nowMonth.value=mm;
		frm.submit();
	}

	function newBook(sDate) {
		if (sDate < "<%=date()%>"){
			alert("오늘 이전 날짜는 예약하실 수 없습니다.");
			return;
		}
		
		if(sDate<"<%=availSdate%>" || sDate>"<%=availEdate%>"){
			alert("아직 오픈되지 않은 날짜입니다. 신청 가능한 기간을 확인해주세요.");
			return;
		}
		
		var p = window.open("pop_tenbyten_Agit_Edit.asp?ChkStart="+sDate,"popNAgit","width=700,height=700,scrollbars=yes,resize=yes");
		p.focus();
	}

	function modiBook(idx) {
		var p = window.open("pop_tenbyten_Agit_Edit.asp?idx="+idx,"popMAgit","width=700,height=700,scrollbars=yes,resize=yes");
		p.focus();
	}

	function fnSelAgit(ArDv) {
		document.frm.AreaDiv.value=ArDv;
		document.frm.submit();
	}

	function fnGoToday() {
		frm.nowYear.value="<%=Year(date)%>";
		frm.nowMonth.value="<%=Num2str(Month(date),2,"0","R")%>";
		frm.submit();
	}
	
	function setHoliday(){
		var h = window.open("pop_Calendar_regHoliday.asp","popCal","width=700,height=400,scrollbars=yes,resize=yes");
		h.focus();
	}
//-->
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="#EEEEEE">검색<br>조건</td>
		<td align="left">
			아지트 : 
			<select name="AreaDiv" class="select">
				<option value="">전체</option>
				<!--<option value="1">제주도</option>-->
				<!--<option value="2">양평</option>-->
				<option value="3">속초</option>
			</select>
			<script type="text/javascript">
				document.frm.AreaDiv.value="<%=AreaDiv%>";
			</script> /
    		일자 : <% call DrawYMSelBox("nowYear","nowMonth",nowYear,nowMonth) %>
    		<input type="button" value="이번달" class="button" onclick="fnGoToday();" /><br>
    		부서 : <%= drawSelectBoxDepartmentALL("department_id", department_id) %> /
    		직원ID : <input type="text" class="text" name="userid" value="<%=userid%>" size="12" maxlength="32">
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
	    <input type="button" value="◀" onclick="goPage('<%=nowYear-1%>','<%=nowMonth%>')" class="btnCircle" style="padding-right:3px;">
	    <b><%=nowYear%></b>
	    <input type="button" value="▶" onclick="goPage('<%=nowYear+1%>','<%=nowMonth%>')" class="btnCircle">
	    &nbsp;/&nbsp;
	    <input type="button" value="◀" onclick="goPage('<%=chkIIF(nowMonth-1<1,nowYear-1,nowYear)%>','<%=chkIIF(nowMonth-1<1,"12",nowMonth-1)%>')" class="btnCircle" style="padding-right:3px;">
	    <b><%=nowMonth%></b>
	    <input type="button" value="▶" onclick="goPage('<%=chkIIF(nowMonth+1>12,nowYear+1,nowYear)%>','<%=chkIIF(nowMonth+1>12,"1",nowMonth+1)%>')" class="btnCircle">
	</td>
	<td width="200" align="right">
	<%	if date()>=availSdate and date()<=availEdate then %>
		<input type="button" class="button" value="신규등록" onClick="newBook('<%=date()%>')">
	<%	end if %>
	<%
		'// 인사에서만 등록가능
	 	if getEditAble then
	%> 
		<input type="button" class="button" value="휴일등록" onClick="setHoliday()">
	<%  else Response.Write "&nbsp;": end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 예약달력 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10px">
<tr>
	<td colspan="5">
		<table cellpadding="0" cellspacing="0" class="tabArea">
		<tr>
			<td class="btnTabTop <%=chkIIF(AreaDiv="","currentTop","")%>">&nbsp;</td>
			<!--<td class="btnTabTop <%=chkIIF(AreaDiv="1","currentTop1","")%>">&nbsp;</td>-->
			<!--<td class="btnTabTop <%=chkIIF(AreaDiv="2","currentTop2","")%>">&nbsp;</td>-->
			<td class="btnTabTop <%=chkIIF(AreaDiv="3","currentTop3","")%>">&nbsp;</td>
		</tr>
		<tr>
			<td class="btnTabBody <%=chkIIF(AreaDiv="","currentBody","")%>" onclick="fnSelAgit('')">전체</td>
			<!--<td class="btnTabBody <%=chkIIF(AreaDiv="1","currentBody1","")%>" onclick="fnSelAgit('1')">제주도</td>-->
			<!--<td class="btnTabBody <%=chkIIF(AreaDiv="2","currentBody2","")%>" onclick="fnSelAgit('2')">양평</td>-->
			<td class="btnTabBody <%=chkIIF(AreaDiv="3","currentBody3","")%>" onclick="fnSelAgit('3')">속초</td>
		</tr>
		</table>
	</td>
	<td colspan="2" align="right" style="color:#777;">
		<div class="infoBdg bgBdgBl">S</div>
		<div style="float:right; margin-top:4px; padding:0 5px 0 10px;">속초:</div>
		<!--
		<div class="infoBdg bgBdgRd">Y</div>
		<div style="float:right; margin-top:4px; padding:0 5px 0 10px;">양평:</div>
		<div class="infoBdg bgBdgGr">J</div>
		<div style="float:right; margin-top:4px; padding-right:5px;">제주:</div>
		-->
	</td>
</tr>
<tr height="30" align="center" bgcolor="#FFFFFF">
	<td width="14%" class="calTT">일</td>
	<td width="14%" class="calTT">월</td>
	<td width="14%" class="calTT">화</td>
	<td width="14%" class="calTT">수</td>
	<td width="14%" class="calTT">목</td>
	<td width="14%" class="calTT">금</td>
	<td width="14%" class="calTT">토</td>
</tr>
</table>
<% if oAgitCal.FResultCount>0 then %>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#D0D0D0" style="padding:5px 0 5px 0;">
<tr height="120" align="center" valign="top" bgcolor="#FFFFFF">
<%
	'// 해당월 1일의 요일
	weekno = DatePart("w", DateSerial(nowYear,nowMonth,"01"))

	'// 달력시작 빈칸 표시
	if weekno>1 then
		for lp=1 to (weekno-1)
			Response.Write "<td width='14%' class='calNull'>&nbsp;</td>"
		next
	end if

	for lp=0 to (oAgitCal.FResultCount-1)
		weekno = DatePart("w", oAgitCal.FItemList(lp).FDate)
	 
		
%>
	<td width="14%" <%=chkIIF(oAgitCal.FItemList(lp).FDate=cstr(date),"class='calToday'","")%>>
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td style="font-size:8px; color:#DDD;">
			<%
				if datediff("d",availSdate,oAgitCal.FItemList(lp).FDate)>=0 and datediff("d",oAgitCal.FItemList(lp).FDate,availEdate)>=0 then
					Response.Write "+"
				end if
			%>
			</td>
			<td align="right" class="<%=chkIIF(weekno=1 or oAgitCal.FItemList(lp).Fholiday>1,"calNoR","calNoB")%>"  onClick="newBook('<%=oAgitCal.FItemList(lp).FDate%>');" style='cursor:pointer'><%=lp+1%></td>
		</tr>
		<tr>
			<td align="left" colspan="2">
				<% if Not(oAgitCal.FItemList(lp).Fholiday_name="" or isNull(oAgitCal.FItemList(lp).Fholiday_name)) then %><div class="calHoly<%=chkIIF(oAgitCal.FItemList(lp).Fholiday>1,"","B")%>"><%=oAgitCal.FItemList(lp).Fholiday_name%></div><% end if %>
				<% ''if fnPeakSeason(oAgitCal.FItemList(lp).FDate) then: Response.Write "<div class=""calHolyO"">성수기</div>": end if %>
				<%
					'// 예약 내용 표시
					if oAgitBook.FResultCount>0 then
					For bx=0 to oAgitBook.FResultCount-1
						if oAgitCal.FItemList(lp).FDate>=Left(oAgitBook.FItemList(bx).FChkStart,10) and oAgitCal.FItemList(lp).FDate<=Left(oAgitBook.FItemList(bx).FChkEnd,10) then
							
							Select Case oAgitBook.FItemList(bx).FareaDiv
							Case "1"
							'// 제주 아지트
								if getEditAble or  oAgitBook.FItemList(bx).Fempno =session("ssBctSn") or  oAgitBook.FItemList(bx).Fuserid =session("ssBctId") then
									strRst = "<div class='calItem calFill" & (oAgitBook.FItemList(bx).Fidx mod 3)+1 & "' onClick='modiBook(" & oAgitBook.FItemList(bx).Fidx & ")' style='cursor:pointer'><em class=""bgBdgGr"">J</em>"
								else
									strRst = "<div class='calItem calFill" & (oAgitBook.FItemList(bx).Fidx mod 3)+1 & "'><em class=""bgBdgGr"">J</em>"
								end if
							Case "2"
							'// 양평 아지트
								if getEditAble or  oAgitBook.FItemList(bx).Fempno =session("ssBctSn") or  oAgitBook.FItemList(bx).Fuserid =session("ssBctId") then
									strRst = "<div class='calItem calFill" & (oAgitBook.FItemList(bx).Fidx mod 3)+4 & "' onClick='modiBook(" & oAgitBook.FItemList(bx).Fidx & ")' style='cursor:pointer'><em class=""bgBdgRd"">Y</em>"
								else
									strRst = "<div class='calItem calFill" & (oAgitBook.FItemList(bx).Fidx mod 3)+4 & "'><em class=""bgBdgRd"">Y</em>"
								end if
							Case "3"
							'// 속초 아지트
								if getEditAble or  oAgitBook.FItemList(bx).Fempno =session("ssBctSn") or  oAgitBook.FItemList(bx).Fuserid =session("ssBctId") then
									strRst = "<div class='calItem calFill" & (oAgitBook.FItemList(bx).Fidx mod 3)+4 & "' onClick='modiBook(" & oAgitBook.FItemList(bx).Fidx & ")' style='cursor:pointer'><em class=""bgBdgBl"">S</em>"
								else
									strRst = "<div class='calItem calFill" & (oAgitBook.FItemList(bx).Fidx mod 3)+4 & "'><em class=""bgBdgBl"">S</em>"
								end if
							End Select
							
				
							'if left(oAgitBook.FItemList(bx).FChkStart,10)=oAgitCal.FItemList(lp).FDate then strRst = strRst & "<b>[체크인 " & hour(oAgitBook.FItemList(bx).FChkStart) & "시]</b><br>"
							'if left(oAgitBook.FItemList(bx).FChkEnd,10)=oAgitCal.FItemList(lp).FDate then strRst = strRst & "<b>[체크아웃 " & hour(oAgitBook.FItemList(bx).FChkEnd) & "시]</b><br>"
							
							'짧은 부서명 쓰기
							aDepNm = split(oAgitBook.FItemList(bx).Fdepartmentnamefull,"-")
							if ubound(aDepNm)>0 then
								sDepNm = aDepNm(ubound(aDepNm)) & " : "
							else 
								sDepNm = ""
							end if
							
							strRst = strRst & sDepNm & oAgitBook.FItemList(bx).FuserName & "<br>"
							strRst = strRst & "인원 : " & oAgitBook.FItemList(bx).FusePersonNo & "명<br>"
							''strRst = strRst & "(" & oAgitBook.FItemList(bx).FuserHP & ")<br>"							''개인정보 삭제
							strRst = strRst & "</div>"
							Response.Write strRst
					 
						end if
					Next
					end if
				%>
			</td>
		</tr>
		</table>
	</td>
<%
		'행구분
		if weekno=7 and day(dateAdd("d",1,oAgitCal.FItemList(lp).FDate))>1 then Response.Write "</tr><tr height='120' align='center' valign='top' bgcolor='#FFFFFF'>"
	next

	'// 달력끝 여백 표시
	if weekno<7 then
		for lp=(weekno+1) to 7
			Response.Write "<td width='14%' class='calNull'>&nbsp;</td>"
		next
	end if
%>
</tr>
</table>
<% end if %>
<!-- 예약달력 끝 -->
<%
	Set oAgitCal = Nothing
	Set oAgitBook = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->