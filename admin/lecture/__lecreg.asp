<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/lecturecls.asp"-->
<%

dim olec
dim idx,mode
dim linkitemid,lectitle,lecturer,lecsum,matinclude,matsum
dim leccount,lectime,tottime,matdesc,properperson,minperson
dim reservestart,reserveend,lecdate01,lecdate02,lecdate03
dim lecdate04,lecdate05,lecdate06,lecdate07,lecdate08,lecdate01_end
dim lecdate02_end,lecdate03_end,lecdate04_end,lecdate05_end,lecdate06_end
dim lecdate07_end,lecdate08_end,leccontents,lecetc
dim lecturerid,lecperiod,leccurry,lecspace
dim yyyymm, regfinish
dim isusing

idx = request("idx")
mode = request("mode")

if idx="" then idx=0

lecturerid = request("lecturerid")
lecperiod = request("lecperiod")
leccurry = request("leccurry")
linkitemid = request("linkitemid")
lectitle = request("lectitle")
lecturer = request("lecturer")
lecsum = request("lecsum")
matinclude = request("matinclude")
matsum = request("matsum")
lecspace = request("lecspace")
leccount = request("leccount")
lectime = request("lectime")
tottime = request("tottime")
matdesc = request("matdesc")
properperson = request("properperson")
minperson = request("minperson")
reservestart = request("reservestart")
reserveend = request("reserveend")
lecdate01 = request("lecdate01")
lecdate02 = request("lecdate02")
lecdate03 = request("lecdate03")
lecdate04 = request("lecdate04")
lecdate05 = request("lecdate05")
lecdate06 = request("lecdate06")
lecdate07 = request("lecdate07")
lecdate08 = request("lecdate08")
lecdate01_end = request("lecdate01_end")
lecdate02_end = request("lecdate02_end")
lecdate03_end = request("lecdate03_end")
lecdate04_end = request("lecdate04_end")
lecdate05_end = request("lecdate05_end")
lecdate06_end = request("lecdate06_end")
lecdate07_end = request("lecdate07_end")
lecdate08_end = request("lecdate08_end")
leccontents = request("leccontents")
lecetc = request("lecetc")
yyyymm = request("yyyymm")
regfinish = request("regfinish")
isusing = request("isusing")

set olec = new CLectureDetail
olec.GetLectureDetail idx

%>
<script language="JavaScript">
<!--
function CheckForm(){
	if (document.lecform.yyyymm.value.length < 1){
		alert("월 구분을 등록해주세요");
		document.lecform.yyyymm.focus();
	}else if (document.lecform.linkitemid.value.length < 1){
		alert("제품번호를 등록해주세요");
		document.lecform.linkitemid.focus();
	}
	else if (document.lecform.lectitle.value.length < 1){
		alert("강좌명을 등록해주세요");
		document.lecform.lectitle.focus();
	}
	else if (document.lecform.lecturer.value.length < 1){
		alert("강사명을 등록해주세요");
		document.lecform.lecturer.focus();
	}
	else{
		document.lecform.action="lecture_act.asp";
		document.lecform.submit();
	}
}

function calender_open(objectname) {
//       document.all.cal.style.display="";
//	   document.all.cal.style.left = event.offsetX;
//	   document.all.cal.style.top = event.offsetY + 200;
//	   document.lecform.objname.value = objectname;

//	   alert("X-좌표 : " + event.offsetX + "\n" + "Y-좌표 : " + event.offsetY);
}

//-->

function popLectureItemList(frm){
	var popwin = window.open('lecregitems.asp','lecitem','width=600,height=500,status=no,resizable=yes,scrollbars=yes');
	popwin.focus();
}
function LectureAdd(){
	document.lecform.action="lecreg.asp";
	document.lecform.idx.value="";
	document.lecform.mode.value="add";
	document.lecform.submit();
}
function popLectureImg(){
	window.open ('lecregimg.asp','lecimg','width=800,height=500,status=no,resizable=yes,scrollbars=yes');
}
</script>
<form method=post name="lecform">
<input type="hidden" name="idx" value="<% = idx %>">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="objname">
<table width="800" border="0" cellpadding="0" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF">
	<td >Idx</td>
	<td bgcolor="#FFFFFF"> <% = olec.Fidx %></td>
</tr>
<% if mode = "add" then %>
<tr bgcolor="#DDDDFF">
	<td >월 구분</td>
	<td bgcolor="#FFFFFF"><input type="text" name="yyyymm" value="<%= yyyymm %>" size="7" maxlength="7">(2004-06)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >상품ID</td>
	<td bgcolor="#FFFFFF"><input type="text" name="linkitemid" value="0" size="6" maxlength="6">
	<input type="button" value="목록에서선택" onClick="popLectureItemList();">
	<input type="button" value="이미지불러오기" onClick="popLectureImg();>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강좌명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lectitle" value="<%= lectitle %>" size="50" maxlength="64"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >소속아이디</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lecturerid" value="<%= lecturerid %>" size="30" maxlength="32"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강사명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lecturer" value="<% =lecturer %>" size="30" maxlength="32"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강좌비</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="lecsum" value="<% if lecsum<>"" then response.write lecsum else response.write "0" end if %>" size="12" maxlength="12">
		<input type="checkbox" name="matinclude" <% if matinclude<>"" then response.write "checked" %>>재료비포함
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >재료비</td>
	<td bgcolor="#FFFFFF"><input type="text" name="matsum" value="<% if matsum<>"" then response.write matsum else response.write "0" end if %>" size="12" maxlength="12"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >장소</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lecspace" size="30" value="<%= lecspace %>" maxlength="64"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강좌횟수</td>
	<td bgcolor="#FFFFFF"><input type="text" name="leccount" value="<% if leccount<>"" then response.write leccount else response.write "0" end if %>" size="6" maxlength="12"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강의시간</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lectime" value="<% if lectime<>"" then response.write lectime else response.write "0" end if %>" size="20" maxlength="12"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >총강의시간</td>
	<td bgcolor="#FFFFFF"><input type="text" name="tottime" value="<% if tottime<>"" then response.write tottime else response.write "0" end if %>" size="6" maxlength="12"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강의기간<br>(주기)</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lecperiod" value="<% if lecperiod<>"" then response.write lecperiod else response.write "0" end if %>" size="30" maxlength="64">(ex : 매주 금요일 몇시~몇시)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >재료비설명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="matdesc" value="<%= matdesc %>" size="100" maxlength="128"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >적정인원</td>
	<td bgcolor="#FFFFFF"><input type="text" name="properperson" value="<% if properperson<>"" then response.write properperson else response.write "0" end if %>" size="6" maxlength="12"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>최소인원</td>
	<td bgcolor="#FFFFFF"><input type="text" name="minperson" value="<% if minperson<>"" then response.write minperson else response.write "0" end if %>" size="6" maxlength="12"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>예약등록일</td>
	<td bgcolor="#FFFFFF"><input type="text" name="reservestart" value="<%= reservestart %>" size="15" maxlength="10" onclick="calender_open('reservestart');"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>예약마감일</td>
	<td bgcolor="#FFFFFF"><input type="text" name="reserveend" value="<%= reserveend %>" size="15" maxlength="10" onclick="calender_open('reserveend');"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>강좌내용<br>(커리큘럼)</td>
	<td bgcolor="#FFFFFF">
			<table border="0" cellpadding="0" cellspacing="1" bgcolor="#3d3d3d" class="a">
			<tr bgcolor="#DDDDFF">
				<td>1주</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecdate01" value="<%= lecdate01 %>" size="20" maxlength="19" onclick="calender_open('lecdate01');">~<input type="text" name="lecdate01_end" value="<%= lecdate01_end %>" size="20" maxlength="19" onclick="calender_open('lecdate01_end');">(2004-06-06 14:00:00)</td>
			</tr>
			<tr bgcolor="#DDDDFF">
				<td>2주</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecdate02" value="<%= lecdate02 %>" size="20" maxlength="19" onclick="calender_open('lecdate02');">~<input type="text" name="lecdate02_end" value="<%= lecdate02_end %>" size="20" maxlength="19" onclick="calender_open('lecdate02_end');"></td>
			</tr>
			<tr bgcolor="#DDDDFF">
				<td>3주</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecdate03" value="<%= lecdate03 %>" size="20" maxlength="19" onclick="calender_open('lecdate03');">~<input type="text" name="lecdate03_end" value="<%= lecdate03_end %>" size="20" maxlength="19" onclick="calender_open('lecdate03_end');"></td>
			</tr>
			<tr bgcolor="#DDDDFF">
				<td>4주</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecdate04" value="<%= lecdate04 %>" size="20" maxlength="19" onclick="calender_open('lecdate04');">~<input type="text" name="lecdate04_end" value="<%= lecdate04_end %>" size="20" maxlength="19" onclick="calender_open('lecdate04_end');"></td>
			</tr>
			<tr bgcolor="#DDDDFF">
				<td>5주</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecdate05" value="<%= lecdate05 %>" size="20" maxlength="19" onclick="calender_open('lecdate05');">~<input type="text" name="lecdate05_end" value="<%= lecdate05_end %>" size="20" maxlength="19" onclick="calender_open('lecdate05_end');"></td>
			</tr>
			<tr bgcolor="#DDDDFF">
				<td>6주</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecdate06" value="<%= lecdate06 %>" size="20" maxlength="19" onclick="calender_open('lecdate06');">~<input type="text" name="lecdate06_end" value="<%= lecdate06_end %>" size="20" maxlength="19" onclick="calender_open('lecdate06_end');"></td>
			</tr>
			<tr bgcolor="#DDDDFF">
				<td>7주</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecdate07" value="<%= lecdate07 %>" size="20" maxlength="19" onclick="calender_open('lecdate07');">~<input type="text" name="lecdate07_end" value="<%= lecdate07_end %>" size="20" maxlength="19" onclick="calender_open('lecdate07_end');"></td>
			</tr>
			<tr bgcolor="#DDDDFF">
				<td>8주</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecdate08" value="<%= lecdate08 %>" size="20" maxlength="19" onclick="calender_open('lecdate08');">~<input type="text" name="lecdate08_end" value="<%= lecdate08_end %>" size="20" maxlength="19" onclick="calender_open('lecdate08_end');"></td>
			</tr>
			</table>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>강좌개요</td>
	<td bgcolor="#FFFFFF"><textarea name="leccontents" rows="10" cols="80"><%= leccontents %></textarea></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>커리큘럼소개</td>
	<td bgcolor="#FFFFFF"><textarea name="leccurry" rows="10" cols="80"><%= leccurry %></textarea></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>기타사항</td>
	<td bgcolor="#FFFFFF"><textarea name="lecetc" rows="10" cols="80"><%= lecetc %></textarea></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>접수종료</td>
	<td bgcolor="#FFFFFF">
	&nbsp;&nbsp;&nbsp;
	<% if regfinish="Y" then %>
	<input type=radio name=regfinish value=N > 접수중
	<input type=radio name=regfinish value=Y checked > 접수종료
	<% else %>
	<input type=radio name=regfinish value=N checked > 접수중
	<input type=radio name=regfinish value=Y > 접수종료
	<% end if %>
	</td>
	
</tr>
<tr bgcolor="#DDDDFF">
	<td>사용여부</td>
	<td bgcolor="#FFFFFF">
	&nbsp;&nbsp;&nbsp;
	<% if isusing ="Y" then %>
	<input type=radio name=isusing value=Y checked > 사용중(전시함)
	<input type=radio name=isusing value=N  > 사용안함(전시안함)
	<% else %>
	<input type=radio name=isusing value=Y  > 사용중(전시함)
	<input type=radio name=isusing value=N checked > 사용안함(전시안함)
	<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="right" height="30"><input type="button" value="내용저장" onclick="CheckForm();return false;">&nbsp;&nbsp;&nbsp;</td>
</tr>

<% else %>
<tr bgcolor="#DDDDFF">
	<td >월 구분</td>
	<td bgcolor="#FFFFFF"><input type="text" name="yyyymm" value="<% = olec.FMastercode %>" size="7" maxlength="7">(2004-06)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >상품ID</td>
	<td bgcolor="#FFFFFF">
	<input type="text" name="linkitemid" value="<% = olec.Flinkitemid %>" size="6" maxlength="6">
	<input type="button" value="목록에서선택" onClick="popLectureItemList();">
	<input type="button" value="새강좌에적용" onClick="LectureAdd();">
	<input type="button" value="이미지불러오기" onClick="popLectureImg()";>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강좌명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lectitle" value="<% = olec.Flectitle %>" size="50" maxlength="64"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >소속아이디</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lecturerid" value="<% = olec.Flecturerid %>" size="30" maxlength="32"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강사명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lecturer" value="<% = olec.Flecturer %>" size="30" maxlength="32"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강좌비</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="lecsum" value="<% = olec.Flecsum %>" size="12" maxlength="12">
		<input type="checkbox" name="matinclude" <% if olec.Fmatinclude = "Y" then response.write"checked" %>>재료비포함
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >재료비</td>
	<td bgcolor="#FFFFFF"><input type="text" name="matsum" value="<% = olec.Fmatsum %>" size="12" maxlength="12"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >장소</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lecspace" value="<% = olec.Flecspace %>" size="30" maxlength="64"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강좌횟수</td>
	<td bgcolor="#FFFFFF"><input type="text" name="leccount" value="<% = olec.Fleccount %>" size="6" maxlength="12"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강의시간</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lectime" value="<% = olec.Flectime %>" size="20" maxlength="12"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >총강의시간</td>
	<td bgcolor="#FFFFFF"><input type="text" name="tottime" value="<% = olec.Ftottime %>" size="6" maxlength="12"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강의기간<br>(주기)</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lecperiod" value="<% = olec.Flecperiod %>" size="30" maxlength="64">(ex : 매주 금요일 몇시~몇시)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >재료비설명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="matdesc" value="<% = olec.Fmatdesc %>" size="100" maxlength="128"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >적정인원</td>
	<td bgcolor="#FFFFFF"><input type="text" name="properperson" value="<% = olec.Fproperperson %>" size="6" maxlength="12"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>최소인원</td>
	<td bgcolor="#FFFFFF"><input type="text" name="minperson" value="<% = olec.Fminperson %>" size="6" maxlength="12"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>예약등록일</td>
	<td bgcolor="#FFFFFF"><input type="text" name="reservestart" value="<% = olec.Freservestart %>" size="15" maxlength="10" onclick="calender_open('reservestart');"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>예약마감일</td>
	<td bgcolor="#FFFFFF"><input type="text" name="reserveend" value="<% = olec.Freserveend %>" size="15" maxlength="10" onclick="calender_open('reserveend');"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>강좌내용<br>(커리큘럼)</td>
	<td bgcolor="#FFFFFF">
			<table border="0" cellpadding="0" cellspacing="1" bgcolor="#3d3d3d" class="a">
			<tr bgcolor="#DDDDFF">
				<td>1주</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecdate01" value="<% = olec.Flecdate01 %>" size="20" maxlength="19" onclick="calender_open('lecdate01');">~<input type="text" name="lecdate01_end" value="<% = olec.Flecdate01_end %>" size="20" maxlength="19" onclick="calender_open('lecdate01_end');"></td>
			</tr>
			<tr bgcolor="#DDDDFF">
				<td>2주</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecdate02" value="<% = olec.Flecdate02 %>" size="20" maxlength="19" onclick="calender_open('lecdate02');">~<input type="text" name="lecdate02_end" value="<% = olec.Flecdate02_end %>" size="20" maxlength="19" onclick="calender_open('lecdate02_end');"></td>
			</tr>
			<tr bgcolor="#DDDDFF">
				<td>3주</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecdate03" value="<% = olec.Flecdate03 %>" size="20" maxlength="19" onclick="calender_open('lecdate03');">~<input type="text" name="lecdate03_end" value="<% = olec.Flecdate03_end %>" size="20" maxlength="19" onclick="calender_open('lecdate03_end');"></td>
			</tr>
			<tr bgcolor="#DDDDFF">
				<td>4주</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecdate04" value="<% = olec.Flecdate04 %>" size="20" maxlength="19" onclick="calender_open('lecdate04');">~<input type="text" name="lecdate04_end" value="<% = olec.Flecdate04_end %>" size="20" maxlength="19" onclick="calender_open('lecdate04_end');"></td>
			</tr>
			<tr bgcolor="#DDDDFF">
				<td>5주</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecdate05" value="<% = olec.Flecdate05 %>" size="20" maxlength="19" onclick="calender_open('lecdate05');">~<input type="text" name="lecdate05_end" value="<% = olec.Flecdate05_end %>" size="20" maxlength="19" onclick="calender_open('lecdate05_end');"></td>
			</tr>
			<tr bgcolor="#DDDDFF">
				<td>6주</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecdate06" value="<% = olec.Flecdate06 %>" size="20" maxlength="19" onclick="calender_open('lecdate06');">~<input type="text" name="lecdate06_end" value="<% = olec.Flecdate06_end %>" size="20" maxlength="19" onclick="calender_open('lecdate06_end');"></td>
			</tr>
			<tr bgcolor="#DDDDFF">
				<td>7주</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecdate07" value="<% = olec.Flecdate07 %>" size="20" maxlength="19" onclick="calender_open('lecdate07');">~<input type="text" name="lecdate07_end" value="<% = olec.Flecdate07_end %>" size="20" maxlength="19" onclick="calender_open('lecdate07_end');"></td>
			</tr>
			<tr bgcolor="#DDDDFF">
				<td>8주</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecdate08" value="<% = olec.Flecdate08 %>" size="20" maxlength="19" onclick="calender_open('lecdate08');">~<input type="text" name="lecdate08_end" value="<% = olec.Flecdate08_end %>" size="20" maxlength="19" onclick="calender_open('lecdate08_end');"></td>
			</tr>
			</table>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>강좌개요</td>
	<td bgcolor="#FFFFFF"><textarea name="leccontents" rows="10" cols="80"><% = olec.Fleccontents %></textarea></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>커리큘럼소개</td>
	<td bgcolor="#FFFFFF"><textarea name="leccurry" rows="10" cols="80"><% = olec.Fleccurry %></textarea></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>기타사항</td>
	<td bgcolor="#FFFFFF"><textarea name="lecetc" rows="10" cols="80"><% = olec.Flecetc %></textarea></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>접수종료</td>
	<td bgcolor="#FFFFFF">
	&nbsp;&nbsp;&nbsp;
	<% if olec.FRegFinish="Y" then %>
	<input type=radio name=regfinish value=N > 접수중
	<input type=radio name=regfinish value=Y checked > 접수종료
	<% else %>
	<input type=radio name=regfinish value=N checked > 접수중
	<input type=radio name=regfinish value=Y > 접수종료
	<% end if %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>사용여부</td>
	<td bgcolor="#FFFFFF">
	&nbsp;&nbsp;&nbsp;
	<% if olec.FIsUsing ="Y" then %>
	<input type=radio name=isusing value=Y checked > 사용중(전시함)
	<input type=radio name=isusing value=N  > 사용안함(전시안함)
	<% else %>
	<input type=radio name=isusing value=Y  > 사용중(전시함)
	<input type=radio name=isusing value=N checked > 사용안함(전시안함)
	<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="right" height="30"><input type="button" value="내용저장" onclick="CheckForm();return false;">&nbsp;&nbsp;&nbsp;</td>
</tr>
<% end if %>
</table>
</form>
<%
set olec = Nothing
%>

<div style="display:none;position:absolute; width:200px; height:100px; z-index:1" id="cal">
<table cellpadding="0" cellspacing="0" border="0" bgcolor="white">
<tr>
	<td align="center">
		<table width="245" cellspacing="0" cellpadding="0" border="0" align="center">
				<tr>
						<td align="center" width="40" height="30"><input type="button" class="button" value="◀◀" onclick="to_PreYear()"></td>
						<td align="center" width="30"><input type="button" class="button" value="◁" onclick="to_PreMonth()"></td>
						<td align="center" width="105"><div id="cal_title" style="color:#8FACCC"></div></td>
						<td align="center" width="30"><input type="button" class="button" value="▷" onclick="to_NextMonth()"></td>
				<td align="center" width="40"><input type="button" class="button" value="▶▶" onclick="to_NextYear()"></td>
				</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
<!-- 달력 출력 부분 -->
		<table width="245" cellspacing="0" cellpadding="0" align="center" id="cal_Table">
		</table>
	</td>
</tr>
<tr>
	<td align="center">
<!-- Button -->
		<table width="245" cellspacing="0" cellpadding="0" border="0">
			<tr>
				<td height="10"></td>
			</tr>
			<tr>
				<td align="center"><input type="button" name='today' class="button" value="Today" style="font-family:verdana" onClick="writeValue()"></td>
				<td align="center"><input type="button" name='none' class="button" value="None" style="font-family:verdana" onClick="writeValue()"></td>
			</tr>
		</table>
	</td>
</tr>
</table>
</div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->