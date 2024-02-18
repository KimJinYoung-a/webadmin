<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  세미나실 관리
' History : 2009.04.07 서동석 생성
'			2010.12.27 한용민 수정
'           2012.01.10 허진원 수정; 세미나실 정리/추가
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/academy/seminar_roomcls.asp"-->
<%
dim idx,mode,tdate,ttime,roomid , osemi
	idx = request("idx")
	mode = request("mode")
	roomid = request("roomid")
	tdate = request("tdate")
	ttime = request("ttime")
%>

<script language='javascript'>

function SubmitForm(){
	<% if mode = "write" then %>
		if (document.SubmitFrm.roomid.value.length<1){
			alert('방을 선택해주세요');
			document.SubmitFrm.roomid.focus();
			return;
		}
		if (document.SubmitFrm.username.value.length<1){
			alert('고객명을 적어주세요');
			document.SubmitFrm.username.focus();
			return;
		}
		if (document.SubmitFrm.userphone.value.length < 1){
			alert('고객 연락처를 적어주세요');
			document.SubmitFrm.userphone.focus();
			return;
		}
	<% end if %>

	if (document.SubmitFrm.basictime.value.length < 1){
		alert('시간을 선택해주세요');
		document.SubmitFrm.basictime.focus();
		return;
	}

	var ret = confirm('저장 하시겠습니까?');
	if (ret) {
		document.SubmitFrm.submit();
	}
}

</script>

<table width="400" border="1" cellpadding="0" cellspacing="0" class="a" bordercolordark="White" bordercolorlight="black" align="center">
<form name="SubmitFrm" method="post" action="doseminarroom.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="<% = mode %>">
<%
if mode = "modify" then

set osemi = new CSeminarRoomDetail
	osemi.read idx
%>
<input type="hidden" name="idx" value="<% = idx %>">
<tr>
  <td width="100">방선택</td>
  <td>
		<select name="roomid">
			<option value="">선택</option>
			<!--
			<option value="05" <% if osemi.Froomid = "05" then response.write "selected" %>>Television</option>
			<option value="07" <% if osemi.Froomid = "07" then response.write "selected" %>>Bingo</option>
			<option value="06" <% if osemi.Froomid = "06" then response.write "selected" %>>Chocolate</option>
			-->
			<option value="01" <% if osemi.Froomid = "01" then response.write "selected" %>>Idea</option>
			<option value="02" <% if osemi.Froomid = "02" then response.write "selected" %>>Paper</option>
			<option value="03" <% if osemi.Froomid = "03" then response.write "selected" %>>Heart</option>
			<option value="04" <% if osemi.Froomid = "04" then response.write "selected" %>>Fingers</option>
			<option value="08" <% if osemi.Froomid = "08" then response.write "selected" %>>Moon</option>
			<option value="09" <% if osemi.Froomid = "09" then response.write "selected" %>>Star</option>
			<option value="10" <% if osemi.Froomid = "10" then response.write "selected" %>>Step</option>
			<option value="11" <% if osemi.Froomid = "11" then response.write "selected" %>>Music</option>
		</select>
  </td>
</tr>
<tr>
  <td width="100">예약일</td>
  <td>
  <input type="text" name="tdate" value="<% = tdate %>" class="input_b" size="10">일
  <select name="usetime">
		<option value="8" <% if osemi.Fusetime = "8" then response.write "selected" %>>10:00</option>
		<option value="9" <% if osemi.Fusetime = "9" then response.write "selected" %>>10:30</option>
		<option value="10" <% if osemi.Fusetime = "10" then response.write "selected" %>>11:00</option>
		<option value="11" <% if osemi.Fusetime = "11" then response.write "selected" %>>11:30</option>
		<option value="12" <% if osemi.Fusetime = "12" then response.write "selected" %>>12:00</option>
		<option value="13" <% if osemi.Fusetime = "13" then response.write "selected" %>>12:30</option>
		<option value="14" <% if osemi.Fusetime = "14" then response.write "selected" %>>13:00</option>
		<option value="15" <% if osemi.Fusetime = "15" then response.write "selected" %>>13:30</option>
		<option value="16" <% if osemi.Fusetime = "16" then response.write "selected" %>>14:00</option>
		<option value="17" <% if osemi.Fusetime = "17" then response.write "selected" %>>14:30</option>
		<option value="18" <% if osemi.Fusetime = "18" then response.write "selected" %>>15:00</option>
		<option value="19" <% if osemi.Fusetime = "19" then response.write "selected" %>>15:30</option>
		<option value="20" <% if osemi.Fusetime = "20" then response.write "selected" %>>16:00</option>
		<option value="21" <% if osemi.Fusetime = "21" then response.write "selected" %>>16:30</option>
		<option value="22" <% if osemi.Fusetime = "22" then response.write "selected" %>>17:00</option>
		<option value="23" <% if osemi.Fusetime = "23" then response.write "selected" %>>17:30</option>
		<option value="24" <% if osemi.Fusetime = "24" then response.write "selected" %>>18:00</option>
		<option value="25" <% if osemi.Fusetime = "25" then response.write "selected" %>>18:30</option>
		<option value="26" <% if osemi.Fusetime = "26" then response.write "selected" %>>19:00</option>
		<option value="27" <% if osemi.Fusetime = "27" then response.write "selected" %>>19:30</option>
		<option value="28" <% if osemi.Fusetime = "28" then response.write "selected" %>>20:00</option>
		<option value="29" <% if osemi.Fusetime = "29" then response.write "selected" %>>20:30</option>
		<option value="30" <% if osemi.Fusetime = "30" then response.write "selected" %>>21:00</option>
		<option value="31" <% if osemi.Fusetime = "31" then response.write "selected" %>>21:30</option>
		<option value="32" <% if osemi.Fusetime = "32" then response.write "selected" %>>22:00</option>
		<option value="33" <% if osemi.Fusetime = "33" then response.write "selected" %>>22:30</option>
		<option value="34" <% if osemi.Fusetime = "34" then response.write "selected" %>>23:00</option>
		<option value="35" <% if osemi.Fusetime = "35" then response.write "selected" %>>23:30</option>
  </select>시
  <select name="basictime">
  		<option value="1" <% if osemi.Fbasictime = "1" then response.write "selected" %>>1</option>
		<option value="1.5" <% if osemi.Fbasictime = "1.5" then response.write "selected" %>>1.5</option>
		<option value="2" <% if osemi.Fbasictime = "2" then response.write "selected" %>>2</option>
		<option value="2.5" <% if osemi.Fbasictime = "2.5" then response.write "selected" %>>2.5</option>
		<option value="3" <% if osemi.Fbasictime = "3" then response.write "selected" %>>3</option>
		<option value="3.5" <% if osemi.Fbasictime = "3.5" then response.write "selected" %>>3.5</option>
		<option value="4" <% if osemi.Fbasictime = "4" then response.write "selected" %>>4</option>
		<option value="4.5" <% if osemi.Fbasictime = "4.5" then response.write "selected" %>>4.5</option>
		<option value="5" <% if osemi.Fbasictime = "5" then response.write "selected" %>>5</option>
		<option value="5.5" <% if osemi.Fbasictime = "5.5" then response.write "selected" %>>5.5</option>
		<option value="6" <% if osemi.Fbasictime = "6" then response.write "selected" %>>6</option>
		<option value="6.5" <% if osemi.Fbasictime = "6.5" then response.write "selected" %>>6.5</option>
		<option value="7" <% if osemi.Fbasictime = "7" then response.write "selected" %>>7</option>
		<option value="7.5" <% if osemi.Fbasictime = "7.5" then response.write "selected" %>>7.5</option>
		<option value="8" <% if osemi.Fbasictime = "8" then response.write "selected" %>>8</option>
		<option value="8.5" <% if osemi.Fbasictime = "8.5" then response.write "selected" %>>8.5</option>
		<option value="9" <% if osemi.Fbasictime = "9" then response.write "selected" %>>9</option>
		<option value="9.5" <% if osemi.Fbasictime = "9.5" then response.write "selected" %>>9.5</option>
		<option value="10" <% if osemi.Fbasictime = "10" then response.write "selected" %>>10</option>
  </select>시간
  </td>
</tr>
<tr>
  <td width="100">그룹명</td>
  <td><input type="text" name="groupname" value="<% = osemi.Fgroupname %>" size="15" class="input_b"></td>
</tr>
<tr>
  <td width="100">고객명</td>
  <td><input type="text" name="username" value="<% = osemi.Fusername %>" size="15" class="input_b"></td>
</tr>
<tr>
  <td width="100">고객연락처</td>
  <td><input type="text" name="userphone" value="<% = osemi.Fuserphone %>" size="15" class="input_b"></td>
</tr>
<tr>
  <td width="100">사용인원</td>
  <td><input type="text" name="usepeople" value="<% = osemi.Fusepeople %>" size="15" class="input_b"></td>
</tr>
<tr>
  <td width="100">기타사항</td>
  <td><textarea name="etc" rows="5" cols="40" class="input_b"><% = osemi.Fetc %></textarea></td>
</tr>
<tr>
  <td width="100">사용여부</td>
  <td>
  	<input type="radio" name="isusing" value="Y" <% if osemi.FIsUsing="Y" then response.write "checked" %> >Y
  	<input type="radio" name="isusing" value="N" <% if osemi.FIsUsing="N" then response.write "checked" %> >N
  </td>
</tr>
<tr>
	<td>강좌번호</td>
	<td>
		<input type="text" name="lecturer_idx" value="<%= osemi.flecturer_idx %>" size=10> ※없으면 공란
	</td>
</tr>
<tr>
  <td colspan="2" align="center">
  	<input type="button" value="저장" onClick="SubmitForm()" class="button">
  </td>
</tr>
</form>
</table>
<%
else
%>
<tr>
  <td width="100">방선택</td>
  <td>
		<select name="roomid">
			<option value="">선택</option>
			<!--
			<option value="05" <% if roomid = "05" then response.write "selected" %>>Television</option>
			<option value="07" <% if roomid = "07" then response.write "selected" %>>Bingo</option>
			<option value="06" <% if roomid = "06" then response.write "selected" %>>Chocolate</option>
			-->
			<option value="01" <% if roomid = "01" then response.write "selected" %>>Idea</option>
			<option value="02" <% if roomid = "02" then response.write "selected" %>>Paper</option>
			<option value="03" <% if roomid = "03" then response.write "selected" %>>Heart</option>
			<option value="04" <% if roomid = "04" then response.write "selected" %>>Fingers</option>
			<option value="08" <% if roomid = "08" then response.write "selected" %>>Moon</option>
			<option value="09" <% if roomid = "09" then response.write "selected" %>>Star</option>
			<option value="10" <% if roomid = "10" then response.write "selected" %>>Step</option>
			<option value="11" <% if roomid = "11" then response.write "selected" %>>Music</option>
		</select>
  </td>
</tr>
<tr>
  <td width="100">예약일</td>
  <td>
		<input type="text" name="tdate" value="<% = tdate %>" class="input_b" size="10">일
		<select name="usetime">
			<option value="8" <% if ttime = "8" then response.write "selected" %>>10:00</option>
			<option value="9" <% if ttime = "9" then response.write "selected" %>>10:30</option>
			<option value="10" <% if ttime = "10" then response.write "selected" %>>11:00</option>
			<option value="11" <% if ttime = "11" then response.write "selected" %>>11:30</option>
			<option value="12" <% if ttime = "12" then response.write "selected" %>>12:00</option>
			<option value="13" <% if ttime = "13" then response.write "selected" %>>12:30</option>
			<option value="14" <% if ttime = "14" then response.write "selected" %>>13:00</option>
			<option value="15" <% if ttime = "15" then response.write "selected" %>>13:30</option>
			<option value="16" <% if ttime = "16" then response.write "selected" %>>14:00</option>
			<option value="17" <% if ttime = "17" then response.write "selected" %>>14:30</option>
			<option value="18" <% if ttime = "18" then response.write "selected" %>>15:00</option>
			<option value="19" <% if ttime = "19" then response.write "selected" %>>15:30</option>
			<option value="20" <% if ttime = "20" then response.write "selected" %>>16:00</option>
			<option value="21" <% if ttime = "21" then response.write "selected" %>>16:30</option>
			<option value="22" <% if ttime = "22" then response.write "selected" %>>17:00</option>
			<option value="23" <% if ttime = "23" then response.write "selected" %>>17:30</option>
			<option value="24" <% if ttime = "24" then response.write "selected" %>>18:00</option>
			<option value="25" <% if ttime = "25" then response.write "selected" %>>18:30</option>
			<option value="26" <% if ttime = "26" then response.write "selected" %>>19:00</option>
			<option value="27" <% if ttime = "27" then response.write "selected" %>>19:30</option>
			<option value="28" <% if ttime = "28" then response.write "selected" %>>20:00</option>
			<option value="29" <% if ttime = "29" then response.write "selected" %>>20:30</option>
			<option value="30" <% if ttime = "30" then response.write "selected" %>>21:00</option>
			<option value="31" <% if ttime = "31" then response.write "selected" %>>21:30</option>
			<option value="32" <% if ttime = "32" then response.write "selected" %>>22:00</option>
			<option value="33" <% if ttime = "33" then response.write "selected" %>>22:30</option>
			<option value="34" <% if ttime = "34" then response.write "selected" %>>23:00</option>
			<option value="35" <% if ttime = "35" then response.write "selected" %>>23:30</option>
		</select>시
		<select name="basictime">
			 <option value="1" >1</option>
			 <option value="1.5" >1.5</option>
			 <option value="2" selected >2</option>
			 <option value="2.5">2.5</option>
			 <option value="3">3</option>
			 <option value="3.5">3.5</option>
			 <option value="4">4</option>
			 <option value="4.5">4.5</option>
			 <option value="5">5</option>
			 <option value="5.5">5.5</option>
			 <option value="6">6</option>
			 <option value="6.5">6.5</option>
			 <option value="7">7</option>
			 <option value="7.5">7.5</option>
			 <option value="8">8</option>
			 <option value="8.5">8.5</option>
			 <option value="9">9</option>
			 <option value="9.5">9.5</option>
			 <option value="10">10</option>
		</select>시간
  </td>
</tr>
<tr>
  <td width="100">고객명</td>
  <td><input type="text" name="username" size="15" class="input_b"></td>
</tr>
<tr>
  <td width="100">고객연락처</td>
  <td><input type="text" name="userphone" size="15" class="input_b"></td>
</tr>
<tr>
  <td width="100">사용인원</td>
  <td><input type="text" name="usepeople" value="0" size="15" class="input_b"></td>
</tr>
<tr>
  <td width="100">기타사항</td>
  <td><textarea name="etc" rows="5" cols="40" class="input_b"></textarea></td>
</tr>
<tr>
  <td width="100">사용여부</td>
  <td>
  	<input type="radio" name="isusing" value="Y">Y
  	<input type="radio" name="isusing" value="N">N
  </td>
</tr>
<tr>
	<td>강좌번호</td>
	<td>
		<input type="text" name="lecturer_idx" size=10> ※없으면 공란
	</td>
</tr>
<tr>
	<td colspan="2" align="center">
		<input type="button" value="저장" onClick="SubmitForm()"  class="button">
	</td>
</tr>
</form>
</table>
<% end if %>

<%
set osemi = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->