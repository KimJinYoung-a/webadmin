<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 한줄낙서
' Hieditor : 2010.11.23 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim ooneline,i , onelineid,startdate,enddate,winnerdate,comment,regdate,isusing , winner ,winnercomment
	onelineid = requestcheckvar(request("onelineid"),8)

'//상세
set ooneline = new coneline_list
	ooneline.frectonelineid = onelineid
	
	'//수정일경우에만 쿼리
	if onelineid <> "" then
		ooneline.foneline_oneitem()
	end if
		
	if ooneline.ftotalcount > 0 then
		onelineid = ooneline.FOneItem.fonelineid
		startdate = ooneline.FOneItem.fstartdate
		enddate = ooneline.FOneItem.fenddate
		winnerdate = ooneline.FOneItem.fwinnerdate
		comment = ooneline.FOneItem.fcomment
		regdate = ooneline.FOneItem.fregdate
		isusing = ooneline.FOneItem.fisusing
		winner = ooneline.FOneItem.fwinner
		winnercomment = ooneline.FOneItem.fwinnercomment			
	end if
%>

<script language="javascript">

	//저장
	function reg(){
		if (frm.startdate.value==''){
		alert('시작일을 입력해주세요');
		frm.startdate.focus();
		return;
		}		
		if (frm.enddate.value==''){
		alert('종료일을 입력해주세요');
		frm.enddate.focus();
		return;
		}
		if (frm.winnerdate.value==''){
		alert('당첨일을 입력해주세요');
		frm.winnerdate.focus();
		return;
		}							
		if (frm.isusing.value==''){
		alert('사용여부를 선택해주세요');
		return;
		}
		
		frm.action='oneline_process.asp';
		frm.mode.value='edit';
		frm.submit();
	}
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post">
<input type="hidden" name="mode">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>ID</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= onelineid %><input type="hidden" name="onelineid" value="<%= onelineid %>">		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><b>기간</b><br></td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="startdate" size=10 value="<%= startdate %>">			
		<a href="javascript:calendarOpen3(frm.startdate,'시작일',frm.startdate.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a> -
		<input type="text" name="enddate" size=10  value="<%= left(enddate,10) %>">
		<a href="javascript:calendarOpen3(frm.enddate,'마지막일',frm.enddate.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>	
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><b>당첨일</b><br></td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="winnerdate" size=10 value="<%= winnerdate %>">			
		<a href="javascript:calendarOpen3(frm.winnerdate,'당첨일',frm.winnerdate.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>내용</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="comment" style="width:450px; height:100px;"><%=comment%></textarea>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>당첨자ID</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="winner" value="<%=winner%>"> ※노출이 필요한 경우만 입력하세요	
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>당첨내역</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="winnercomment" style="width:450px; height:100px;"><%=winnercomment%></textarea>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>사용여부</td>
	<td bgcolor="#FFFFFF" align="left">
		<select name="isusing" value="<%=isusing%>">
			<option value="" <% if isusing = "" then response.write " selected" %>>사용여부</option>
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>			
	</td>
</tr>
<tr align="center" bgcolor="FFFFFF">
	<td colspan=2><input type="button" onclick="reg();" value="저장" class="button"></td>
</tr>
</form>
</table>
<%
	set ooneline = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->