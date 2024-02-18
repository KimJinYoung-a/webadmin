<%@ language=vbscript %>
<% option explicit %>
<%
'############### 2008년 11월 6일 한용민 생성
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->

  
※ 해당이벤트가 사용여부Y , 이벤트진행중 일경우만 노출됩니다. 조건을 벗어나면 등록하셔도 자동으로 노출되지 않습니다. 
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" action="/admin/diary2009/event_process.asp" mothod="get">
	<input type="hidden" name="mode" value="new">
	<input type="hidden" name="event_type" value="organizer">	
	<tr bgcolor="FFFFFF" align="center">
		<td>이벤트코드</td>
		<td>노출순서</td>
		<td>사용여부</td>
	</tr>
	
		
	<tr bgcolor="FFFFFF" align="center">
		<td><input type="text" name="evt_code" value=""></td>
		<td><input type="text" name="idx_order" value=""></td>
		<td>
			<select name="isusing">
				<option>선택하세요</option>
				<option value="Y">Y</option>
				<option value="N" checked>N</option>				
			</select>
		</td>
	</tr>
	
	<tr bgcolor="FFFFFF" align="left">
		<td colspan=3><input type="button" class="button" value="저장" onclick="javascript:frm.submit();"></td>
	</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->