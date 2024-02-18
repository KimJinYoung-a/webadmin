<%@ language=vbscript %>
<% option explicit %>
<%
'############### 2008년 11월 4일 한용민 생성
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->
<%
dim idx
	idx = request("idx")

'//수정모드
if idx = "" then
	response.write "<script>"
	response.write "alert('idx 파라메타 값이없습니다');"
	response.write "history.go(-1);"
	response.write "</script>"
end if

dim oip
set oip = new DiaryCls
	oip.frectidx = idx
	oip.geteventone
%>

※ 해당이벤트가 사용여부Y , 이벤트진행중 일경우만 노출됩니다. 조건을 벗어나면 등록하셔도 자동으로 노출되지 않습니다.
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" action="/admin/diary2009/event_process.asp" mothod="get">
	<input type="hidden" name="mode" value="edit">
	<input type="hidden" name="idx" value="<%= oip.fitem.fidx %>">
	<input type="hidden" name="event_type" value="<%= oip.fitem.fevent_type %>">
	<tr bgcolor="FFFFFF" align="center">
		<td>
			링크타입
		</td>
		<td>
			<select name="event_link">
				<option value="event" <% if oip.fitem.fevent_link = "event" or oip.fitem.fevent_link=""  then response.write " selected"%>>event</option>
				<option value="item" <% if oip.fitem.fevent_link = "item" then response.write " selected"%>>item</option>
			</select><br>
			item를 선택시 하단상품코드 상품페이지로 이동<br>
			event를 선택시 하단이벤트코드 이벤트페이지로 이동
		</td>
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td>이벤트코드</td>
		<td><input type="text" name="evt_code" value="<%= oip.fitem.fevt_code %>"></td>
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td>상품코드</td>
		<td><input type="text" name="itemid" value="<%= oip.fitem.fitemid %>"></td>
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td>구분</td>
		<td><% SelectList "cate", oip.fitem.FCateCode %></td>
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td>노출순서</td>
		<td><input type="text" name="idx_order" value="<%= oip.fitem.fidx_order %>">기본값0</td>
	</tr>

	<tr bgcolor="FFFFFF" align="center">
		<td>사용여부</td>
		<td>
			<select name="isusing">
				<option value="Y" <% if oip.fitem.fisusing = "Y" then response.write " selected"%>>Y</option>
				<option value="N" <% if oip.fitem.fisusing = "N" or oip.fitem.fisusing="" then response.write " selected"%>>N</option>
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


