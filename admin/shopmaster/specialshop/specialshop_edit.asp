<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 우수회원샵
' Hieditor : 2009.12.28 한용민 생성
'			 2022.07.06 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/specialshop/specialshop_cls.asp"-->

<%
dim id,openDate,status,regdate , i , statusstr , itemcount , isusing, title, endDate
dim mode
	id = requestCheckVar(getNumeric(request("id")),10)
	mode = requestCheckVar(request("mode"),32)

dim ospecialshop
set ospecialshop = new cspecialshop_list
	ospecialshop.frectid = id
	
	'//수정모드 일경우만 쿼리
	if id <> "" then
	ospecialshop.fspecialshop_oneitem()

		if ospecialshop.ftotalcount > 0 then
			statusstr = ospecialshop.FOneItem.fstatusstr
			openDate = formatdate(ospecialshop.FOneItem.fopenDate,"0000-00-00")
			status = ospecialshop.FOneItem.fstatus
			regdate = ospecialshop.FOneItem.fregdate
			itemcount = ospecialshop.FOneItem.fitemcount
			isusing = ospecialshop.FOneItem.fisusing
			title = ReplaceBracket(ospecialshop.FOneItem.ftitle)
			endDate = ospecialshop.FOneItem.FendDate
		end if
	end if
%>

<script type='text/javascript'>

// 등록&수정
function reg(){
	if (frm.title.value==''){
		alert('테마를 등록하세요');
		frm.title.focus();
		return;
	}
	
	if (frm.openDate.value==''){
		alert('오픈일을 등록하세요');
		frm.openDate.focus();
		return;
	}
	
	<% If status = "" OR status = "1" Then %>
	if (frm.endDate.value==''){
		alert('종료일을 등록하세요');
		frm.endDate.focus();
		return;
	}
	<% End If %>

	if (frm.status.value==''){
		alert('상태를 선택 하세요');
		frm.status.focus();
		return;
	}
	
	if (frm.isusing.value==''){
		alert('사용여부를 선택 하세요');
		frm.isusing.focus();
		return;
	}	
	
	frm.mode.value='reg';	
	frm.action='/admin/shopmaster/specialshop/specialshop_process.asp';
	frm.submit();
}

</script>

<form name="frm" method="post" style="margin:0px;">
<input type="hidden" name="mode">
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25" width="70">ID</td>
	<td bgcolor="#FFFFFF" align="left">&nbsp;
		<%= id %><input type="hidden" name="id" value="<%= id %>">		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">테마</td>
	<td bgcolor="#FFFFFF" align="left">&nbsp;
		<input type="text" name="title" size="70" value="<%= title %>">			
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">오픈일</td>
	<td bgcolor="#FFFFFF" align="left">&nbsp;
	<% If status = "" OR status = "0" OR status = "1" Then %>
		<input type="text" name="openDate" size=10 value="<%= openDate %>">	
		<a href="javascript:calendarOpen3(frm.openDate,'시작일',frm.openDate.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
	<% Else %>
		<%= openDate %><input type="hidden" name="openDate" value="<%= openDate %>">
	<% End If %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">종료일</td>
	<td bgcolor="#FFFFFF" align="left">&nbsp;
	<% If status = "" OR status = "0" OR status = "1" Then %>
		<input type="text" name="endDate" size=10 value="<%= endDate %>">			
		<a href="javascript:calendarOpen3(frm.endDate,'종료일',frm.endDate.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
	<% Else %>
		<%= endDate %><input type="hidden" name="endDate" value="<%= endDate %>">
	<% End If %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="75">상태</td>
	<td bgcolor="#FFFFFF" align="left">&nbsp;
	<% If status = "" then status = "0" end if %>
		<% drawstatus "status" , status, id %>
		<br><br>&nbsp;* 대기로 설정하면 오픈일 00시에 사용여부가 Y인것들은 자동 오픈 됩니다.
		<br>&nbsp;* 종료일이 지나면 자동 종료 됩니다.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">사용여부</td>
	<td bgcolor="#FFFFFF" align="left">&nbsp;
		<select name="isusing" value="<%=isusing%>">
			<option value="" <% if isusing = "" then response.write " selected" %>>선택</option>
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>			
	</td>
</tr>
<tr align="center" bgcolor="FFFFFF">
	<td colspan=2 height="50"><input type="button" onclick="reg();" value=" 저 장 " class="button" style="width:80px;height:40px;"></td>
</tr>
</table>
</form>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->