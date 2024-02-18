<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 메일진 등록
' History : 2007.12.20 한용민 수정
'###########################################################
%>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/mailzine/mailzinecls.asp"-->
<%
Dim omail,ix,page , isusing , username , usermail
	page = requestCheckVar(getNumeric(trim(request("page"))),10)
	isusing = requestCheckVar(trim(request("isusing")),1)
	username = requestCheckVar(trim(request("username")),32)
	usermail = requestCheckVar(trim(request("usermail")),128)

if page = "" then page = 1
if isusing = "" then isusing = "Y"
	
set omail = new CMailzineList
	omail.FPageSize = 50
	omail.FCurrPage = page
	omail.frectisusing = isusing
	omail.frectusername = username
	omail.frectusermail = usermail
	omail.MailzineList
%>
<script type='text/javascript'>

function addreg(idx){
	var addreg = window.open('/cscenter/mailzine/cs_mailzine_reg.asp?idx='+idx,'addreg','width=800,height=400,scrollbars=yes,resizable=yes');
	addreg.focus();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method=get action="" style="margin:0px;">
<input type="hidden" name="editor_no">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		사용여부: <select name="isusing" value="<%=isusing%>">
			<option value="" <% if isusing = "" then response.write " selected" %>>사용여부</option>
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>			 			
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		이름: <input type="text" name="username" value="<%=username%>">
		&nbsp;이메일: <input type="text" name="usermail" value="<%=usermail%>">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<Br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		* 이름이나 이메일 주소를 공백 없이 입력하셔야 내역이 나옵니다.
		<br>* 사용여부 Y 인경우에만 비회원고객님께 메일진이 발송 됩니다.
	</td>
	<td align="right">
		<input type="button" value="신규등록" onclick="addreg('','add');" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<form method=post name="monthly" style="margin:0px;">
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="17">
		검색결과 : <b><%= omail.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= omail.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">idx</td>
	<td align="center">이름</td>
	<td align="center">이메일</td>
	<td align="center">등록일</td>		
	<td align="center">사용여부</td>
	<td align="center">비고</td>	
</tr>
<% if omail.FresultCount>0 then %>	
	<% for ix=0 to omail.FresultCount-1 %>
		<tr align="center" bgcolor="#FFFFFF">
			<td align="center"><% = omail.FItemList(ix).Fidx %></td>
			<td align="center"><% = ReplaceBracket(omail.FItemList(ix).fusername) %></td>
			<td align="center"><% = ReplaceBracket(omail.FItemList(ix).fusermail) %></td>		
			<td align="center"><% = FormatDate(omail.FItemList(ix).fregdate,"0000.00.00") %></td>
			<td align="center"><% = omail.FItemList(ix).fisusing %></td>
			<td align="center">
				<input type="button" value="수정" class="button" onclick="addreg(<% = omail.FItemList(ix).Fidx %>);">
			</td>				
		</tr>
	<% next %>
	
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if omail.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= omail.StarScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for ix = 0 + omail.StarScrollPage to omail.StarScrollPage + omail.FScrollCount - 1 %>
				<% if (ix > omail.FTotalpage) then Exit for %>
				<% if CStr(ix) = CStr(omail.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= ix %></b></font></span>
				<% else %>
				<a href="?page=<%= ix %>" class="list_link"><font color="#000000"><%= ix %></font></a>
				<% end if %>
			<% next %>
			<% if omail.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= ix %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="7" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
</form>

<% set omail = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
