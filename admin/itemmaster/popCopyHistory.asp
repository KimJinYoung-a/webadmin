<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
Dim oItem, page, i
Dim itemid, resultCode, lastUserid, sellyn, pagesize, errMsg
itemid		= request("itemid")
resultCode	= request("resultCode")
page 		= request("page")
lastUserid	= request("lastUserid")
sellyn		= request("sellyn")
pagesize	= request("pagesize")
errMsg		= request("errMsg")

If page = "" Then page = 1
If pagesize = "" Then pagesize = 100

If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If

Set oItem = new CItem
	oItem.FPageSize 		= pagesize
	oItem.FCurrPage			= page
	oItem.FRectItemid 		= itemid
	oItem.FRectResultCode 	= resultCode
	oItem.FRectErrMsg	 	= errMsg
	oItem.getItemCopyHistoryList
%>
<script>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		성공여부 :
		<select name="resultCode" class="select">
			<option value="">전체</option>
			<option value="OK"  	<%= Chkiif(resultCode = "OK", "selected", "")%> >성공</option>
			<option value="ERR"		<%= Chkiif(resultCode = "ERR", "selected", "")%> >에러</option>
		</select>
		&nbsp;
		표시갯수 :
		<select name="pagesize" class="select">
			<option value="20"  <%= Chkiif(pagesize = "20", "selected", "")%> >20</option>
			<option value="100"  <%= Chkiif(pagesize = "100", "selected", "")%> >100</option>
			<option value="200"  <%= Chkiif(pagesize = "200", "selected", "")%> >200</option>
			<option value="500"  <%= Chkiif(pagesize = "500", "selected", "")%> >500</option>
		</select>
		&nbsp;
		Message 검색 : <input type="text" name="errMsg" id="errMsg" value="<%= errMsg %>">
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="subcmd" value="">
<input type="hidden" name="chgSellYn" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="14">
		검색결과 : <b><%= FormatNumber(oItem.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oItem.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="7%">상품코드</td>
	<td width="7%">복제상품코드</td>
	<td width="15%">등록시간</td>
	<td width="15%">API완료시간</td>
	<td width="6%">성공여부</td>
	<td width="10%">수행ID</td>
	<td width="40%">Message</td>
</tr>
<% For i = 0 To oItem.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= oItem.FItemlist(i).FItemid %></td>
	<td><%= oItem.FItemlist(i).FCopyitemid %></td>
	<td><%= oItem.FItemlist(i).FRegdate %></td>
	<td><%= oItem.FItemlist(i).FFindate %></td>
	<td>
	<%
		Select Case oItem.FItemlist(i).FResultCode
			Case "OK"		response.write "<font color='BLUE'>"&oItem.FItemlist(i).FResultCode&"</font>"
			Case "ERR"		response.write "<font color='RED'>"&oItem.FItemlist(i).FResultCode&"</font>"
			Case Else		response.write "<font color='GRAY'>"&oItem.FItemlist(i).FResultCode&"</font>"
		End Select
	%>
	</td>
	<td><%= oItem.FItemlist(i).FLastUserid %></td>
	<td width="300"><font title='<%= oItem.FItemlist(i).FLastErrMsg %>'><%= left(oItem.FItemlist(i).FLastErrMsg, 120) %></font></td>
</tr>
<%
	Next
%>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="14" align="center">
	<% If oItem.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oItem.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + oItem.StartScrollPage To oItem.FScrollCount + oItem.StartScrollPage - 1 %>
		<% If i>oItem.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If oItem.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
</form>
</table>
<% Set oItem = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
