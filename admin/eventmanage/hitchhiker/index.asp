<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/hitchhikerCls.asp"-->
<%
Dim hlist, page, i, isusing
page = request("page")
isusing = request("isusing")
If page = "" Then page = 1
	Set hlist = new viphitchhker
		hlist.FIsusing = isusing
		hlist.FPageSize = 20
		hlist.FCurrPage = page
		hlist.fnhitchlist
%>
<script language="javascript">
function newReg(con){
	var pop_view = window.open('popup_reg.asp?idx='+con+'','popup_reg','width=500,height=450,scrollbars=no,resizable=no');
	pop_view.focus();
}
function gosubmit(page){
    frm.page.value=page;
	frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<form name="frm" action="index.asp" method="get">
<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
<input type="hidden" name="page">
<tr align="center" bgcolor="#FFFFFF" height="40" >
	<td width="80" bgcolor="#EEEEEE">검색 조건</td>
	<td align="left" valign="middle">
		사용여부 :
		<select name="isusing">
			<option value="">선택</option>
			<option value="Y" <%=chkiif(isusing = "Y","selected","")%>>Y</option>
			<option value="N" <%=chkiif(isusing = "N","selected","")%>>N</option>
		</select>&nbsp;&nbsp;&nbsp;
		<img src="/images/icon_search.gif" border="0" style="cursor:pointer" onClick="frm.submit();" onfocus="this.blur();">
	</td>
</tr>
</form>
</table>
<table width="100%" cellpadding="0" cellspacing="0" class="a">
<tr height="50">
	<td>
		Total Count : <b><%=hlist.FTotalCount%></b>
	</td>
	<td align="right">
		<input type = "button" class="button" onclick="javascript:location.href='/admin/hitchhiker/';" value="신청리스트로">
		<a href="javascript:newReg('')"><img src="/images/icon_new_registration.gif" border="0"></a>
	</td>
</tr>
</table>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#E6E6E6" height="25">
	<td>번호</td>
	<td>회차</td>
	<td>web 이벤트코드</td>
	<td>mobile 이벤트코드</td>
	<td>시작일</td>
	<td>종료일</td>
	<td>배송일</td>
	<td>등록일</td>
	<td>사용여부</td>
</tr>
<%
	If hlist.FResultCount <> 0 Then
		For i = 0 to hlist.FResultCount - 1
%>
<tr bgcolor="FFFFFF" height="25" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" onClick="newReg('<%=hlist.FhitchList(i).FIdx%>');" style="cursor:pointer;">
	<td align="center"><%=hlist.FhitchList(i).Fidx%></td>
	<td align="center">Vol.<%=hlist.FhitchList(i).FHvol%></td>
	<td align="center"><%=hlist.FhitchList(i).Fevt_code%></td>
	<td align="center"><%=hlist.FhitchList(i).Fmevt_code%></td>
	<td align="center"><%=hlist.FhitchList(i).Fstartdate%></td>
	<td align="center"><%=hlist.FhitchList(i).Fenddate%></td>
	<td align="center"><%=hlist.FhitchList(i).Fdelidate%></td>
	<td align="center"><%=hlist.FhitchList(i).Fregdate%></td>
	<td align="center"><%=hlist.FhitchList(i).Fisusing%></td>
</tr>
<%
		Next
%>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If hlist.HasPreScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= hlist.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + hlist.StartScrollPage to hlist.StartScrollPage + hlist.FScrollCount - 1 %>
			<% If (i > hlist.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(hlist.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If hlist.HasNextScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
<%
	Else
%>
<tr bgcolor="FFFFFF" height="25" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td align="center" colspan="8" height="100">데이터가 없습니다</td>
</tr>
<%
	End If
%>
<% Set hlist = nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->