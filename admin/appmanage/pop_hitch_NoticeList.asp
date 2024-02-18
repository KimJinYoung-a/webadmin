<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/appmanage/hitchhiker.asp" -->
<%
Dim hitch, page, i
page = request("page")
If page = "" Then page = 1

Set hitch = new Hitchhiker
	hitch.FPageSize = 10
	hitch.FCurrPage = page
	hitch.HitchNoticeList
%>
<script language="javascript">
function gosubmit(page){
    frm.page.value=page;
	frm.submit();
}
function noticewrite(mode,idx){
	var winImg;
	winImg = window.open('/admin/appmanage/pop_hitch_NoticeWrite.asp?mode='+mode+'&idx='+idx+'','popImg','width=600,height=320, status=yes');
	winImg.focus();
}
</script>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:2;">
<form name="frm" method="get">
<input type="hidden" name="page">
<tr>
	<td align="left">
		<input type= "button" value="등록" class="button" onclick="javascript:noticewrite('I','');">
	</td>
</tr>
</table>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="5%">idx</td>
	<td width="10%">Device</td>
	<td width="10%">시작일</td>
	<td width="10%">종료일</td>
	<td>내용</td>
	<td width="10%">등록일</td>
	<td width="10%">사용여부</td>
</tr>
<% For i = 0 to hitch.FResultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onclick="javascript:noticewrite('U','<%=hitch.FhitchList(i).FNidx%>');" style="cursor:pointer;">
	<td><%=hitch.FhitchList(i).FNidx%></td>
	<td>
	<%
		Select Case hitch.FhitchList(i).FNassigndevice
			Case "0" response.write "전체"
			Case "1" response.write "IOS"
			Case "2" response.write "Android"
		End Select
	%>
	</td>
	<td><%=hitch.FhitchList(i).FNstartdate%></td>
	<td><%=hitch.FhitchList(i).FNenddate%></td>
	<td><%=nl2br(hitch.FhitchList(i).FNcontents)%></td>
	<td><%=Left(hitch.FhitchList(i).FNregdate,10)%></td>
	<td><%=hitch.FhitchList(i).FNisusing%></td>
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If hitch.HasPreScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= hitch.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + hitch.StartScrollPage to hitch.StartScrollPage + hitch.FScrollCount - 1 %>
			<% If (i > hitch.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(hitch.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If hitch.HasNextScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
</table>
<% Set hitch = nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
