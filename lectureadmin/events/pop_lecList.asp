<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/event/eventCls.asp"-->
<%
Dim i, page, searchKey, searchString
Dim oLec, lecturer_id, research
research		= requestCheckvar(request("research"),10)
page			= requestCheckvar(request("page"),10)
searchKey		= requestcheckvar(Request("searchKey"),10)
searchString	= Request("searchString")
lecturer_id		= session("ssBctId")
if searchString <> "" then
	if checkNotValidHTML(searchString) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "</script>"
	response.End
	end if
end if
If page = "" Then page = 1

Set oLec = new CEvent
	oLec.FCurrPage			= page
	oLec.FPageSize			= 12
	oLec.FRectSearchKey		= searchKey
	oLec.FRectSearchString	= searchString
	oLec.FRectlecturerID	= lecturer_id
	oLec.getLecList
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function openerRegLecIdx(v){
	opener.$("#lecidx").val(v);
	window.close();
}
</script>
</head>
<body>
<table width="100%" align="center" cellpadding="3" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" action="" method="POST">
<input type="hidden" name="page">
<tr height="80" bgcolor="FFFFFF">
	<td>
		<font size="4"><strong>진행중인 강좌</strong></font>&nbsp;&nbsp;
		총 <%= oLec.FTotalCount %>건
	</td>
	<td align="right">
		<select name="searchKey" class="select">
			<option value="lectitle" <%= chkiif(searchKey = "lectitle", "selected", "") %>>강좌명</option>
			<option value="lecidx" <%= chkiif(searchKey = "lecidx", "selected", "") %>>강좌코드</option>
		</select>
		<input type="text" class="text" name="searchString" value="<%=searchString%>">
		<input type="button" class="button" value="검색" onclick="document.frm.submit();">
	</td>
</tr>
</form>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">강좌 코드</td>
	<td>강좌명</td>
	<td width="300">일정</td>
	<td width="100">잔여인원</td>
</tr>
<% For i = 0 to oLec.FResultCount - 1 %>
<tr height="30" bgcolor="FFFFFF" align="center" style="cursor:pointer;" onclick="openerRegLecIdx('<%= oLec.FItemList(i).FLecIdx %>')" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background='ffffff';>
	<td><%= oLec.FItemList(i).FLecIdx %></td>
	<td><%= oLec.FItemList(i).FLecTitle %></td>
	<td><%= oLec.FItemList(i).FLecperiod %> <% If oLec.FItemList(i).FLecLimitCount <> 0 Then %> (한정<%= oLec.FItemList(i).FLecLimitCount %>명)<% End If %></td>
	<td><%= oLec.FItemList(i).FOptLimitCnt %> / <%=oLec.FItemList(i).FLecLimitCount/oLec.FItemList(i).FOptionCnt%></td>	
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oLec.HasPreScroll then %>
		<a href="javascript:goPage('<%= oLec.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oLec.StartScrollPage to oLec.FScrollCount + oLec.StartScrollPage - 1 %>
    		<% if i>oLec.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oLec.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</body>
</html>
<% Set oLec = nothing %>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->