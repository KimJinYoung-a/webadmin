<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/search/itemCls.asp" -->
<%
Dim page, oKeyList, rowNum, i, search, searchstring
Dim sSdate, sEdate, mode

page			= requestCheckvar(request("page"),10)
mode			= requestCheckvar(request("mode"),1)
search			= requestCheckvar(request("search"),11)
searchstring	= requestCheckvar(request("searchstring"),100)
sSdate			= request("iSD")
sEdate			= request("iED")


If page = "" Then page = 1

SET oKeyList = new cItemContent
	oKeyList.FCurrPage			= page
	oKeyList.FPageSize			= 20
	oKeyList.FRectSdate			= sSdate
	oKeyList.FRectEdate			= sEdate
	oKeyList.FRectMode			= mode
	oKeyList.FRectSearch		= search
	oKeyList.FRectSearchstring	= searchstring
	oKeyList.getKeyWordLogList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function goPage(pg){
	frm.page.value = pg;
	frm.submit();
}

//날짜 지정
function jsSetDate(n, m){
	document.frm.iSD.value = "";
	document.frm.iED.value = "";
	var date = new Date();
	if(n == 7 || n == 15){
		var start = new Date(Date.parse(date) - n * 1000 * 60 * 60 * 24);
		var today = new Date(Date.parse(date) - m * 1000 * 60 * 60 * 24);
	
		var yyyy = start.getFullYear();
		var mm = start.getMonth()+1;
		var dd = start.getDate();

		var t_yyyy = today.getFullYear();
		var t_mm = today.getMonth()+1;
		var t_dd = today.getDate();
	}else{
        var t_mm = date.getMonth() + 1;
        var t_dd = date.getDate();
        var t_yyyy = date.getFullYear();
 		if(n == 30){
        	var preDate = new Date(date.setMonth(t_mm - 1)); 
        }else{
        	var preDate = new Date(date.setMonth(t_mm - 3)); 
        }
        var mm = preDate.getMonth() ; 
        var dd = preDate.getDate();
        var yyyy = preDate.getFullYear();
	}
	if (t_mm <10){
		t_mm = "0"+t_mm;
	}
	if (mm <10){
		mm = "0"+mm;
	}
	if (dd <10){
		dd = "0"+dd;
	}
	if (t_dd <10){
		t_dd = "0"+t_dd;
	}
	document.frm.iSD.value = yyyy + "-" + mm + "-" + dd; 
	document.frm.iED.value = t_yyyy + "-" + t_mm + "-" + t_dd;
}
function pop_keywordLogDetail(v){
    document.location.href = "/admin/search/popkeywordLogDetail.asp?idx="+v;
}
</script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		변경일 : 
		<input id="iSD" name="iSD" value="<%=sSdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		<input id="iED" name="iED" value="<%=sEdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "iSD", trigger    : "iSD_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "iED", trigger    : "iED_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		&nbsp;
		<input type="button" value="최근7일" class="button" onClick="jsSetDate(7,0)">
		<input type="button" value="최근15일" class="button" onClick="jsSetDate(15,0)">
		&nbsp;
		변경구분 : 
		<select name="mode" class="select">
			<option value="">전체</option>
			<option value="I" <%= Chkiif(mode="I", "selected", "") %> >등록</option>
			<option value="U" <%= Chkiif(mode="U", "selected", "") %> >수정</option>
			<option value="D" <%= Chkiif(mode="D", "selected", "") %> >삭제</option>
		</select>
		<br /><br />
		검색어 :
		<select name="search" class="select">
			<option value="">전체</option>
			<option value="nextkeyword">변경 키워드</option>
			<option value="subject">변경 내용</option>
			<option value="username">변경자</option>
		</select>
		<input type="text" name="searchstring" size="50" class="text" value="<%= searchstring %>">
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<p />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30" bgcolor="#FFFFFF">
	<td colspan="11">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				검색결과 : <b><%= FormatNumber(oKeyList.FTotalCount,0) %></b>
				&nbsp;
				페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oKeyList.FTotalPage,0) %></b>
			</td>
			<td align="right">
				*목록에서 바로 변경 적용한 키워드 정보는 공란입니다.
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="40">
	<td width="80">번호</td>
	<td width="100">변경 구분</td>
	<td width="150">변경 키워드</td>
	<td>변경 내용</td>
	<td width="100">변경자</td>
	<td width="100">변경일</td>
</tr>
<%
rowNum = oKeyList.FTotalcount - (page -1) * 20
For i = 0 To oKeyList.FResultCount - 1
%>
<tr align="center" bgcolor="#FFFFFF" height="30" onclick="pop_keywordLogDetail('<%= oKeyList.FItemList(i).FIdx %>');" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
	<td><%= rowNum %></td>
	<td>
		<%
			Select Case oKeyList.FItemList(i).FMode
				Case "I"		response.write "등록"
				Case "U"		response.write "수정"
				Case "D"		response.write "삭제"
			End Select
		%>
	</td>
	<td><%= oKeyList.FItemList(i).FNextkeyword %></td>
	<td><%= oKeyList.FItemList(i).FSubject %></td>
	<td><%= oKeyList.FItemList(i).FUsername %></td>
	<td><%= LEFT(oKeyList.FItemList(i).FRegdate, 10) %></td>
</tr>
<%
	rowNum = rowNum - 1 
Next
%>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18" align="center">
	<% If oKeyList.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oKeyList.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + oKeyList.StartScrollPage To oKeyList.FScrollCount + oKeyList.StartScrollPage - 1 %>
		<% If i>oKeyList.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If oKeyList.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
</table>
<% SET oKeyList = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->