<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2019.10.16 한용민 생성
'	Description : Link 발송
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/rndSerial.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/LinkSendCls.asp"-->
<%
dim page, linkidx, title, isusing, i
	page = requestCheckVar(getNumeric(request("page")),10)
	linkidx = requestCheckVar(getNumeric(request("linkidx")),10)
	title = requestCheckVar(request("title"),128)
	isusing = requestCheckVar(request("isusing"),1)

if page="" then page=1
if isusing="" then isusing="Y"

dim oLink
set oLink = New CLinkSend
    oLink.FCurrPage = page
    oLink.FPageSize=20
    oLink.FRectlinkidx = linkidx
    oLink.FRecttitle = title
    oLink.FRectisusing = isusing
    oLink.GetLinkSend

%>
<script type='text/javascript'>

function popsendlist(linkidx){
	var popwin = window.open('/admin/sitemaster/link/LinkSend_reg.asp?linkidx='+linkidx,'addreg','width=1400,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function gotoPage(pg) {
    document.Listfrm.page.value=pg;
    document.Listfrm.submit();
}

function fnLinkURLCopy(link) {
	window.clipboardData.setData("Text", link);
	alert('링크가 복사되었습니다.\n원하시는 곳에 Ctrl+V 하시면됩니다.');
}

</script>

<!-- 검색폼 시작 -->
<form name="Listfrm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		* 링크번호 :
		<input type="text" name="linkidx" size="10" value="<%= linkidx %>">
        &nbsp;
		* 링크명 :
		<input type="text" name="title" size="25" value="<%= title %>">
		&nbsp;
        * 사용여부 : <% drawSelectBoxisusingYN "isusing",isusing,"" %>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" onclick="gotoPage('1');" value="검색">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td align="right">
        <input type="button" value="신규등록" onclick="popsendlist('')" class="button">
    </td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="18">
        검색결과 : <b><%= oLink.FTotalCount %></b>
        &nbsp;
        페이지 : <b><%= page %>/ <%= oLink.FTotalPage %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=60>링크번호</td>
	<td width=220>링크명</td>
	<td>외부노출링크</td>
	<td>실제링크</td>
	<td width=50>사용여부</td>
	<td width=60>클릭수</td>
	<td width=80>최종수정</td>
    <td width=40>비고</td>
</tr>
<% if oLink.FResultCount>0 then %>
<% for i=0 to oLink.FResultCount-1 %>
<tr bgcolor="<%=chkIIF(oLink.FItemList(i).fisusing="Y","#FFFFFF","#E0E0E0")%>">
	<td align="center"><%= oLink.FItemList(i).flinkidx %></td>
	<td align="left"><%= chrbyte(ReplaceBracket(oLink.FItemList(i).ftitle),30,"Y") %></td>
	<td align="left">
		http://www.10x10.co.kr/apps/Link/LinkSend.asp?key=<%= rdmSerialEnc(oLink.FItemList(i).flinkidx) %>
		<input type="button" value="링크복사" onclick="fnLinkURLCopy('http://www.10x10.co.kr/apps/Link/LinkSend.asp?key=<%= rdmSerialEnc(oLink.FItemList(i).flinkidx) %>')" class="button">
	</td>
	<td align="left"><%= ReplaceBracket(oLink.FItemList(i).flinkurl) %></td>
	<td align="center"><%= oLink.FItemList(i).fisusing %></td>
	<td align="center"><%= oLink.FItemList(i).fviewcount %></td>
	<td align="center">
        <%= left(oLink.FItemList(i).flastupdate,10) %>
        <br>
        <%= mid(oLink.FItemList(i).flastupdate,11,22) %>
        <% if oLink.FItemList(i).flastadminid <> "" then %>
            <br>(<%= oLink.FItemList(i).flastadminid %>)
        <% end if %>
    </td>
	<td align="center">
        <input type="button" value="수정" onclick="popsendlist('<%= oLink.FItemList(i).flinkidx %>')" class="button">
    </td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="9" align="center">
	<% if oLink.HasPreScroll then %>
		<a href="javascript:gotoPage(<%= oLink.StarScrollPage-1 %>)">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oLink.StarScrollPage to oLink.FScrollCount + oLink.StarScrollPage - 1 %>
		<% if i>oLink.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:gotoPage(<%= i %>)">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oLink.HasNextScroll then %>
		<a href="javascript:gotoPage(<%= i %>)">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="18" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
<% set oLink = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->