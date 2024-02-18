<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/board/upche_qnacls.asp" -->
<%
dim i, ix
dim page
Dim gubun
page = request("page")
gubun = request("gubun")
if (page = "") then
    page = "1"
end if
%>
<script language="JavaScript">
<!--
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
//-->
</script>

<%
If (session("ssBctDiv") < 10) Then
dim itemqanotinclude, research
dim boardqna
set boardqna = New CUpcheQnA
    boardqna.FPageSize = 200
    boardqna.FCurrPage = page
    boardqna.FRectRelpy = "N"
    boardqna.FRectGubun = gubun
    boardqna.list
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		구분 :
		<select name="gubun" class="select">
            <option value="">-전체-
			<option value="01" <%= CHkIIF(gubun="01","selected","") %> >배송문의
			<option value="02" <%= CHkIIF(gubun="02","selected","") %> >반품문의
			<option value="03" <%= CHkIIF(gubun="03","selected","") %> >교환문의
            <option value="04" <%= CHkIIF(gubun="04","selected","") %> >정산문의
            <option value="05" <%= CHkIIF(gubun="05","selected","") %> >입고문의
            <option value="06" <%= CHkIIF(gubun="06","selected","") %> >재고문의
            <option value="07" <%= CHkIIF(gubun="07","selected","") %> >상품등록문의
            <option value="08" <%= CHkIIF(gubun="08","selected","") %> >이벤트시행문의
            <option value="20" <%= CHkIIF(gubun="20","selected","") %> >기타문의
		</select>&nbsp;
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<br /><br />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="5">
		검색결과 : <b><%= FormatNumber(boardqna.FTotalCount,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20%">업체명</td>
	<td>제목</td>
	<td width="10%">구분</td>
	<td width="10%">업체구분</td>
	<td width="10%">작성일</td>
</tr>
<% For i = 0 to (boardqna.FResultCount - 1) %>
<tr align="center" bgcolor= "#FFFFFF">
    <td align="center">&nbsp;<%= boardqna.FItemList(i).Fusername %>(<%= boardqna.FItemList(i).Fuserid %>)</td>
    <td>&nbsp;<a href="/admin/board/upche_qna_board_reply.asp?idx=<%= boardqna.FItemList(i).Fidx %>"><%= (boardqna.FItemList(i).Ftitle) %></a>
    <% if datediff("d",boardqna.FItemList(i).Fregdate,now())<6 then %>
	&nbsp;&nbsp;&nbsp;<img src="/images/new.gif">
	<% end if %>
    </td>
    <td align="center"><%= boardqna.FItemList(i).GubunName %></td>
    <td align="center"><%= boardqna.FItemList(i).UpcheGubun %></td>
    <td align="center"><%= FormatDate(boardqna.FItemList(i).Fregdate, "0000.00.00") %></td>
<% Next %>
<tr height="20">
    <td colspan="17" align="center" bgcolor="#FFFFFF">
        <% if boardqna.HasPreScroll then %>
		<a href="javascript:goPage('<%= boardqna.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + boardqna.StartScrollPage to boardqna.FScrollCount + boardqna.StartScrollPage - 1 %>
    		<% if i>boardqna.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if boardqna.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<% set boardqna = Nothing %>
<% End If %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->