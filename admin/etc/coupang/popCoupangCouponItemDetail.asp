<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/ssg/ssgItemcls.asp"-->
<%
Dim midx, page, i
page		= request("page")
midx 		= request("midx")

If page = "" Then page = 1

If NOT isNumeric(midx) Then
	Response.Write "<script language=javascript>alert('잘못된 접근입니다.');window.close();</script>"
	dbget.close()	:	response.End
End If

Dim oSsg
Set oSsg = new Cssg
	oSsg.FCurrPage					= page
	oSsg.FPageSize					= 50
	oSsg.FRectMasterIdx				= midx
	oSsg.getssgMarginItemDetailList
%>
<link rel="stylesheet" href="/bct.css" type="text/css">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function popCateSelect(){
	$.ajax({
		url: "/admin/etc/ssg/act_CategorySelect.asp",

		cache: false,
		success: function(message) {
			$("#lyrCateAdd").empty().append(message).fadeIn();
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}

function jsAddItemID() {
	var frm = document.frm;

	if (frm.itemid.value == '') {
		alert('상품코드를 입력하세요.');
		return;
	}

	if (confirm('저장하시겠습니까?')) {
		frm.delIdx.value = '';
		frm.submit();
	}
}

function delItem(v)
{
	$("#delIdx").val(v);
	document.frm.submit();
}

function goPage(pg){
    //frm.page.value = pg;
    //frm.submit();
	location.href = '?page=' + pg + '&midx=<%= midx %>';
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="3" bgcolor="<%= adminColor("tabletop") %>">기간별 마진 리스트</td>
</tr>
</table>

<br />

<form name="frm" action="procSsgMargin.asp" methd="post" style="margin:0px;">
<input type="hidden" name="mode" value="itemDetail">
<input type="hidden" name="midx" value="<%= midx %>">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" id="delIdx" name="delIdx" value="">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" class="a">
		<tr>
			<td width="80%">
				상품ID :
				<textarea class="textarea" name="itemid" cols="16" rows="2"></textarea>
				<input type="button" value="저 장" onClick="jsAddItemID()">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="5">
		검색결과 : <b><%= FormatNumber(oSsg.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oSsg.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="100">IDX</td>
	<td>상품코드</td>
	<td width="100">관리</td>
</tr>
<% For i=0 to oSsg.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= oSsg.FItemList(i).Fidx %></td>
	<td><%= oSsg.FItemList(i).Fitemid %></td>
	<td><input type="button" class="button" value="삭제" onclick="delItem(<%= oSsg.FItemList(i).FIdx %>);"></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oSsg.HasPreScroll then %>
		<a href="javascript:goPage('<%= oSsg.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oSsg.StartScrollPage to oSsg.FScrollCount + oSsg.StartScrollPage - 1 %>
    		<% if i>oSsg.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oSsg.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<% Set oSsg = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
