<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/daumEp/epShopCls.asp"-->
<%
Dim mode, NOTmakerid, sqlStr, makerid
Dim nMaker, page, itemidarr, isusingarr
page		= request("page")
mode		= requestCheckvar(request("mode"),1)
NOTmakerid	= requestCheckvar(request("NOTmakerid"),32)
itemidarr	= request("itemidarr")
isusingarr	= request("isusingarr")
makerid		= requestCheckvar(request("makerid"),32)

If page = "" Then page = 1

SET nMaker = new epShop
	nMaker.FCurrPage					= page
	nMaker.FPageSize					= 20
	nMaker.FMakerId						= makerid
    nMaker.EpshopnotinmakeridList
%>
<script language='javascript'>
function goPage(pg){
    var frm = document.frmsearch;
    frm.page.value=pg;
	frm.submit();
}
</script>
<!-- #include virtual="/admin/etc/daumEp/inc_daumHead.asp" -->
<!-- 검색 시작 -->
<form name="frmsearch" method="post" action="notinmakerid.asp" >
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
		    <td width="90%">Mall 구분 : 다음EP</td>
		    <td rowspan="4" width="10%"><input type="submit" value="검 색" style="width:50px;height:50px;"></td>
		</tr>
		<tr>
			<td >
			브랜드ID : <input type="text" class="text" name="makerid" value="<%=makerid%>" size="20"> <input type="button" class="button" value="ID검색" onclick="jsSearchBrandID(this.form.name,'makerid');" >
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fitem" method="post" style="margin:0px;">
<input type="hidden" name="sortarr" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
    <td>몰구분</td>
	<td>브랜드ID</td>
	<td>등록일</td>
	<td>등록자</td>
	<td>최종수정일</td>
	<td>최종수정자</td>
	<td>판매설정</td>
</tr>
<% If nMaker.FResultCount > 0 Then %>
<% For i = 0 To nMaker.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="30" align="center" height="25">
	<td><%=nMaker.FItemList(i).FMallgubun%></td>
	<td><%=nMaker.FItemList(i).FMakerid%></td>
	<td><%=nMaker.FItemList(i).FRegdate%></td>
	<td><%=nMaker.FItemList(i).FRegid%></td>
	<td><%=nMaker.FItemList(i).FLastupdate%></td>
	<td><%=nMaker.FItemList(i).FUpdateid%></td>
	<td>
		<%
			If nMaker.FItemList(i).FIsusing = "Y" Then
				response.write "판매함"
			Else
				response.write "판매안함"
			End If
		%>
	</td>
</tr>
<% Next %>
<tr height="30">
	<td colspan="16" align="center" bgcolor="#FFFFFF">
	<% If nMaker.HasPreScroll Then %>
		<a href="javascript:goPage('<%= nMaker.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + nMaker.StartScrollPage To nMaker.FScrollCount + nMaker.StartScrollPage - 1 %>
		<% If i>nMaker.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If nMaker.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
<% Else %>
<tr height="50">
	<td colspan="16" align="center" bgcolor="#FFFFFF">
		등록된 브랜드가 없습니다
	</td>
</tr>
<% End If %>
</form>
</table>
<% SET nMaker = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->