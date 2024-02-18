<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/potal/potalCls.asp"-->
<%
Dim mallid, mallName
Dim mode, NOTmakerid, sqlStr, makerid, orderby
Dim oMaker, page, itemidarr, isusingarr, isusing
mallid		= requestCheckvar(request("mallid"),32)
page		= requestCheckvar(request("page"), 20)
mode		= request("mode")
NOTmakerid	= requestCheckvar(request("NOTmakerid"), 32)
itemidarr	= request("itemidarr")
isusingarr	= request("isusingarr")
isusing		= requestCheckvar(request("isusing"),1)
makerid		= requestCheckvar(request("makerid"), 32)
orderby		= requestCheckvar(request("orderby"), 10)

Select Case mallid
	Case "ggshop"		mallName = "구글쇼핑"
	Case "naverEP"		mallName = "네이버EP"
	Case "daumEP"		mallName = "다음EP"
End Select

If page = "" Then page = 1

SET oMaker = new CPotal
	oMaker.FCurrPage			= page
	oMaker.FPageSize			= 20
	oMaker.FMakerId				= makerid
	oMaker.FRectIsusing			= isusing
	oMaker.FRectOrderby			= orderby
	oMaker.FRectMallGubun		= mallid
    oMaker.getPotalNotInMakeridList
%>
<script language='javascript'>
function goPage(pg){
    var frm = document.frmsearch;
    frm.page.value=pg;
	frm.submit();
}
function searchFrm(){
	var frm = document.frmsearch;
	frmsearch.submit();
}
</script>
<% If mallid = "ggshop" Then %>
<!-- #include virtual="/admin/etc/potal/inc_googleHead.asp" -->
<% ElseIf mallid = "naverEP" Then %>
<!-- #include virtual="/admin/etc/potal/inc_naverHead.asp" -->
<% End If %>
<!-- 검색 시작 -->
<form name="frmsearch" method="post" action="notinMakerid.asp" >
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mallid" value="<%= mallid %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
		    <td width="90%">Mall 구분 : <%= mallName %></td>
		    <td rowspan="4" width="10%"><input type="button" onclick="searchFrm();" class="button" value="검 색" style="width:50px;height:50px;"></td>
		</tr>
		<tr>
			<td >
				브랜드ID : <input type="text" class="text" name="makerid" value="<%=makerid%>" size="20"> <input type="button" class="button" value="ID검색" onclick="jsSearchBrandID(this.form.name,'makerid');" >&nbsp;
				판매여부 :
				<select name="isusing" class="select">
					<option value="">-Choice-</option>
					<option value="Y" <%= Chkiif(isusing = "Y", "selected", "") %> >판매</option>
					<option value="N" <%= Chkiif(isusing = "N", "selected", "") %> >판매안함</option>
				</select>
				&nbsp;
				정렬기준 :
				<select name="orderby" class="select">
					<option value="">-Choice-</option>
					<option value="lastupdate" <%= Chkiif(orderby = "lastupdate", "selected", "") %> >최종수정일</option>
					<option value="best" <%= Chkiif(orderby = "best", "selected", "") %> >베스트브랜드</option>
				</select>
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
	<td>판매여부</td>
</tr>
<% If oMaker.FResultCount > 0 Then %>
<% For i = 0 To oMaker.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="30" align="center" height="25">
	<td><%=oMaker.FItemList(i).FMallgubun%></td>
	<td><%=oMaker.FItemList(i).FMakerid%></td>
	<td><%=oMaker.FItemList(i).FRegdate%></td>
	<td><%=oMaker.FItemList(i).FRegid%></td>
	<td><%=oMaker.FItemList(i).FLastupdate%></td>
	<td><%=oMaker.FItemList(i).FUpdateid%></td>
	<td>
		<%
			If oMaker.FItemList(i).FIsusing = "Y" Then
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
	<% If oMaker.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oMaker.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + oMaker.StartScrollPage To oMaker.FScrollCount + oMaker.StartScrollPage - 1 %>
		<% If i>oMaker.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If oMaker.HasNextScroll Then %>
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
<% SET oMaker = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->