<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
	Dim searchKey, searchString
	Dim sqlStr
	searchKey		= Trim(Request("searchKey"))
	searchString	= Trim(Request("searchString"))

	if searchKey="" then searchKey="itemname"
%>
<script FOR="window" EVENT="onload" LANGUAGE="javascript">
frmItemSearch.searchKey.value="<%=searchKey%>";
frmItemSearch.searchString.focus()
</script>
<script language="javascript">
<!--
	function putItemItem(id,nm)
	{
		opener.inputfrm.itemid.value=id;
		opener.document.all.itemname.innerText=nm;
		self.close();
	}
//-->
</script>
<table width="300" cellpadding="2" cellspacing="1" border="0" class="a">
<form name="frmItemSearch" method="post">
<tr>
	<td bgcolor="F0F0F0"><b>상품 검색</b></td>
</tr>
<tr>
	<td bgcolor="F8F8F8" align="center">
		<select name="searchKey">
			<option value="itemid">상품번호</option>
			<option value="itemname">상품명</option>
			<option value="socname_kor">브랜드(한글)</option>
			<option value="socname">브랜드(영문)</option>
		</select>
		<input type="text" name="searchString" size="15" value="<%=searchString%>">
		<input type="button" value="검색" onClick="frmItemSearch.submit()">
	</td>
</tr>
</form>
</table>
<table width="300" cellpadding="2" cellspacing="1" border="0" class="a">
<%
	'검색어가 있으면 검색
	if Not(searchString="" or IsNull(searchString)) then
		sqlStr = "Select t1.itemid, t1.itemname, t1.brandname " &_
				"From db_item.[dbo].tbl_item as t1 " &_
				"	Join db_user.[dbo].tbl_user_c as t2 " &_
				"		on t1.makerid=t2.userid " &_
				"Where t1.isusing='Y' and t2.isusing='Y' " &_
				"	and " & searchKey & " like '%" & searchString & "%'"
		rsget.Open sqlStr,dbget,1

		if Not(rsget.EOF or rsget.BOF) then
%>
<table width="300" cellpadding="2" cellspacing="1" border="0" class="a">
<tr>
	<td align="center" align="center" bgcolor="#F8F8F8">브랜드</td>
	<td align="center" bgcolor="#F8F8F8">상품</td>
</tr>
<%	Do Until rsget.EOF %>
<tr>
	<td style="border-bottom:1px solid #F0F0F0"><%=db2html(rsget("brandname"))%></td>
	<td style="border-bottom:1px solid #F0F0F0"><a href="javascript:putItemItem('<%=rsget("itemid")%>','<%=db2html(rsget("itemname"))%>');"><%=db2html(rsget("itemname"))%></a></td>
</tr>
<%
		rsget.MoveNext
	Loop
%>
</table>
<%
		else
%>
<table width="300" cellpadding="2" cellspacing="1" border="0" class="a">
<tr><td align="center" height="180">"<%=SearchString%>"(으)로 검색된 상품이 없습니다!</td></tr>
</table>
<%
		end if
		rsget.Close
%>
<%	else %>
<table width="300" cellpadding="2" cellspacing="1" border="0" class="a">
<tr><td align="center" height="180">상품을 검색해주세요.</td></tr>
</table>
<%	end if %>
<table width="300" cellpadding="2" cellspacing="1" border="0" class="a">
<tr>
	<td align="center" bgcolor="#F0F0F0"><input type="button" value=" 닫기 " onClick="self.close()"></td>
</tr>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
