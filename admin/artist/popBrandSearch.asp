<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
	Dim searchKey, searchString
	Dim sqlStr, tmpName
	searchKey		= Trim(Request("searchKey"))
	searchString	= Trim(Request("searchString"))

	if searchKey="" then searchKey="userid"
%>
<script FOR="window" EVENT="onload" LANGUAGE="javascript">
frmItemSearch.searchKey.value="<%=searchKey%>";
frmItemSearch.searchString.focus()
</script>
<script language="javascript">
<!--
	function putBrandItem(id,nm)
	{
		opener.inputfrm.designerid.value=id;
		opener.document.all.designerName.innerText=nm;
		self.close();
	}
//-->
</script>
<table width="300" cellpadding="2" cellspacing="1" border="0" class="a">
<form name="frmItemSearch" method="post">
<tr>
	<td bgcolor="F0F0F0"><b>��ǰ �˻�</b></td>
</tr>
<tr>
	<td bgcolor="F8F8F8" align="center">
		<select name="searchKey">
			<option value="userid">�귣��ID</option>
			<option value="socname_kor">�귣��(�ѱ�)</option>
			<option value="socname">�귣��(����)</option>
		</select>
		<input type="text" name="searchString" size="15" value="<%=searchString%>">
		<input type="button" value="�˻�" onClick="frmItemSearch.submit()">
	</td>
</tr>
</form>
</table>
<table width="300" cellpadding="2" cellspacing="1" border="0" class="a">
<%
	'�˻�� ������ �˻�
	if Not(searchString="" or IsNull(searchString)) then
		sqlStr = "Select userid, socname, socname_kor " &_
				"From db_user.dbo.tbl_user_c " &_
				"Where " & searchKey & " like '%" & searchString & "%'"
		rsget.Open sqlStr,dbget,1

		if Not(rsget.EOF or rsget.BOF) then
%>
<table width="300" cellpadding="2" cellspacing="1" border="0" class="a">
<tr>
	<td align="center" align="center" bgcolor="#F8F8F8">ID</td>
	<td align="center" bgcolor="#F8F8F8">�귣��</td>
</tr>
<%
	Do Until rsget.EOF
		tmpName = db2html(rsget("socname")) & " (" & db2html(rsget("socname_kor")) & ")"
%>
<tr>
	<td style="border-bottom:1px solid #F0F0F0"><%=db2html(rsget("userid"))%></td>
	<td style="border-bottom:1px solid #F0F0F0"><a href="javascript:putBrandItem('<%=rsget("userid")%>','<%=Replace(tmpName,"'","\'")%>');"><%=tmpName%></a></td>
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
<tr><td align="center" height="180">"<%=SearchString%>"(��)�� �˻��� �귣�尡 �����ϴ�!</td></tr>
</table>
<%
		end if
		rsget.Close
%>
<%	else %>
<table width="300" cellpadding="2" cellspacing="1" border="0" class="a">
<tr><td align="center" height="180">��ǰ�� �˻����ּ���.</td></tr>
</table>
<%	end if %>
<table width="300" cellpadding="2" cellspacing="1" border="0" class="a">
<tr>
	<td align="center" bgcolor="#F0F0F0"><input type="button" value=" �ݱ� " onClick="self.close()"></td>
</tr>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
