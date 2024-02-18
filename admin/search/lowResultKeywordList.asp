<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbEVTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/seachkeywordCls.asp" -->
<%

dim i
dim research
dim yyyy1, mm1, dd1, currDate
dim mxrectCNT, searchcnt

research 	= requestCheckvar(request("research"),32)
yyyy1 		= requestCheckvar(request("yyyy1"),32)
mm1 		= requestCheckvar(request("mm1"),32)
dd1 		= requestCheckvar(request("dd1"),32)

mxrectCNT 	= requestCheckvar(request("mxrectCNT"),32)
searchcnt 	= requestCheckvar(request("searchcnt"),32)

if (mxrectCNT = "") then
	mxrectCNT = 10
end if

if (searchcnt = "") then
	searchcnt = 10
end if

if (yyyy1 = "") then
	currDate = Now()
	yyyy1 = Year(currDate)
	mm1 = Month(currDate)
	dd1 = Day(currDate)
else
	currDate = DateSerial(yyyy1, mm1, dd1)
end if

'// ============================================================================
dim osearchKeyword

set osearchKeyword = new CSearchKeyword

osearchKeyword.FRectYYYYMMDD	= Left(currDate,10)
osearchKeyword.FRectMxrectCNT	= mxrectCNT
osearchKeyword.FRectSearchCNT	= searchcnt

osearchKeyword.GetLowResultKeywordList

%>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left" height="30" >
			* ��¥ : <% Call DrawOneDateBoxdynamic("yyyy1", yyyy1, "mm1", mm1, "dd1", dd1, "", "", "", "") %> ~
			&nbsp;
			* �˻�Ƚ�� : <input type="text" class="text" name="searchcnt" size="4" value="<%= searchcnt %>"> �� �̻�
			&nbsp;
			* ��հ˻������ : <input type="text" class="text" name="mxrectCNT" size="4" value="<%= mxrectCNT %>"> �� ����
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p />

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="21">
			�˻���� : <b><%= osearchKeyword.FResultCount %></b>
		</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="150">�˻���</td>
		<td width="80">�˻�Ƚ��</td>
		<td width="80">���<br />�˻������</td>
		<td>���</td>
	</tr>
	<%
	for i = 0 To osearchKeyword.FResultCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" height="30">
			<%= osearchKeyword.FItemList(i).Frect %>
		</td>
		<td align="center"><%= osearchKeyword.FItemList(i).Fsumsearchcnt %></td>
		<td align="center"><%= osearchKeyword.FItemList(i).FmxrectCNT %></td>
		<td></td>
	</tr>
	<%
	next
	%>
	<% if (osearchKeyword.FResultCount = 0) then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td height="30" colspan="15">
			�˻������ �����ϴ�.
		</td>
	</tr>
	<% end if %>
</table>
<%
set osearchKeyword = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbEVTclose.asp" -->
