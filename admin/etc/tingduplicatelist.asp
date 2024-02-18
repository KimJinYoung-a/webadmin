<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/tingcls.asp"-->
<%
dim yyyy1,mm1

yyyy1 = request("yyyy1")
mm1   = request("mm1")

if (yyyy1="") then
	yyyy1 = Left(Cstr(Now()),4)
	mm1   = Mid(Cstr(Now()),6,2)
end if

dim oting
set oting = new CTingItemList
oting.FRectYYYY = yyyy1
oting.FRectMM   = mm1
oting.DuplicateListByID

dim i
%>

<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
		<% DrawYMBox yyyy1,mm1  %>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<div class="a">userid�ߺ�</div>
<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
      <td width="20" >ID</td>
      <td width="120">�ֹ���ȣ</td>
      <td width="60">��ۻ���</td>
      <td width="60">Userid</td>
      <td width="40">������</td>
      <td width="50">������</td>
      <td width="50">������</td>
      <td width="40">TingQ</td>
      <td width="80">�ֹ���</td>

    </tr>
    <% for i=0 to oting.FResultCount-1 %>
    <tr bgcolor="#FFFFFF">
      <td><%= oting.FTingList(i).FID %></td>
      <td ><%= oting.FTingList(i).FOrderSerial %></td>
      <td ><font color="<%= IpkumDivColor(oting.FTingList(i).FIpkumdiv) %>"><%= IpkumDivName(oting.FTingList(i).FIpkumdiv) %></font></td>
      <td ><%= oting.FTingList(i).FUserID %></td>
      <td><%= oting.FTingList(i).FBuyName %></td>
      <td><%= oting.FTingList(i).FReqName %></td>
      <td><%= oting.FTingList(i).FItemName %></td>
      <td align="right"><%= FormatNumber(oting.FTingList(i).FTingQ,0) %></td>
      <td ><%= oting.FTingList(i).FOrderdate %></td>

    </tr>
    <% next %>
</table>

<%
'oting.DuplicateListByEmail
oting.DuplicateListByReqHp
%>
<div class="a">email�ߺ�</div>
<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
      <td width="20" >ID</td>
      <td width="120">�ֹ���ȣ</td>
      <td width="60">��ۻ���</td>
      <td width="60">Userid</td>
      <td width="40">������</td>
      <td width="50">������</td>
      <td width="50">������</td>
      <td width="40">TingQ</td>
      <td width="80">�ֹ���</td>
      <td width="80">�̸���</td>
      <td width="80">������HP</td>
      <td width="100">�ּ�</td>
    </tr>
    <% for i=0 to oting.FResultCount-1 %>
    <tr bgcolor="#FFFFFF">
      <td><%= oting.FTingList(i).FID %></td>
      <td ><%= oting.FTingList(i).FOrderSerial %></td>
      <td ><font color="<%= IpkumDivColor(oting.FTingList(i).FIpkumdiv) %>"><%= IpkumDivName(oting.FTingList(i).FIpkumdiv) %></font></td>
      <td ><%= oting.FTingList(i).FUserID %></td>
      <td><%= oting.FTingList(i).FBuyName %></td>
      <td><%= oting.FTingList(i).FReqName %></td>
      <td><%= oting.FTingList(i).FItemName %></td>
      <td align="right"><%= FormatNumber(oting.FTingList(i).FTingQ,0) %></td>
      <td ><%= oting.FTingList(i).FOrderdate %></td>
      <td ><%= oting.FTingList(i).FBuyEmail %></td>
      <td ><%= oting.FTingList(i).FReqHp %></td>
      <td ><%= oting.FTingList(i).FReqaddr1 + " " + oting.FTingList(i).FReqaddr2 %></td>
    </tr>
    <% next %>
</table>
<%
set oting = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->