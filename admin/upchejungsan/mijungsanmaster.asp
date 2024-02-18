<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->

<%
dim yyyy1,mm1
yyyy1 = request("yyyy1")
mm1 = request("mm1")

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FRectYYYYMM = yyyy1 + "-" + mm1
ojungsan.GetMijungsanList
dim i
%>
<table width="760" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
		정산대상년월:<% DrawYMBox yyyy1,mm1 %>
		&nbsp;&nbsp;
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<div class="a">정산 미처리 리스트 (총 <%= ojungsan.FResultCount %>건)</div>

<table width="760" cellspacing="1"  class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
      <td width="80">디자이너ID</td>
      <td width="160">디자이너</td>
      <td width="100">업체배송총예상액</td>
      <td width="100">매입총예상액</td>
      <td width="100">위탁총예상액</td>
      <td width="100">총정산예상액</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <tr bgcolor="#FFFFFF">
      <td ><%= ojungsan.FItemList(i).Fdesignerid %></td>
      <td ><a href="mijungsanlist.asp?designer=<%= ojungsan.FItemList(i).Fdesignerid %>"><%= ojungsan.FItemList(i).Ftitle %></a></td>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).Fub_totalsuplycash,0) %></td>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).Fme_totalsuplycash,0) %></td>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).Fwi_totalsuplycash,0) %></td>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).Fwi_totalsuplycash,0) %></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td width="80">총계</td>
      <td width="160"></td>
      <td width="100"></td>
      <td width="100"></td>
      <td width="100"></td>
      <td width="100"></td>
    </tr>
</table>
<%
set ojungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
