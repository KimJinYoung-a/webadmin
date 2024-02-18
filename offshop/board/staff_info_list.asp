<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionoffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/lib/classes/offshop/offshop_staffcls.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<%

dim i, j, page, shopid, isusing, research

shopid = session("ssBctId")
isusing= request("isusing")
research= request("research")
page = request("page")
if page="" then page=1
if (research="") and (isusing="") then isusing="Y"

dim nstaff
set nstaff = New COffshopStaff

nstaff.FRectIsusing = isusing
nstaff.FRectShopid = shopid
nstaff.FPageSize = 20
nstaff.FCurrPage = page
nstaff.FScrollCount = 10
nstaff.GetOffshopStaffList

%>
<STYLE TYPE="text/css">
<!--
    A:link, A:visited, A:active { text-decoration: none; }
    A:hover { text-decoration:underline; }
    BODY, TD, UL, OL, PRE { font-size: 9pt; }
    INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #CACACA; color: #000000; }
-->
</STYLE>
<script language='javascript'>
function  TnSearch(frm){
	if (frm.rectuserid.length<1){
		alert('검색어를 입력하세요.');
		return;
	}
	frm.method="get";
	frm.submit();
}
function NextPage(ipage){
	document.frmSrc.page.value= ipage;
	document.frmSrc.submit();
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frmSrc" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
			사용구분 :
			<select name="isusing" class="select" >
			    <option value="">ALL
			    <option value="Y" <%= chkIIF(isusing="Y","selected","") %> >Y
			    <option value="N" <%= chkIIF(isusing="N","selected","") %> >N
			</select>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frmSrc.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table border="1" bordercolordark="White" bordercolorlight="black" cellpadding="0" cellspacing="0">
  <tr bgcolor="#DDDDFF" height="25">
    <td width="50" align="center">번호</td>
    <td width="100" align="center">샵</td>
   <td width="100" align="center">이름</td>
   <td width="100" align="center">사진</td>
    <td width="100" align="center">입사일자</td>
    <td width="50" align="center">사용유무</td>
	<td width="100" align="center">작성일</td>
  </tr>
<% for i = 0 to (nstaff.FResultCount - 1) %>
  <tr height="20">
    <td align="center"><%= nstaff.FItemList(i).Fidx %></td>
    <td align="center"><%= nstaff.FItemList(i).Fshopname %></td>
    <td align="center"><a href="staff_info_write.asp?idx=<%= nstaff.FItemList(i).Fidx %>&mode=edit&menupos=<%= menupos %>"><%= nstaff.FItemList(i).Fusername %></a></td>
    <td align="center"><img src="<% = nstaff.FItemList(i).Ficon1 %>" width="50" height="60" border="0"></td>
	<td align="center"><%= FormatDate(nstaff.FItemList(i).Fipsadate, "0000.00.00") %></td>
	<td align="center"><%= nstaff.FItemList(i).Fisusing %></td>
	<td align="center"><%= FormatDate(nstaff.FItemList(i).Fregdate, "0000.00.00") %></td>
  </tr>
<% next %>
</table>
<table width="500" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td align="center" height="30">
		<% if nstaff.HasPreScroll then %>
			<a href="javascript:NextPage('<%= nstaff.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + nstaff.StartScrollPage to nstaff.FScrollCount + nstaff.StartScrollPage - 1 %>
			<% if i>nstaff.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if nstaff.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<tr>
	<td align="right"><a href="staff_info_write.asp?mode=add&menupos=<%= menupos %>"><font color="red">Staff 등록</font></a>&nbsp;&nbsp;&nbsp;</td>
</tr>
</table>
<br><br>
<% set nstaff = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->