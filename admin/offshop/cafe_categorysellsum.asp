<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 카테고리 통계
' History : 최초생성자모름
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/cafecategorycls.asp"-->

<%
dim shopid
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim fromDate, toDate
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)

if session("ssBctDiv")="201" then
	shopid = "cafe002"
elseif session("ssBctDiv")="301" then
	shopid = "cafe003"
else
	''shopid = "cafe001"
end if

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))


fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

dim ooffsell,i
set ooffsell = new CCafeCategorySell
ooffsell.FRectStartDay = fromDate
ooffsell.FRectEndDay = toDate
ooffsell.FRectShopID = shopid
ooffsell.GetCafeCategorySell

dim omioffsell
set omioffsell = new CCafeCategorySell
omioffsell.FRectStartDay = fromDate
omioffsell.FRectEndDay = toDate
omioffsell.FRectShopID = shopid
omioffsell.GetCafeCategoryMiMatch
%>
<script language=javascript>
function PopCategory(iitemid,iitemname){
	var popwin = window.open("","popcafecategory","width=640 height=580 scrollbars=yes");
	document.bufform.itemid.value = iitemid;
	document.bufform.itemname.value = iitemname;
	document.bufform.target = "popcafecategory";
	document.bufform.submit();

}

function PopCateEdit(){
	var popwin = window.open("/admin/offshop/popcafecategorymaster.asp");
}
</script>
<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
      <input type="hidden" name="showtype" value="showtype">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		샵선택
		<select name="shopid">
			<option value="cafe001" <% if shopid = "cafe001" then response.write "selected"%>>1층cafe</option>
			<option value="cafe002" <% if shopid = "cafe002" then response.write "selected"%>>Zoom</option>
			<option value="cafe003" <% if shopid = "cafe003" then response.write "selected"%>>College</option>
		</select>
		<td class="a" align="right">
			<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0">
		</td>
	</tr>
	</form>
</table>
<span class=a>* 카테고리별 매출</span>
<table width="500" border="0" cellspacing="1" cellpadding="3" bgcolor="#3d3d3d" class=a>
<tr bgcolor="#DDDDFF">
	<td width=100>카테고리</td>
	<td width=100>매출건수</td>
	<td >매출액</td>
	<td >점유율</td>
</tr>
<% for i=0 to ooffsell.FResultCount -1 %>
<tr bgcolor="#FFFFFF">
	<td ><%= ooffsell.FItemList(i).FCateName %></td>
	<td align=right><%= FormatNumber(ooffsell.FItemList(i).FSellCount,0) %></td>
	<td align=right><%= FormatNumber(ooffsell.FItemList(i).FSellSum,0) %></td>
	<td align=center>
	<% if ooffsell.FSumTotal<>0 then %>
	<%= Clng(ooffsell.FItemList(i).FSellSum/ooffsell.FSumTotal * 10000)/100 %> %
	<% end if %>
	</td>
</tr>
<% next %>
<tr bgcolor="#DDDDFF">
	<td width=100>총계</td>
	<td width=100><%= FormatNumber(ooffsell.FCountTotal,0) %></td>
	<td align=right><%= FormatNumber(ooffsell.FSumTotal,0) %></td>
	<td align=right></td>
</tr>
</table>
<br>
<table width="500" border="0" cellspacing="1" cellpadding="3" class=a>
<tr bgcolor="#FFFFFF">
	<td >* 매칭 된 카테고리</td>
	<td align=right ><input type=button value="카테고리관리" onClick="PopCateEdit()"></td>
</tr>
<table width="500" border="0" cellspacing="1" cellpadding="3" bgcolor="#3d3d3d" class=a>
<tr bgcolor="#DDDDFF">
	<td width=100>상품코드</td>
	<td width=100>상품명</td>
	<td >카테고리지정</td>
</tr>
<% for i=0 to omioffsell.FResultCount -1 %>
<tr bgcolor="#FFFFFF">
	<td ><%= omioffsell.FItemList(i).FItemId %></td>
	<td ><%= omioffsell.FItemList(i).FItemName %></td>
	<td align=center>
		<% if IsNull(omioffsell.FItemList(i).FCateName) or (omioffsell.FItemList(i).FCateName="") then %>
		<a href="javascript:PopCategory('<%= omioffsell.FItemList(i).FItemId %>','<%= replace(omioffsell.FItemList(i).FItemName,"&","||") %>');">--&gt;</a>
		<% else %>
		<a href="javascript:PopCategory('<%= omioffsell.FItemList(i).FItemId %>','<%= replace(omioffsell.FItemList(i).FItemName,"&","||") %>');"><%= omioffsell.FItemList(i).FCateName %></a>
		<% end if %>
	</td>
</tr>
<% next %>
</table>
<%
set ooffsell = Nothing
set omioffsell = Nothing
%>
<form name=bufform method=post action="popcafecategory.asp">
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemname" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->