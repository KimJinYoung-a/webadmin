<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<%
response.write "사용중지"
dbget.close()	:	response.End
dim designer,page,ckonlyoff,ckonlyusing
dim research,pricediff,imageview

designer = request("designer")
page  = request("page")
ckonlyoff = request("ckonlyoff")
ckonlyusing = request("ckonlyusing")
research = request("research")
pricediff = request("pricediff")
imageview = request("imageview")

if page="" then page=1
if research<>"on" then ckonlyusing="on"

dim ioffitem
set ioffitem  = new COffShopItem
ioffitem.FPageSize = 50
ioffitem.FCurrPage = page
ioffitem.FRectDesigner = designer
ioffitem.FRectOnlyOffLine = ckonlyoff
ioffitem.FRectOnlyUsing = ckonlyusing

if pricediff="on" then
	ioffitem.GetOffShopPriceDiffItemList
else
	if designer<>"" then
		ioffitem.GetOffNOnLineShopItemList
	end if
end if

dim i
%>
<script language='javascript'>
function Pop_barcode_input
	window.open('/admin/itemmaster/pop_barcode_input.asp?barcode='+barcode ,'width=500,height=300,scrollbars=yes,status=no');
}

</script>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<tr>
		<td class="a" >
		<input type="checkbox" name="imageview" value="on" <% if imageview="on" then response.write "checked" %> >이미지보기
		&nbsp;&nbsp;&nbsp;
		<input type="checkbox">사용중인상품만
		&nbsp;&nbsp;&nbsp;
		<input type="checkbox">온라인상품(10)
		&nbsp;&nbsp;
		<input type="checkbox">오프라인전용상품(90)
		
		<br>
		
		브랜드 : <% drawSelectBoxDesignerwithName "designer",designer  %>
		&nbsp;&nbsp;
		상품-Barcode : <input type="text" name="barcode" value="" size="9" maxlength="32" style="border:1px #999999 solid; ">
		</td>
			
		<td class="a" align="right">
			<a href="javascript:location.reload();"><img src="/admin/images/icon_reload.gif" width="24" height="20" border="0" alt="새로고침"></a>
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="100%" cellspacing="1" class="a" bgcolor=#3d3d3d>
	<tr bgcolor="#FFFFFF">
		<td colspan="12" align="right"><input type="button" value="선택아이템저장" onclick="ModiArr()"></td>
	</tr>
	<tr bgcolor="#DDDDFF" align="center">
    		<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
    		<td width="70">브랜드ID</td>
    		<td width="80">상품code</td>
    		<td >상품명</td>
    		<td width="150">옵션명</td>
    		<td width="60">판매가</td>
    		<td width="100">업체관리코드</td>
    		<td width="100">범용바코드</td>

	</tr>
	<% for i=0 to ioffitem.FresultCount -1 %>
	<tr bgcolor="#FFFFFF"  align="center">
	
    		<td><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
    		<td align="left"><%= ioffitem.FItemlist(i).FMakerID %></td>
    		<td><a href="javascript:Pop_barcode_input('<%= ioffitem.FItemlist(i).GetBarCode %>')"><%= ioffitem.FItemlist(i).GetBarCode %></a></td>
    		<td align="left"><%= ioffitem.FItemlist(i).FShopItemName %></td>
    		<td align="left"><%= ioffitem.FItemlist(i).FShopitemOptionname %></td>
    		<td align="right"><%= ioffitem.FItemlist(i).FOnLineItemprice %></td>
    		<td><input type="text" name="extbarcode" value="<%= ioffitem.FItemlist(i).FextBarcode %>" size="9" maxlength="32" style="border:1px #999999 solid; "></td>
    		<td><input type="text" name="extbarcode" value="<%= ioffitem.FItemlist(i).FextBarcode %>" size="9" maxlength="32" style="border:1px #999999 solid; "></td>
	
	</tr>
	<% next %>
	
	
<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->