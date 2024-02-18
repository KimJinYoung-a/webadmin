<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/shopmaster/mileageshopitemcls.asp"-->
<%
dim sqlstr
dim delitem, itemidarr
delitem = request("delitem")
itemidarr = request("itemidarr")

if Right(itemidarr,1)="," then
	itemidarr = Left(itemidarr,Len(itemidarr)-1)
end if

if delitem<>"" then
	sqlstr = " update [db_item].[dbo].tbl_item" + VbCrlf
	sqlstr = sqlstr + " set itemdiv='01'" + VbCrlf
	sqlstr = sqlstr + " where itemid=" + CStr(delitem) + VbCrlf

	dbget.Execute sqlStr
	
	sqlstr = " delete from db_item.dbo.tbl_mileageshop_item" + VbCrlf
	sqlstr = sqlstr + " where mileageShopItemID in (" + VbCrlf
	sqlstr = sqlstr + " 	select m.mileageShopItemID" + VbCrlf
	sqlstr = sqlstr + " 	from db_item.dbo.tbl_mileageshop_item m" + VbCrlf
	sqlstr = sqlstr + " 	left join [db_item].[dbo].tbl_item i" + VbCrlf
	sqlstr = sqlstr + " 	on m.mileageShopItemID= i.itemid and i.itemdiv='82' " + VbCrlf
	sqlstr = sqlstr + " 	where i.itemid is NULL" + VbCrlf
	sqlstr = sqlstr + " )" + VbCrlf
	
	dbget.Execute sqlStr
end if

if itemidarr<>"" then
	sqlstr = " update [db_item].[dbo].tbl_item" + VbCrlf
	sqlstr = sqlstr + " set itemdiv='82'" + VbCrlf
	sqlstr = sqlstr + " where itemid in (" + CStr(itemidarr) + ")"

	dbget.Execute sqlStr
	
	sqlstr = " insert into db_item.dbo.tbl_mileageshop_item" + VbCrlf
	sqlstr = sqlstr + " (mileageShopItemID)" + VbCrlf
	sqlstr = sqlstr + " select i.itemid" + VbCrlf
	sqlstr = sqlstr + " from  [db_item].[dbo].tbl_item i" + VbCrlf
	sqlstr = sqlstr + " 	left join db_item.dbo.tbl_mileageshop_item m" + VbCrlf
	sqlstr = sqlstr + " 	on i.itemid=m.mileageShopItemID" + VbCrlf
	sqlstr = sqlstr + " where i.itemdiv='82' " + VbCrlf
	sqlstr = sqlstr + " and m.mileageShopItemID is NULL" + VbCrlf
	
	dbget.Execute sqlStr
end if

dim omileageshop
set omileageshop = new CMileageShop
omileageshop.FPageSize=100
omileageshop.GetMileageShopItemList


dim iCols, iRows
iCols=4
iRows = Clng(omileageshop.FResultCount \ iCols)
if (omileageshop.FResultCount mod iCols)>0 then
	iRows = iRows + 1
end if

dim i,j
%>
<script language='javascript'>
function delThis(v){
	if (confirm('일반상품으로 변경 하시겠습니까?')){
		document.location = "?delitem="+ v;
	}
}

function SaveArr(frm){
	if (confirm('마일리지샾 상품으로 변경 하시겠습니까?')){
		frm.submit();
	}
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC" class=a>
	<form name="frm" method="post" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<tr>
		<td><input type=text name=itemidarr size="30" maxlength=64>
		(, 콤머로 구분)
		<input type=button value="마일리지 상품으로 변경" onClick="SaveArr(frm)"></td>
		<td class="a" align="right">
		</td>
	</tr>
	</form>
</table>
<table border="0" width="750" cellpadding="0" cellspacing="0">
<tr>
	<td>
<% if omileageshop.FResultCount<1 then %>
   [ 검색 결과가 없습니다. ]
<% else %>
  <table border="0" width="100%" cellpadding="3" cellspacing="1">
  <% for j=0 to iRows-1 %>
    <tr bgcolor="#FFFFFF" valign="top">
    <% for i=0 to iCols-1 %>
      <% if omileageshop.FResultCount>j*iCols+i then %>
      <td style="padding-bottom:5" width="7%">
        <div align="center">
          <table border="0" width="41" cellpadding="0" cellspacing="0">
            <tr>
              <td><a href="iframespecial_item.asp?itemid=<%= omileageshop.FItemList(j*iCols+i).FItemId %>" target=ispecialitem><img src="<%= omileageshop.FItemList(j*iCols+i).FImageList %>" width="100" height="100" border=0></a></td>
            </tr>
            <tr>
              <td>
                <div align="center">
                  <table border="0" width="100" cellpadding="0" cellspacing="1">
                    <tr>
                    	<td align=center><%= omileageshop.FItemList(j*iCols+i).FItemId %></td>
                    </tr>
                    <% if omileageshop.FItemList(j*iCols+i).IsSailItem then %>
                    <tr>
                      <td class="verdana-small">
                        <div align="center"><del><%= FormatNumber(omileageshop.FItemList(j*iCols+i).FOrgPrice,0) %></del>won</div>
                      </td>
                    </tr>
                    <% end if %>
                    <tr>
                      <td class="verdana-small">
                        <div align="center"><font color="#3366FF"><%= FormatNumber(omileageshop.FItemList(j*iCols+i).getRealPrice(),0) %>won</font></div>
                      </td>
                    </tr>
                    <% if omileageshop.FItemList(j*iCols+i).IsSailItem then %>
                    <tr>
                      <td class="verdana-small">
                        <div align="center"><font color="#3366FF" class="verdana-basic">(<%= omileageshop.FItemList(j*iCols+i).getSailPro() %> %
                          <span class="verdana-small"> sale</span>)</font></div>
                      </td>
                    </tr>
                    <% end if %>
                    <tr>
                    	<td align=center><input type=button value="일반상품으로변경" onclick="delThis('<%= omileageshop.FItemList(j*iCols+i).FItemId %>')"></td>
                    </tr>
                  </table>
                </div>
              </td>
            </tr>
          </table>
        </div>
      </td>
      <% else %>
      <td style="padding-bottom:8" >&nbsp; </td>
      <% end if %>
    <% next %>
    </tr>
    <% next %>
  </table>
<% end if %>
	</td>
</tr>
</table>

<%
set omileageshop = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->