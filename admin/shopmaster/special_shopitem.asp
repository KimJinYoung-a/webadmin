<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/specialshopitemcls.asp"-->
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
	sqlstr = sqlstr + " set specialuseritem=0" + VbCrlf
	sqlstr = sqlstr + " where itemid=" + CStr(delitem) + VbCrlf

	rsget.Open sqlStr,dbget,1
end if

if itemidarr<>"" then
	sqlstr = " update [db_item].[dbo].tbl_item" + VbCrlf
	sqlstr = sqlstr + " set specialuseritem=1" + VbCrlf
	sqlstr = sqlstr + " where itemid in (" + CStr(itemidarr) + ")"

	rsget.Open sqlStr,dbget,1
end if

dim ospecialshop
set ospecialshop = new CSpecialShop
ospecialshop.FPageSize=100
ospecialshop.FRectUserLevelUnder = 3
ospecialshop.GetSpecialItemList

dim iCols, iRows
iCols=7
iRows = Clng(ospecialshop.FResultCount \ iCols)
if (ospecialshop.FResultCount mod iCols)>0 then
	iRows = iRows + 1
end if

dim i,j
%>
<script language='javascript'>
function delThis(v){
	if (confirm('삭제 하시겠습니까?')){
		document.location = "?delitem="+ v;
	}
}

function SaveArr(frm){
	if (confirm('추가 하시겠습니까?')){
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
		<input type=button value="상품추가" onClick="SaveArr(frm)"></td>
		<td class="a" align="right">
		</td>
	</tr>
	</form>
</table>
<table border="0" width="750" cellpadding="0" cellspacing="0">
<tr>
	<td>
<% if ospecialshop.FResultCount<1 then %>
   [ 검색 결과가 없습니다. ]
<% else %>
  <table border="0" width="100%" cellpadding="3" cellspacing="1">
  <% for j=0 to iRows-1 %>
    <tr bgcolor="#FFFFFF" valign="top">
    <% for i=0 to iCols-1 %>
      <% if ospecialshop.FResultCount>j*iCols+i then %>
      <td style="padding-bottom:5" width="7%">
        <div align="center">
          <table border="0" width="41" cellpadding="0" cellspacing="0">
            <tr>
              <td><a href="iframespecial_item.asp?itemid=<%= ospecialshop.FItemList(j*iCols+i).FItemId %>" target=ispecialitem><img src="<%= ospecialshop.FItemList(j*iCols+i).FImageList %>" width="100" height="100" border=0></a></td>
            </tr>
            <tr>
              <td>
                <div align="center">
                  <table border="0" width="100" cellpadding="0" cellspacing="1">
                    <tr>
                    	<td align=center><%= ospecialshop.FItemList(j*iCols+i).FItemId %></td>
                    </tr>
                    <% if ospecialshop.FItemList(j*iCols+i).IsSailItem then %>
                    <tr>
                      <td class="verdana-small">
                        <div align="center"><del><%= FormatNumber(ospecialshop.FItemList(j*iCols+i).FOrgPrice,0) %></del>won</div>
                      </td>
                    </tr>
                    <% end if %>
                    <tr>
                      <td class="verdana-small">
                        <div align="center"><font color="#3366FF"><%= FormatNumber(ospecialshop.FItemList(j*iCols+i).getRealPrice(),0) %>won</font></div>
                      </td>
                    </tr>
                    <% if ospecialshop.FItemList(j*iCols+i).IsSailItem then %>
                    <tr>
                      <td class="verdana-small">
                        <div align="center"><font color="#3366FF" class="verdana-basic">(<%= ospecialshop.FItemList(j*iCols+i).getSailPro() %> %
                          <span class="verdana-small"> sale</span>)</font></div>
                      </td>
                    </tr>
                    <% end if %>
                    <tr>
                    	<td align=center><input type=button value="삭제" onclick="delThis('<%= ospecialshop.FItemList(j*iCols+i).FItemId %>')"></td>
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
set ospecialshop = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->