<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/100proshopCls.asp" -->

<style type="text/css">
.title {font-size:10px;font-famliy:"굴림";background-color:"CCBBBB";}
.content {font-size:12px;font-famliy:"굴림";}
</style>

<script language="javascript">
function TransItemInfo(currFrm){
	frm = eval('document.' + currFrm);

	opener.document.SubmitFrm.itemid.value=frm.itemid.value;
	opener.document.SubmitFrm.startdate.value=frm.startdate.value;
	opener.document.SubmitFrm.enddate.value=frm.enddate.value;
	opener.document.SubmitFrm.couponname.value=frm.couponname.value;
	opener.document.SubmitFrm.couponvalue.value=frm.couponvalue.value;
	opener.document.SubmitFrm.couponstartdate.value=frm.couponstartdate.value;
	opener.document.SubmitFrm.couponexpiredate.value=frm.couponexpiredate.value;
	opener.document.SubmitFrm.minbuyprice.value=frm.minbuyprice.value;
	opener.document.SubmitFrm.couponvalue.value=frm.couponvalue.value;


}
</script>
<%
dim eCode,i
eCode = request("eC")

dim o100pro
set o100pro = new C100ProShop
o100pro.getItemList eCode

%>
<% if o100pro.FResultCount > 0 then %>
<table width="300" border="1" cellpadding="1" cellspacing="0" class="verdana">
	<tr class="title">
		<td width="55" align="center">이미지</td>
		<td width="55" align="center">상품코드</td>
		<td width="70" align="center">쿠폰금액</td>
		<td width="100" align="center">최소구매금액</td>
	</tr>
	<% for i = 0 to o100pro.FResultCount -1 %>
	<tr class="content" onMouseOver="this.bgColor='#CCCCFF'" onMouseOut="this.bgColor=''" style="cursor:pointer" onclick="TransItemInfo('List_<%= i %>');">
	<form name="List_<%= i %>" >
		<input type="hidden" name="itemid" value="<%= o100pro.FItemList(i).Fitemid %>" />
		<input type="hidden" name="startdate" value="<% = FormatDateTime(o100pro.FItemList(i).FStartDate,2) %>" />
		<input type="hidden" name="enddate" value="<%= FormatDateTime(o100pro.FItemList(i).Fenddate,2) %>" />
		<input type="hidden" name="couponname" value="<%= o100pro.FItemList(i).FCouponName %>" />
		<input type="hidden" name="couponvalue" value="<%= o100pro.FItemList(i).FCouponValue %>" />
		<input type="hidden" name="couponstartdate" value="<%= FormatDateTime(o100pro.FItemList(i).FCouponStartDate,2) %>" />
		<input type="hidden" name="couponexpiredate" value="<%= FormatDateTime(o100pro.FItemList(i).FCouponExpireDate,2) %>" />
		<input type="hidden" name="minbuyprice" value="<%= o100pro.FItemList(i).Fminbuyprice %>" />


		<td align="center"><img src="<%= o100pro.FItemList(i).FItemImageSmall %>" border="0"></td>
		<td align="center"><b><%= o100pro.FItemList(i).Fitemid %></b></td>
		<td align="center"><%= FormatNumber(o100pro.FItemList(i).FCouponValue,0) %></td>
		<td align="center"><%= FormatNumber(o100pro.FItemList(i).Fminbuyprice,0) %></td>
	</form>
	</tr>
	<% next %>
</table>
<% else %>
<script language="javascript">
alert('등록된 상품이 없습니다');
window.close();
</script>
<% end if%>

<% set o100pro = nothing %>

<!-- #include virtual="/lib/db/dbclose.asp" -->
