<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/stock/upcheorderitemcls.asp"-->
<%

dim suplyer, shopid, barcode, idx
dim research, onoffgubun
dim ErrStr

suplyer = request("suplyer")
shopid = request("shopid")
barcode = request("barcode")
idx = request("idx")
research = request("research")
onoffgubun = request("onoffgubun")

if (onoffgubun = "") then
	onoffgubun = "online"
end if



dim itemgubun, itemid,itemoption
dim sqlStr

if (Len(barcode)=12) then
	itemgubun = Left(barcode,2)
	itemid = CLng(Mid(barcode,3,6))
	itemoption = Right(barcode,4)

elseif Len(barcode)>8 then
	sqlStr = "select itemgubun,shopitemid,itemoption  " + VbCrlf
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item" + VbCrlf
	sqlStr = sqlStr + " where extbarcode='" + barcode + "'" + VbCrlf

	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		itemgubun = rsget("itemgubun")
		itemid = rsget("shopitemid")
		itemoption = rsget("itemoption")
	end if
	rsget.Close
else

end if



dim ojumunitem
set ojumunitem  = new CUpcheOrderItem
ojumunitem.FPageSize = 50
ojumunitem.FCurrPage = 1
ojumunitem.FRectDesigner = suplyer
'ojumunitem.FRectNoSearchUpcheBeasong = nubeasong
'ojumunitem.FRectNoSearchNotusingItem = nuitem
'ojumunitem.FRectNoSearchNotusingItemOption = nuitemoption
'ojumunitem.FRectNoSearchDanjong = nudanjong
'ojumunitem.FRectNoSearchSoldoutover7days = soldoutover7days
ojumunitem.FRectItemgubun = itemgubun
ojumunitem.FRectItemid = itemid
ojumunitem.FRectItemOption = itemoption
'ojumunitem.FRectShortage7days = chkIIF(ShortageType="7day","on","")
'ojumunitem.FRectShortage14days = chkIIF(ShortageType="14day","on","")
'ojumunitem.FRectShortageRealStock = chkIIF(ShortageType="5under","on","")

if (suplyer<>"") and (itemgubun<>"") and (itemid<>"") and (itemoption<>"") then
	if onoffgubun="offline" then
		ojumunitem.GetOffShopItemList
	else
		ojumunitem.GetOnLineJumunByBrand
	end if
end if

%>
<script language='javascript'>
function search(frm){
	frm.submit();
}

function AddArr(upfrm){

	opener.ReActItems('<%= idx %>', upfrm.itemgubunarr.value,upfrm.itemarr.value,upfrm.itemoptionarr.value,
		upfrm.sellcasharr.value,upfrm.suplycasharr.value,upfrm.buycasharr.value,upfrm.itemnoarr.value,upfrm.itemnamearr.value,
		upfrm.itemoptionnamearr.value,upfrm.designerarr.value,upfrm.mwdivarr.value);

}


function GetOnLoad(){
	document.frm.barcode.focus();
	document.frm.barcode.select();

}
window.onload = GetOnLoad;
</script>
<br>
<table width="500" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="suplyer"value="<%= suplyer %>">
	<input type="hidden" name="shopid"value="<%= shopid %>">
	<input type="hidden" name="idx"value="<%= idx %>">
	<input type="hidden" name="research" value="on">

	<tr>
		<td class="a" >
			<input type="radio" name="onoffgubun" value="online" <% if onoffgubun="online" then response.write "checked" %> >온라인
			<input type="radio" name="onoffgubun" value="offline" <% if onoffgubun="offline" then response.write "checked" %> >오프라인
		</td>
		<td class="a" align="right">
			바코드 :
			<input type="text" name="barcode" value="" size="13" maxlength="20" AUTOCOMPLETE="off">
			<a href="javascript:search(frm);"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<% if ojumunitem.FResultCount>0 then %>
	<!-- 입력값 체크는 차후에 한다.
	<//% if IsNULL(ioffitem.FOneItem.Fdefaultsuplymargin) then %//>
		<//%
		ErrStr = "[계약 안된 브랜드 입니다. 마진 설정후 사용 요망.]"
		%//>
	<//% elseif ioffitem.FOneItem.Fchargediv<>"2" and ioffitem.FOneItem.Fchargediv<>"4" and ioffitem.FOneItem.Fchargediv<>"5" then %//>
		<//%
		ErrStr = "[업체위탁이나 업체매입은 사용할 수 없습니다." +  ioffitem.FOneItem.Fchargediv + "]"
		%//>
	<//% else %//>
	-->
<form name="upfrm" >
<input type="hidden" name="itemgubunarr" value="<%= ojumunitem.FItemList(0).FItemgubun %>|">
<input type="hidden" name="itemarr" value="<%= ojumunitem.FItemList(0).Fitemid %>|">
<input type="hidden" name="itemoptionarr" value="<%= ojumunitem.FItemList(0).Fitemoption %>|">
<input type="hidden" name="sellcasharr" value="<%= ojumunitem.FItemList(0).Fsellcash %>|">
<input type="hidden" name="suplycasharr" value="<%= ojumunitem.FItemList(0).FBuycash %>|">
<input type="hidden" name="buycasharr" value="<%= ojumunitem.FItemList(0).FBuycash %>|">
<input type="hidden" name="itemnoarr" value="1|">
<input type="hidden" name="itemnamearr" value='<%= replace(ojumunitem.FItemList(0).Fitemname,"'","") %>|'>
<input type="hidden" name="itemoptionnamearr" value="<%= ojumunitem.FItemList(0).Fitemoptionname %>|">
<input type="hidden" name="designerarr" value="<%= ojumunitem.FItemList(0).Fmakerid %>|">
<input type="hidden" name="mwdivarr" value="<%= ojumunitem.FItemList(0).Fmwdiv %>|">
</form>
<script language='javascript'>
AddArr(upfrm);
</script>
<table width="500" class="a">
<tr>
	<td align="center"><font color="blue">[<%= ojumunitem.FItemList(0).Fmakerid %>] <%= ojumunitem.FItemList(0).Fitemname %> <%= ojumunitem.FItemList(0).Fitemoptionname %></font> 추가 완료</td>
</tr>
</table>
	<!--<//% end if %//>-->
<% elseif research="on" then %>
	<%
	ErrStr = "[검색결과가 없습니다.]"
	%>
<% end if %>

<% if ErrStr<>"" then %>
<table width="500" class="a">
<tr>
	<td align="center"><font color="red"><%= ErrStr %></font></td>
</tr>
</table>
<script language='javascript'>
alert('<%= ErrStr %>');
</script>
<% end if %>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
