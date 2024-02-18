<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionOffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%
dim shopid, reguser, divcode,baljuname,regname
shopid = session("ssBctId")
reguser = shopid

divcode = session("ssBctDiv")
baljuname = session("ssBctCname")
regname = baljuname


dim osheetmaster, idx
idx = request("idx")
if idx="" then idx=0

dim suplyer,yyyymmdd,comment
suplyer = request("suplyer")
yyyymmdd = request("yyyymmdd")
comment = request("comment")

if suplyer="" then suplyer="10x10"

dim vatcode
dim itemgubunarr, itemidadd, itemoptionarr
dim itemnamearr, itemoptionnamearr
dim sellcasharr, suplycasharr, buycasharr, itemnoarr, designerarr

dim itemgubunarr2, itemidadd2, itemoptionarr2
dim itemnamearr2, itemoptionnamearr2
dim sellcasharr2, suplycasharr2, buycasharr2, itemnoarr2, designerarr2

dim itemgubunarr3, itemidadd3, itemoptionarr3
dim itemnamearr3, itemoptionnamearr3
dim sellcasharr3, suplycasharr3, buycasharr3, itemnoarr3, designerarr3

dim i,j,cnt,cnt2

itemgubunarr = request("itemgubunarr")
itemidadd	= request("itemidadd")
itemoptionarr = request("itemoptionarr")
itemnamearr		= request("itemnamearr")
itemoptionnamearr = request("itemoptionnamearr")
sellcasharr = request("sellcasharr")
suplycasharr = request("suplycasharr")
buycasharr = request("buycasharr")
itemnoarr = request("itemnoarr")
designerarr = request("designerarr")

itemgubunarr2 = request("itemgubunarr2")
itemidadd2	= request("itemidadd2")
itemoptionarr2 = request("itemoptionarr2")
itemnamearr2	= request("itemnamearr2")
itemoptionnamearr2 = request("itemoptionnamearr2")
sellcasharr2 = request("sellcasharr2")
suplycasharr2 = request("suplycasharr2")
buycasharr2 = request("buycasharr2")
itemnoarr2 = request("itemnoarr2")
designerarr2 = request("designerarr2")


'chargeid = request("chargeid")
'shopid = session("ssBctID")
'vatcode = request("vatcode")
'divcode  = request("divcode")


itemgubunarr = split(itemgubunarr,"|")
itemidadd	= split(itemidadd,"|")
itemoptionarr = split(itemoptionarr,"|")
itemnamearr		= split(itemnamearr,"|")
itemoptionnamearr = split(itemoptionnamearr,"|")
sellcasharr = split(sellcasharr,"|")
suplycasharr = split(suplycasharr,"|")
buycasharr = split(buycasharr,"|")
itemnoarr = split(itemnoarr,"|")
designerarr = split(designerarr,"|")

itemgubunarr2 = split(itemgubunarr2,"|")
itemidadd2	= split(itemidadd2,"|")
itemoptionarr2 = split(itemoptionarr2,"|")
itemnamearr2		= split(itemnamearr2,"|")
itemoptionnamearr2 = split(itemoptionnamearr2,"|")
sellcasharr2 = split(sellcasharr2,"|")
suplycasharr2 = split(suplycasharr2,"|")
buycasharr2 = split(buycasharr2,"|")
itemnoarr2 = split(itemnoarr2,"|")
designerarr2 = split(designerarr2,"|")

cnt = uBound(itemidadd)
cnt2 = uBound(itemidadd2)


dim isPreExists

for j=0 to cnt2-1
	isPreExists = false
	for i=0 to cnt-1
		if (itemgubunarr(i)=itemgubunarr2(j)) and (itemidadd(i)=itemidadd2(j)) and (itemoptionarr(i)=itemoptionarr2(j)) then
			itemnoarr(i) = CStr(CLng(itemnoarr(i)) + CLng(itemnoarr2(j)))
			isPreExists = true
			exit for
		end if
	next

	if Not isPreExists then
		itemgubunarr3 = itemgubunarr3 + itemgubunarr2(j) + "|"
		itemidadd3	= itemidadd3 + itemidadd2(j) + "|"
		itemoptionarr3 = itemoptionarr3 + itemoptionarr2(j) + "|"
		itemnamearr3		= itemnamearr3 + itemnamearr2(j) + "|"
		itemoptionnamearr3  = itemoptionnamearr3 + itemoptionnamearr2(j) + "|"
		sellcasharr3 = sellcasharr3 + sellcasharr2(j) + "|"
		suplycasharr3 = suplycasharr3 + suplycasharr2(j) + "|"
		buycasharr3 = buycasharr3 + buycasharr2(j) + "|"
		itemnoarr3 = itemnoarr3 + itemnoarr2(j) + "|"
		designerarr3 = designerarr3 + designerarr2(j) + "|"
	end if
next

itemgubunarr2 = ""
itemidadd2	= ""
itemoptionarr2 = ""
itemnamearr2	= ""
itemoptionnamearr2 = ""
sellcasharr2 = ""
suplycasharr2 = ""
buycasharr2 = ""
itemnoarr2 = ""
designerarr2 = ""

for i=0 to cnt-1
	itemgubunarr2 = itemgubunarr2 + itemgubunarr(i) + "|"
	itemidadd2	= itemidadd2 + itemidadd(i) + "|"
	itemoptionarr2 = itemoptionarr2 + itemoptionarr(i) + "|"
	itemnamearr2	= itemnamearr2 + itemnamearr(i) + "|"
	itemoptionnamearr2 = itemoptionnamearr2 + itemoptionnamearr(i) + "|"
	sellcasharr2 = sellcasharr2 + sellcasharr(i) + "|"
	suplycasharr2 = suplycasharr2 + suplycasharr(i) + "|"
	buycasharr2 = buycasharr2 + buycasharr(i) + "|"
	itemnoarr2 = itemnoarr2 + itemnoarr(i) + "|"
	designerarr2 = designerarr2 + designerarr(i) + "|"
next

itemgubunarr = itemgubunarr2 + itemgubunarr3
itemidadd	= itemidadd2 + itemidadd3
itemoptionarr = itemoptionarr2 + itemoptionarr3
itemnamearr	= itemnamearr2 + itemnamearr3
itemoptionnamearr = itemoptionnamearr2 + itemoptionnamearr3
sellcasharr = sellcasharr2 + sellcasharr3
suplycasharr = suplycasharr2 + suplycasharr3
buycasharr = buycasharr2 + buycasharr3
itemnoarr = itemnoarr2 + itemnoarr3
designerarr = designerarr2 + designerarr3
%>
<script language='javascript'>
function ReActItems(iidx, igubun,iitemid,iitemoption,isellcash,isuplycash,ibuycash,iitemno,iitemname,iitemoptionname,iitemdesigner){
	if (iidx!='0'){
		alert('주문서가 일치하지 않습니다. 다시시도해 주세요.');
		return;
	}

	frmMaster.itemgubunarr2.value = igubun;
	frmMaster.itemidadd2.value = iitemid;
	frmMaster.itemoptionarr2.value = iitemoption;
	frmMaster.sellcasharr2.value = isellcash;
	frmMaster.suplycasharr2.value = isuplycash;
	frmMaster.buycasharr2.value = ibuycash;
	frmMaster.itemnoarr2.value = iitemno;
	frmMaster.itemnamearr2.value = iitemname;
	frmMaster.itemoptionnamearr2.value = iitemoptionname;
	frmMaster.designerarr2.value = iitemdesigner;

	frmMaster.submit();
}

function AddItems(frm){
	var popwin;
	var suplyer;

	if (frm.suplyer.value.length<1){
		alert('공급처를 먼저 선택하세요.');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;

	popwin = window.open('popshopjumunitem.asp?suplyer=' + suplyer + '&idx=0','offjumuninputadd','width=880,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ConFirmIpChulList(){
	var msfrm = document.frmMaster;
	var upfrm = document.frmArrupdate;
	var frm;

	if (msfrm.yyyymmdd.value.length<1){
		alert('입고요청일을 입력해 주세요.');
		//msfrm.yyyymmdd.value.focus();
		return;
	}

	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.sellcasharr.value = "";
	upfrm.suplycasharr.value = "";
	upfrm.buycasharr.value = "";
	upfrm.itemnoarr.value = "";
	upfrm.designerarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {

			if (!IsInteger(frm.itemno.value)){
				alert('갯수는 정수만 가능합니다.');
				frm.itemno.focus();
				return;
			}

			upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
			upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
			upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
			upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
			upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
			upfrm.buycasharr.value = upfrm.buycasharr.value + frm.buycash.value + "|";
			upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + "|";
			upfrm.designerarr.value = upfrm.designerarr.value + frm.desingerid.value + "|";
		}
	}

	var ret = confirm('저장 하시겠습니까?');

	if (ret){
		upfrm.yyyymmdd.value = msfrm.yyyymmdd.value;
		upfrm.comment.value = msfrm.comment.value;

		upfrm.submit();
	}
}
</script>
<table width="100%" cellspacing="1" class="a" bgcolor=#3d3d3d>
<form name="frmMaster" method="post" action="">
<input type="hidden" name="mode" value="addmaster">
<input type="hidden" name="itemgubunarr" value="<%= itemgubunarr %>">
<input type="hidden" name="itemidadd" value="<%= itemidadd %>">
<input type="hidden" name="itemoptionarr" value="<%= itemoptionarr %>">
<input type="hidden" name="itemnamearr" value="<%= itemnamearr %>">
<input type="hidden" name="itemoptionnamearr" value="<%= itemoptionnamearr %>">
<input type="hidden" name="sellcasharr" value="<%= sellcasharr %>">
<input type="hidden" name="suplycasharr" value="<%= suplycasharr %>">
<input type="hidden" name="buycasharr" value="<%= buycasharr %>">
<input type="hidden" name="itemnoarr" value="<%= itemnoarr %>">
<input type="hidden" name="designerarr" value="<%= designerarr %>">

<input type="hidden" name="itemgubunarr2" value="">
<input type="hidden" name="itemidadd2" value="">
<input type="hidden" name="itemoptionarr2" value="">
<input type="hidden" name="itemnamearr2" value="">
<input type="hidden" name="itemoptionnamearr2" value="">
<input type="hidden" name="sellcasharr2" value="">
<input type="hidden" name="suplycasharr2" value="">
<input type="hidden" name="buycasharr2" value="">
<input type="hidden" name="itemnoarr2" value="">
<input type="hidden" name="designerarr2" value="">
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>공급처</td>
	<% if suplyer<>"" then %>
	<input type=hidden name="suplyer" value="<%= suplyer %>">
	<td><%= suplyer %></td>
	<% else %>
	<td><% SelectBoxOffShopSuplyer "suplyer", suplyer, shopid, session("ssBctDiv") %></td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>입고요청일</td>
	<td><input type=text name="yyyymmdd" value="<%= yyyymmdd %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.yyyymmdd);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a> (원하는 입고 날짜를 입력하세요.)</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>기타요청사항</td>
	<td>
	<textarea name=comment cols=80 rows=6><%= comment %></textarea>
	</td>
</tr>
</form>
<tr  bgcolor="#FFFFFF">
	<td colspan="2">
	<br>
	* 5일내 출고 : 업체 배송 상품 (물류센터로 입고 되는대로 매장으로 발송 해드리겠습니다.) <br>
	* 재고 부족 : 물류센터 재고 부족으로 인해 업체로 발주가 들어가 있는 상태입니다. <br>
				2~3일 내로 입고 될 수 있는 상품 입니다. 따로 보내드리지 않으며, <B>다음 주문시 추가(재주문)</B>해 주셔야 합니다.<br>
	* 일시품절 : 업체 재고부족으로 인해 재생산중인 상품입니다.(단기간 내에 입고 되기 어려운 상품입니다.)
	<br>
	</td>
</tr>
</table>
<table width="100%" cellspacing="0" class="a" bgcolor=#ffffff>
<tr>
	<td align=right><input type=button value="상품추가" onclick="AddItems(frmMaster)"></td>
</tr>
</table>

<%
itemgubunarr = split(itemgubunarr,"|")
itemidadd	= split(itemidadd,"|")
itemoptionarr = split(itemoptionarr,"|")
itemnamearr		= split(itemnamearr,"|")
itemoptionnamearr = split(itemoptionnamearr,"|")
sellcasharr = split(sellcasharr,"|")
suplycasharr = split(suplycasharr,"|")
buycasharr = split(buycasharr,"|")
itemnoarr = split(itemnoarr,"|")
designerarr = split(designerarr,"|")

cnt = ubound(itemidadd)

dim selltotal, suplytotal
selltotal =0
suplytotal =0
%>

<table width="100%" cellspacing="1" class="a" bgcolor=#3d3d3d>
	<tr bgcolor="#FFFFFF">
		<td colspan="6" align="right">총건수:  <%= cnt+1 %>&nbsp;</td>
	</tr>
	<tr bgcolor="#DDDDFF">
		<td width="100">바코드</td>
		<td width="200">상품명</td>
		<td width="80">옵션명</td>
		<td width="80">판매가</td>
		<td width="80">공급가</td>
		<td width="60">갯수</td>
	</tr>
	<% for i=0 to cnt-1 %>
	<%
	selltotal  = selltotal + sellcasharr(i) * itemnoarr(i)
	suplytotal = suplytotal + suplycasharr(i) * itemnoarr(i)
	%>
	<form name="frmBuyPrc_<%= i %>" method="post" action="">
	<input type="hidden" name="itemgubun" value="<%= itemgubunarr(i) %>">
	<input type="hidden" name="itemid" value="<%= itemidadd(i) %>">
	<input type="hidden" name="itemoption" value="<%= itemoptionarr(i) %>">
	<input type="hidden" name="desingerid" value="<%= designerarr(i) %>">
	<input type="hidden" name="sellcash" value="<%= sellcasharr(i) %>">
	<input type="hidden" name="suplycash" value="<%= suplycasharr(i) %>">
	<input type="hidden" name="buycash" value="<%= buycasharr(i) %>">
	<tr bgcolor="#FFFFFF">
		<td ><%= itemgubunarr(i) %><%= CHKIIF(itemidadd(i)>=1000000,format00(8,itemidadd(i)),format00(6,itemidadd(i))) %><%= itemoptionarr(i) %></td>
		<td ><%= itemnamearr(i) %></td>
		<td ><%= itemoptionnamearr(i) %></td>
		<td align=right><%= FormatNumber(sellcasharr(i),0) %></td>
		<td align=right><%= FormatNumber(suplycasharr(i),0) %></td>
		<td ><input type="text" name="itemno" value="<%= itemnoarr(i) %>"  size="4" maxlength="4"></td>
	</tr>
	</form>
	<% next %>

	<% if (cnt>0) then %>
	<tr bgcolor="#FFFFFF">
		<td align="center">총계</td>
		<td colspan="2" align="center">
		<td align=right><%= formatNumber(selltotal,0) %></td>
		<td align=right><%= formatNumber(suplytotal,0) %></td>
		<td></td>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="6" align="center">
		<input type="button" value="내역확정" onclick="ConFirmIpChulList()">
		</td>
	</tr>
	<% end if %>
</table>
<form name="frmArrupdate" method="post" action="shopjumun_process.asp">
<input type="hidden" name="mode" value="addshopjumun">
<input type="hidden" name="yyyymmdd" value="">

<input type="hidden" name="baljuid" value="<%= shopid %>">
<input type="hidden" name="targetid" value="<%= suplyer %>">
<input type="hidden" name="reguser" value="<%= reguser %>">
<input type="hidden" name="divcode" value="<%= divcode %>">
<input type="hidden" name="vatinclude" value="Y">
<input type="hidden" name="comment" value="">
<input type="hidden" name="regname" value="<%= regname %>">
<input type="hidden" name="baljuname" value="<%= baljuname %>">

<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="sellcasharr" value="">
<input type="hidden" name="suplycasharr" value="">
<input type="hidden" name="buycasharr" value="">
<input type="hidden" name="itemnoarr" value="">
<input type="hidden" name="designerarr" value="">

</form>
<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->