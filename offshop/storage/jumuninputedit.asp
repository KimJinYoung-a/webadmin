<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionOffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%
dim idx, isfixed
idx = request("idx")
if idx="" then idx=0

dim ojumunmaster, ojumundetail

set ojumunmaster = new COrderSheet
ojumunmaster.FRectIdx = idx
ojumunmaster.GetOneOrderSheetMaster
isfixed = ojumunmaster.FOneItem.IsFixed

set ojumundetail= new COrderSheet
ojumundetail.FRectIdx = idx
ojumundetail.GetOrderSheetDetail

dim yyyymmdd
yyyymmdd = Left(ojumunmaster.FOneItem.Fscheduledate,10)
%>
<script language='javascript'>

<% if (ojumunmaster.FOneItem.FStatecd="0") or (ojumunmaster.FOneItem.FStatecd=" ") then %>
var jumunwait = true;
<% else %>
var jumunwait = false;
<% end if %>

<% if (Left(ojumunmaster.FOneItem.Fbaljucode,2) = "RJ") then %>
var rejumun = true;
<% else %>
var rejumun = false;
<% end if %>

function AddItems(frm){
	if (jumunwait!=true){
		alert('주문접수 상태가 아니면 수정하실 수 없습니다.');
		return;
	}

	if (rejumun == true){
		alert('재작성된 주문에서는 상품을 추가할수 없습니다.(상품준비중입니다.) \n다른 상품을 주문하시려면 별도의 주문서를 작성하세요.');
		return;
	}

	var popwin;
	var suplyer;

	if (frm.suplyer.value.length<1){
		alert('공급처를 먼저 선택하세요.');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;

	popwin = window.open('popshopjumunitem.asp?suplyer=' + suplyer + '&idx=' + frm.masteridx.value ,'offjumuninputeditadd','width=880,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ModiThis(frm){
	if (jumunwait!=true){
		alert('주문접수 상태가 아니면 수정하실 수 없습니다.');
		return;
	}

	if (rejumun == true){
		alert('재작성된 주문에서는 상품을 수정 하실 수 없습니다.(상품준비중입니다.) \n다른 상품을 주문하시려면 별도의 주문서를 작성하세요.');
		return;
	}


	var ret = confirm('수정 하시겠습니까?');

	if (ret){
		frm.mode.value="modidetail";
		frm.submit();
	}
}

function DelThis(frm){
	if (jumunwait!=true){
		alert('주문접수 상태가 아니면 수정하실 수 없습니다.');
		return;
	}

	if (rejumun == true){
		alert('재작성된 주문에서는 상품을 수정 하실 수 없습니다.(상품준비중입니다.) \n다른 상품을 주문하시려면 별도의 주문서를 작성하세요.');
		return;
	}

	var ret = confirm('삭제 하시겠습니까?');

	if (ret){
		frm.mode.value="deldetail";
		frm.submit();
	}
}

function DelMaster(frm){
	if (jumunwait!=true){
		alert('주문접수 상태가 아니면 수정하실 수 없습니다.');
		return;
	}

	if (rejumun == true){
		alert('재작성된 주문에서는 상품을 수정 하실 수 없습니다.(상품준비중입니다.) \n다른 상품을 주문하시려면 별도의 주문서를 작성하세요.');
		return;
	}

	var ret = confirm('삭제 하시겠습니까?');

	if (ret){
		frm.mode.value="delmaster";
		frm.submit();
	}
}

function ModiMaster(frm){
	if (jumunwait!=true){
		alert('주문접수 상태가 아니면 수정하실 수 없습니다.');
		return;
	}

	if (rejumun == true){
		alert('재작성된 주문에서는 상품을 수정 하실 수 없습니다.(상품준비중입니다.) \n다른 상품을 주문하시려면 별도의 주문서를 작성하세요.');
		return;
	}

	var ret = confirm('수정 하시겠습니까?');

	if (ret){
		frm.mode.value="modimaster";
		frm.submit();
	}
}

function ReActItems(iidx, igubun,iitemid,iitemoption,isellcash,isuplycash,ibuycash,iitemno,iitemname,iitemoptionname,iitemdesigner){
	if (iidx!='<%= idx %>'){
		alert('주문서가 일치하지 않습니다. 다시시도해 주세요.');
		return;
	}

	if (rejumun == true){
		alert('재작성된 주문에서는 상품을 수정 하실 수 없습니다.(상품준비중입니다.) \n다른 상품을 주문하시려면 별도의 주문서를 작성하세요.');
		return;
	}

	frmadd.itemgubunarr.value = igubun;
	frmadd.itemarr.value = iitemid;
	frmadd.itemoptionarr.value = iitemoption;
	frmadd.sellcasharr.value = isellcash;
	frmadd.suplycasharr.value = isuplycash;
	frmadd.buycasharr.value = ibuycash;
	frmadd.itemnoarr.value = iitemno;

	frmadd.submit();
}
</script>
<table width="760" cellspacing="1" class="a" bgcolor=#3d3d3d>
<form name="frmMaster" method="post" action="shopjumun_process.asp">
<input type=hidden name="mode" value="">
<input type=hidden name="masteridx" value="<%= idx %>">
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>공급처</td>
	<td>
	<input type=hidden name="suplyer" value="<%= ojumunmaster.FOneItem.Ftargetid %>">
	<%= ojumunmaster.FOneItem.Ftargetid %>&nbsp;(<%= ojumunmaster.FOneItem.Ftargetname %>)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>발주처</td>
	<td>
	<%= ojumunmaster.FOneItem.Fbaljuid %>&nbsp;(<%= ojumunmaster.FOneItem.Fbaljuname %>)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>등록일</td>
	<td><%= ojumunmaster.FOneItem.Fregdate %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>입고요청일</td>
	<td><input type=text name="yyyymmdd" value="<%= yyyymmdd %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.yyyymmdd);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a> (원하는 입고 날짜를 입력하세요.)</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>현재상태</td>
	<td><font color="<%= ojumunmaster.FOneItem.GetStateColor %>"><%= ojumunmaster.FOneItem.GetStateName %></font></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>총 소비자가(주문)</td>
	<td><%= FormatNumber(ojumunmaster.FOneItem.Fjumunsellcash,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>총 매입가(주문)</td>
	<td><%= FormatNumber(ojumunmaster.FOneItem.Fjumunsuplycash,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>총 소비자가(확정)</td>
	<td><%= FormatNumber(ojumunmaster.FOneItem.Ftotalsellcash,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>총 매입가(확정)</td>
	<td><%= FormatNumber(ojumunmaster.FOneItem.Ftotalsuplycash,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" width=100>기타요청사항</td>
	<td>
	<textarea name=comment cols=80 rows=6><%= ojumunmaster.FOneItem.FComment %></textarea>
	</td>
</tr>
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
<tr bgcolor="#FFFFFF">
	<td colspan="2" align=center>
		<input type=button value="수정" onclick="ModiMaster(frmMaster)">
		&nbsp;
		<input type=button value="전체삭제" onclick="DelMaster(frmMaster)">
	</td>
</tr>
</form>
</table>
<br>
<%

dim i,selltotal, suplytotal
selltotal =0
suplytotal =0
%>
<table width="760" cellspacing="0" class="a" bgcolor=#ffffff>
<tr>
	<td align=right><input type=button value="상품추가" onclick="AddItems(frmMaster)"></td>
</tr>
</table>
<table width="760" cellspacing="1" class="a" bgcolor=#3d3d3d>
	<tr bgcolor="#FFFFFF">
		<td colspan="9" align="right">총건수:  <%= ojumundetail.FResultCount %>&nbsp;</td>
	</tr>
	<tr bgcolor="#DDDDFF" align=center>
		<td width="100">바코드</td>
		<td width="100">브랜드</td>
		<td width="200">상품명</td>
		<td width="80">옵션명</td>
		<td width="80">판매가</td>
		<td width="80">공급가</td>
		<td width="60">주문갯수</td>
		<% if isfixed then %>
		<td width="60">확정갯수</td>
		<td width="60">비고</td>
		<% else %>
		<td width="40">수정</td>
		<td width="40">삭제</td>
		<% end if %>
	</tr>
	<% for i=0 to ojumundetail.FResultCount-1 %>
	<%
	selltotal  = selltotal + ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Fbaljuitemno
	suplytotal = suplytotal + ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Fbaljuitemno
	%>
	<form name="frmBuyPrc_<%= i %>" method="post" action="shopjumun_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="masteridx" value="<%= idx %>">
	<input type="hidden" name="itemgubun" value="<%= ojumundetail.FItemList(i).FItemGubun %>">
	<input type="hidden" name="itemid" value="<%= ojumundetail.FItemList(i).FItemID %>">
	<input type="hidden" name="itemoption" value="<%= ojumundetail.FItemList(i).Fitemoption %>">
	<input type="hidden" name="desingerid" value="<%= ojumundetail.FItemList(i).Fmakerid %>">
	<input type="hidden" name="sellcash" value="<%= ojumundetail.FItemList(i).FSellcash %>">
	<input type="hidden" name="suplycash" value="<%= ojumundetail.FItemList(i).FSuplycash %>">
	<input type="hidden" name="buycash" value="<%= ojumundetail.FItemList(i).Fbuycash %>">
	<tr bgcolor="#FFFFFF">
		<td ><%= ojumundetail.FItemList(i).FItemGubun %><%= CHKIIF(ojumundetail.FItemList(i).FItemID>=1000000,format00(8,ojumundetail.FItemList(i).FItemID),format00(6,ojumundetail.FItemList(i).FItemID)) %><%= ojumundetail.FItemList(i).Fitemoption %></td>
		<td ><%= ojumundetail.FItemList(i).Fmakerid %></td>
		<td ><%= ojumundetail.FItemList(i).Fitemname %></td>
		<td ><%= ojumundetail.FItemList(i).Fitemoptionname %></td>

		<td align=right><%= FormatNumber(ojumundetail.FItemList(i).Fsellcash,0) %></td>
		<td align=right><%= FormatNumber(ojumundetail.FItemList(i).Fsuplycash,0) %></td>
		<td align=center><input type="text" name="baljuitemno" value="<%= ojumundetail.FItemList(i).Fbaljuitemno %>"  size="4" maxlength="4"></td>
		<% if isfixed then %>
		<td><%= ojumundetail.FItemList(i).Frealitemno %></td>
		<td><%= ojumundetail.FItemList(i).Fcomment %></td>
		<% else %>
		<td><input type=button value="수정" onclick="ModiThis(frmBuyPrc_<%= i %>)"></td>
		<td><input type=button value="삭제" onclick="DelThis(frmBuyPrc_<%= i %>)"></td>
		<% end if %>
	</tr>
	</form>
	<% next %>

	<% if (ojumundetail.FResultCount>0) then %>
	<tr bgcolor="#FFFFFF">
		<td align="center">총계</td>
		<td colspan="3" align="center">
		<td align=right><%= formatNumber(selltotal,0) %></td>
		<td align=right><%= formatNumber(suplytotal,0) %></td>
		<td></td>
		<td></td>
		<td></td>
		</td>
	</tr>
	<% end if %>
</table>
<%
set ojumunmaster = Nothing
set ojumundetail = Nothing
%>
<form name="frmadd" method=post action="shopjumun_process.asp">
<input type=hidden name="mode" value="shopjumunitemaddarr">
<input type=hidden name="masteridx" value="<%= idx %>">
<input type=hidden name="itemgubunarr" value="">
<input type=hidden name="itemarr" value="">
<input type=hidden name="itemoptionarr" value="">
<input type=hidden name="sellcasharr" value="">
<input type=hidden name="suplycasharr" value="">
<input type=hidden name="buycasharr" value="">
<input type=hidden name="itemnoarr" value="">
</form>
<script language='javascript'>
if (jumunwait!=true){
	alert('주문접수 상태가 아니면 수정하실 수 없습니다.');
}else if (rejumun==true){
	alert('재작성된 주문서는 수정하실 수 없습니다.');
}
</script>
<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->