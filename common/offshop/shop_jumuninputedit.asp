<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프샵 주문서 작성
' History : 2009.04.07 서동석 생성
'			2011.01.13 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/common/lib/incMultiLangConst.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%

dim IS_HIDE_BUYCASH : IS_HIDE_BUYCASH = False
if C_IS_OWN_SHOP or C_IS_SHOP then
	IS_HIDE_BUYCASH = True
end if

dim idx, isfixed ,ojumunmaster, ojumundetail
	idx = requestCheckVar(request("idx"),10)

if idx="" then idx=0

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

<script type='text/javascript'>

<% if (ojumunmaster.FOneItem.FStatecd=" ") then %>
	var jumunwait = true;
<% else %>
	var jumunwait = false;
<% end if %>

<% if (Left(ojumunmaster.FOneItem.Fbaljucode,2) = "RJ") then %>
	var rejumun = true;
<% else %>
	var rejumun = false;
<% end if %>

function jsPopCal(fName,sName){
	var fd = eval("document."+fName+"."+sName);

	var winCal;
	winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function AddItems(frm){
	if (jumunwait!=true){
		alert('주문접수 이후에는 수정하실 수 없습니다.');
		return;
	}

	if (rejumun == true){
		alert('재작성된 주문에서는 상품을 추가할수 없습니다.(상품준비중입니다.) \n다른 상품을 주문하시려면 별도의 주문서를 작성하세요.');
		return;
	}

	var popwin;
	var suplyer;

	if (frm.suplyer.value.length<1){
		alert('<%= CTX_Please_select %> (<%= CTX_WHOLESALEID %>)');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;

	var cwflag;
	for (var i =0 ; i < frm.cwflag.length ; i++){
		if (frm.cwflag[i].checked){
			cwflag = frm.cwflag[i].value;
		}
	}

	popwin = window.open('/common/offshop/popshopjumunitem.asp?suplyer=' + suplyer + '&idx=' + frm.masteridx.value +'&cwflag='+cwflag,'offjumuninputeditadd','width=880,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ModiThis(frm){
	if (jumunwait!=true){
		alert('주문접수 이후에는 수정하실 수 없습니다.');
		return;
	}

	if (rejumun == true){
		alert('재작성된 주문에서는 상품을 수정 하실 수 없습니다.(상품준비중입니다.) \n다른 상품을 주문하시려면 별도의 주문서를 작성하세요.');
		return;
	}


	var ret = confirm('<%= CTX_Are_you_sure_you_want_to_continue %> (<%= CTX_Revise %>)?');

	if (ret){
		frm.mode.value="modidetail";
		frm.submit();
	}
}

function DelThis(frm){
	if (jumunwait!=true){
		alert('주문접수 이후에는 수정하실 수 없습니다.');
		return;
	}

	if (rejumun == true){
		alert('재작성된 주문에서는 상품을 수정 하실 수 없습니다.(상품준비중입니다.) \n다른 상품을 주문하시려면 별도의 주문서를 작성하세요.');
		return;
	}

	var ret = confirm('<%= CTX_Are_you_sure_you_want_to_continue %> (<%= CTX_Delete %>)?');

	if (ret){
		frm.mode.value="deldetail";
		frm.submit();
	}
}

function DelMaster(frm){
	if (jumunwait!=true){
		alert('주문접수 이후에는 수정하실 수 없습니다.');
		return;
	}

	if (rejumun == true){
		alert('재작성된 주문에서는 상품을 수정 하실 수 없습니다.(상품준비중입니다.) \n다른 상품을 주문하시려면 별도의 주문서를 작성하세요.');
		return;
	}

	var ret = confirm('<%= CTX_Are_you_sure_you_want_to_continue %> (<%= CTX_Delete %>)?');

	if (ret){
		frm.mode.value="delmaster";
		frm.submit();
	}
}

function ModiMaster(frm){
	if (jumunwait!=true){
		alert('주문접수 이후에는 수정하실 수 없습니다.');
		return;
	}

	if (rejumun == true){
		alert('재작성된 주문에서는 상품을 수정 하실 수 없습니다.(상품준비중입니다.) \n다른 상품을 주문하시려면 별도의 주문서를 작성하세요.');
		return;
	}

	var ret = confirm('<%= CTX_Are_you_sure_you_want_to_continue %> (<%= CTX_Revise %>)?');

	if (ret){
		frm.mode.value="modimaster";
		frm.submit();
	}
}

function ReActItems(iidx, igubun,iitemid,iitemoption,isellcash,isuplycash,ibuycash,iitemno,iitemname,iitemoptionname,iitemdesigner){
	if (iidx!='<%= idx %>'){
		alert('<%= CTX_Does_not_match %> (<%= CTX_Order_code %> :' + iidx + ')');
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

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmMaster" method="post" action="common_shopjumun_process.asp">
<input type=hidden name="mode" value="">
<input type=hidden name="masteridx" value="<%= idx %>">

<!-- 상단바 시작 -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong><%= CTX_Ordering_Information %></strong></font>
			        &nbsp;
			        <b>[ <%= ojumunmaster.FOneItem.FBaljuCode %> ]</b>
			    </td>
			    <td align="right">
					<!-- <input type="button" class="button" value="목록으로 이동" onclick=""> -->
				</td>
			</tr>
		</table>
	</td>
</tr>
<!-- 상단바 끝 -->
<tr bgcolor="#FFFFFF">
	<td width="110" bgcolor="<%= adminColor("tabletop") %>" ><%= CTX_WHOLESALEID %></td>
	<td width="400">
		<input type="hidden" name="suplyer" value="<%= ojumunmaster.FOneItem.Ftargetid %>">
		<%= ojumunmaster.FOneItem.Ftargetid %>&nbsp;(<%= ojumunmaster.FOneItem.Ftargetname %>)
	</td>
	<td width="120" bgcolor="<%= adminColor("tabletop") %>" ><%= CTX_an_orderer %>(SHOP)</td>
	<td><%= ojumunmaster.FOneItem.Fbaljuid %>&nbsp;(<%= ojumunmaster.FOneItem.Fbaljuname %>)</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>"><%= CTX_Order_Date %></td>
	<td><%= ojumunmaster.FOneItem.Fregdate %></td>
	<td bgcolor="<%= adminColor("tabletop") %>"><%= CTX_were_stocked_requested_date %></td>
	<td>
		<input type="text" class="text" name="yyyymmdd" value="<%= yyyymmdd %>" size=10 readonly >
		<a href="javascript:jsPopCalendar('frmMaster','yyyymmdd');"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>"><%= CTX_Status %></td>
	<td>
	    <% if (ojumunmaster.FOneItem.FStatecd=" ") or (ojumunmaster.FOneItem.FStatecd="0") then %>
	    <input type=radio name="statecd" value=" " <% if ojumunmaster.FOneItem.FStatecd=" " then response.write "checked" %> ><%= CTX_in_process %>
		<input type=radio name="statecd" value="0" <% if ojumunmaster.FOneItem.FStatecd="0" then response.write "checked" %> ><%= CTX_Register %>
	    <% else %>
		<font color="<%= ojumunmaster.FOneItem.GetStateColor %>"><%= ojumunmaster.FOneItem.GetStateName %></font>
		<% end if %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>"><%= CTX_release_divide %></td>
	<td>
		<input type="radio" disabled name="cwflag" value="0" <% if ojumunmaster.FOneItem.fcwflag="0" then response.write " checked" %>><%= CTX_release_Purchase %>
		<input type="radio" disabled name="cwflag" value="1" <% if ojumunmaster.FOneItem.fcwflag="1" then response.write " checked" %>><%= CTX_release_on_consignment %>
		<input type="hidden" name="cwflag" value="<%=ojumunmaster.FOneItem.fcwflag%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>"><%= CTX_total_consumer_price %>(<%= CTX_request %>)</td>
	<td><%= FormatNumber(ojumunmaster.FOneItem.Fjumunsellcash,0) %></td>
	<td bgcolor="<%= adminColor("tabletop") %>"><%= CTX_total_Supply_price %>(<%= CTX_request %>)</td>
	<td><%= FormatNumber(ojumunmaster.FOneItem.Fjumunsuplycash,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>"><%= CTX_total_consumer_price %>(<%= CTX_FIX %>)</td>
	<td><b><%= FormatNumber(ojumunmaster.FOneItem.Ftotalsellcash,0) %></b></td>
	<td bgcolor="<%= adminColor("tabletop") %>"><%= CTX_total_Supply_price %>(<%= CTX_FIX %>)</td>
	<td><b><%= FormatNumber(ojumunmaster.FOneItem.Ftotalsuplycash,0) %></b></td>
</tr>
<!--
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">주문브랜드</td>
	<td colspan="3"><textarea class="textarea" cols="80" rows="3"><%= ojumunmaster.FOneItem.FBrandList %></textarea></td>
</tr>
-->
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>"><%= CTX_Requests %></td>
	<td colspan="3"><textarea class="textarea" name="comment" cols="80" rows="6"><%= ojumunmaster.FOneItem.FComment %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="4">
		* 5일내 출고 : 업체 배송 상품 (물류센터로 입고 되는대로 매장으로 발송 해드리겠습니다.) <br>
		* 재고 부족 : 물류센터 재고 부족으로 인해 업체로 발주가 들어가 있는 상태입니다. <br>
					2~3일 내로 입고 될 수 있는 상품 입니다. 따로 보내드리지 않으며, <B>다음 주문시 추가(재주문)</B>해 주셔야 합니다.<br>
		* 일시품절 : 업체 재고부족으로 인해 재생산중인 상품입니다.(단기간 내에 입고 되기 어려운 상품입니다.)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="4" align="center">
		<input type="button" class="button" value="<%= CTX_Revise %>" onclick="ModiMaster(frmMaster)">
		&nbsp;
		<input type="button" class="button" value="<%= CTX_Delete %>" onclick="DelMaster(frmMaster)">
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

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="11" align="right">
	총건수:  <%= ojumundetail.FResultCount %>
	&nbsp;
	<input type="button" class="button" value="<%= CTX_Add_new_items %>" onclick="AddItems(frmMaster)">
	</td>
</tr>
<tr bgcolor="#DDDDFF" align=center>
    <td width="50"><%= CTX_Image %></td>
	<td width="100"><%= CTX_Brand %></td>
	<td width="90"><%= CTX_Warehouse %>&nbsp;<%= CTX_Barcode %></td>
	<td><%= CTX_Description %><font color="blue">[<%= CTX_Description_Option %>]</font></td>
	<td width="60"><%= CTX_selling_price %></td>
	<td width="60"><%= CTX_Supply_price %></td>
	<td width="50"><%= CTX_request %>&nbsp;<%= CTX_quantity %></td>

	<% if isfixed then %>
		<td width="50"><%= CTX_FIX %>&nbsp;<%= CTX_quantity %></td>
		<td width="70"><%= CTX_total_Supply_price %></td>
		<td width="60"><%= CTX_Note %></td>
	<% else %>
		<td width="40"><%= CTX_Revise %></td>
		<td width="40"><%= CTX_Delete %></td>
		<td width="60"><%= CTX_Note %></td>
	<% end if %>

	<td width="70"><%= CTX_Invoice_Number %></td>
</tr>
<% for i=0 to ojumundetail.FResultCount-1 %>
<%
selltotal  = selltotal + ojumundetail.FItemList(i).FSellcash * ojumundetail.FItemList(i).Fbaljuitemno
suplytotal = suplytotal + ojumundetail.FItemList(i).FSuplycash * ojumundetail.FItemList(i).Fbaljuitemno
%>
<form name="frmBuyPrc_<%= i %>" method="post" action="common_shopjumun_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="masteridx" value="<%= idx %>">
<input type="hidden" name="itemgubun" value="<%= ojumundetail.FItemList(i).FItemGubun %>">
<input type="hidden" name="itemid" value="<%= ojumundetail.FItemList(i).FItemID %>">
<input type="hidden" name="itemoption" value="<%= ojumundetail.FItemList(i).Fitemoption %>">
<input type="hidden" name="desingerid" value="<%= ojumundetail.FItemList(i).Fmakerid %>">
<input type="hidden" name="sellcash" value="<%= ojumundetail.FItemList(i).FSellcash %>">
<input type="hidden" name="suplycash" value="<%= ojumundetail.FItemList(i).FSuplycash %>">
<% if IS_HIDE_BUYCASH = True then %>
<input type="hidden" name="buycash" value="-1">
<% else %>
<input type="hidden" name="buycash" value="<%= ojumundetail.FItemList(i).Fbuycash %>">
<% end if %>
<tr align="center" bgcolor="#FFFFFF">
    <td><img src="<%= ojumundetail.FItemList(i).GetImageSmall %>" border="0" width="50" height="50" onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
	<td><%= ojumundetail.FItemList(i).Fmakerid %></td>
	<td><%= ojumundetail.FItemList(i).FItemGubun %><%= CHKIIF(ojumundetail.FItemList(i).FItemID>=1000000,format00(8,ojumundetail.FItemList(i).FItemID),format00(6,ojumundetail.FItemList(i).FItemID)) %><%= ojumundetail.FItemList(i).Fitemoption %></td>
	<td align="left">
		<%= ojumundetail.FItemList(i).Fitemname %>
		<% if ojumundetail.FItemList(i).Fitemoption <> "0000" then %>
			<font color="blue">[<%= ojumundetail.FItemList(i).Fitemoptionname %>]</font>
		<% end if %>
	</td>
	<td align="right"><%= FormatNumber(ojumundetail.FItemList(i).Fsellcash,0) %></td>
	<td align="right"><%= FormatNumber(ojumundetail.FItemList(i).Fsuplycash,0) %></td>
	<td><input type="text" class="text" name="baljuitemno" value="<%= ojumundetail.FItemList(i).Fbaljuitemno %>"  size="3" maxlength="4"></td>
	<% if isfixed then %>
	<td><%= ojumundetail.FItemList(i).Frealitemno %></td>
	<td align="right"><%= FormatNumber(ojumundetail.FItemList(i).Fsuplycash * ojumundetail.FItemList(i).Frealitemno,0) %></td>
	<td><%= ojumundetail.FItemList(i).Fcomment %></td>
	<% else %>
	<td><input type=button value="<%= CTX_Revise %>" onclick="ModiThis(frmBuyPrc_<%= i %>)" class="button"></td>
	<td><input type=button value="<%= CTX_Delete %>" onclick="DelThis(frmBuyPrc_<%= i %>)" class="button"></td>
	<td></td>
	<% end if %>
	<td><%= ojumundetail.FItemList(i).Fboxsongjangno %></td>
</tr>
</form>
<% next %>

<% if (ojumundetail.FResultCount>0) then %>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan=5><%= CTX_total %></td>
	<td align=right><%= formatNumber(suplytotal,0) %></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
<% end if %>
</table>
<form name="frmadd" method=post action="common_shopjumun_process.asp">
	<input type="hidden" name="mode" value="shopjumunitemaddarr">
	<input type="hidden" name="masteridx" value="<%= idx %>">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="sellcasharr" value="">
	<input type="hidden" name="suplycasharr" value="">
	<input type="hidden" name="buycasharr" value="">
	<input type="hidden" name="itemnoarr" value="">
</form>

<script type='text/javascript'>

	if (jumunwait!=true){
		alert('주문접수 이후에는 수정하실 수 없습니다.');
	}else if (rejumun==true){
		alert('재작성된 주문서는 수정하실 수 없습니다.');
	}

</script>

<%
set ojumunmaster = Nothing
set ojumundetail = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
