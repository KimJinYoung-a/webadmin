<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 주문상세
' Hieditor : 이상구 생성
'			 2018.06.05 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/3pl/orderCls.asp"-->
<!-- #include virtual="/lib/classes/3pl/orderGiftCls.asp"-->

<%
dim orderserial, onimage, research, i, ix
dim itemcostcomment, itemcostcolor
	orderserial = requestcheckvar(request("orderserial"),32)
	onimage     = requestcheckvar(request("onimage"),10)
	research    = requestcheckvar(request("research"),2)

if (onimage = "") and (research="") then  onimage = "on"

dim oordermaster, oorderdetail
set oordermaster = new COrderMaster
	oordermaster.FRectOrderSerial = orderserial
	oordermaster.QuickSearchOrderMaster

	if (oordermaster.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
	    oordermaster.FRectOldOrder = "on"
	    oordermaster.QuickSearchOrderMaster
	end if

set oorderdetail = new COrderMaster
	oorderdetail.FRectOrderSerial = orderserial
	oorderdetail.QuickSearchOrderDetail

	if (oorderdetail.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
	    oorderdetail.FRectOldOrder = "on"
	    oorderdetail.QuickSearchOrderDetail
	end if

''사은품정보 추가 : 상품 출고지시 시 생성됨.
dim oGift
set oGift = new COrderGift
if (oordermaster.FOneItem.Fipkumdiv>1) and (oordermaster.FOneItem.Fjumundiv<>9) then
    oGift.FRectOrderSerial = orderserial
    oGift.GetOneOrderGiftlist
end if

function drawPrice(v)
	'과거 데이타와의 호환성때문에 값이 있을때만 표시하는 함수를 쓴다.
	dim result

	result = ""
	if (v <> 0) then
		result = " : " & FormatNumber(v, 0) & " 원"
	end if

	drawPrice = result
end function

dim IsShowSetStateBtn : IsShowSetStateBtn = False
dim IsShowSetStateBtnAvail : IsShowSetStateBtnAvail = False

if (oordermaster.FOneItem.Fcancelyn = "N") and (oordermaster.FOneItem.Fipkumdiv >= "4") and (InStr(oordermaster.FOneItem.Fjumundiv, "4679") = 0) then
	IsShowSetStateBtnAvail = True
end if

%>
<link rel="stylesheet" href="/cscenter/css/cs.css" type="text/css">
<script type="text/javascript">

function popOrderDetailEdit(idx){
	var popwin = window.open('/common/orderdetailedit.asp?idx=' + idx,'orderdetailedit','width=500,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popOrderDetailEditOption(idx){
	// var popwin = window.open('/cscenter/ordermaster/orderdetail_editoption.asp?idx=' + idx,'orderdetaileditoption','width=500,height=480,scrollbars=yes,resizable=yes');
	var popwin = window.open('/cscenter/ordermaster/orderdetail_simple_editoption.asp?idx=' + idx,'orderdetaileditoption','width=900,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popOrderDetailEditItem(idx){
	var popwin = window.open('/cscenter/ordermaster/orderdetail_edititem.asp?idx=' + idx,'popOrderDetailEditItem','width=600,height=850,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popSimpleBrandInfo(makerid){
	var popwin = window.open('/common/popsimpleBrandInfo.asp?makerid=' + makerid,'popsimpleBrandInfo','width=500,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ReloadThisPage() {
	var frm = document.frm;

	if (frm.onimage.value == "on") {
		frm.onimage.value = "off";
	} else {
		frm.onimage.value = "on";
	}

	frm.submit();
}

function pojangview(orderserial){
	var pojangview = window.open('/cscenter/pack/pojang_view.asp?orderserial='+orderserial,'pojangview','width=1024,height=768,scrollbars=yes,resizable=yes');
	pojangview.focus();
}

<% if IsShowSetStateBtnAvail then %>
function jsSetCurrState() {
	if (confirm("모든 업체배송 상품의 상태를 주문통보로 전환합니다.\n\n진행하시겠습니까?") == true) {
		var frm = document.actFrm;
		frm.mode.value = "modistate2";
		frm.submit();
	}
}
<% end if %>

</script>

<table width="100%" border="0" cellspacing=0 cellpadding=0 class=a bgcolor="FFFFFF">
    <tr>
        <td>
            <table width="100%" border="0" cellspacing=0 cellpadding=0 class=a bgcolor="FFFFFF">
            <form name="frm" method="get" action="">
            <input type="hidden" name="orderserial" value="<%= orderserial %>">
            <input type="hidden" name="research" value="on">
            <input type="hidden" name="onimage" value="<%= onimage %>">
            <tr align="center" height="0">
                <td width="30"></td>
                <td width="50"></td>
            	<td width="100"></td>
				<td width="80"></td>
               	<td width="50"></td>
            	<td width="200" align="left"></td>
            	<td align="left"></td>
            	<td width="30"></td>
				<% if (C_InspectorUser = False) then %>
            	<td width="60" align="right"></td>
            	<td width="60" align="right"></td>
            	<td width="60" align="right"></td>
				<% end if %>
            	<td width="60" align="right"></td>
				<td width="60" align="right"></td>
				<td width="60" align="right"></td>
				<td width="60" align="right"></td>
            	<td width="70"></td>
            	<td width="125"></td>
            </tr>
<% for ix=0 to oorderdetail.FResultCount-1 %>
	<%
	'/배송비
	if oorderdetail.FItemList(ix).Fitemid = 0 then
	%>
		<% if oorderdetail.FItemList(ix).FCancelyn ="Y" then %>
			<tr align="center" height="25" bgcolor="#EEEEEE" class="gray">
        <% else %>
			<tr align="center" height="25">
        <% end if %>
                <td></td>
                <td></td>
            	<td><%= oorderdetail.FItemList(ix).Fprdcode %></td>
				<td><%= oorderdetail.FItemList(ix).Fitemid %></td>
               	<td></td>
            	<td align="left"><%= oorderdetail.FItemList(ix).FMakerid %></td>
            	<td align="left">
            		배송비
            		<font color="blue">[<%= oorderdetail.BeasongCD2Name(oorderdetail.FItemList(ix).Fitemoption) %><% oorderdetail.BeasongOptionString(oorderdetail.FItemList(ix).Fitemoptionname) %>]</font>
            	</td>
            	<td><%= oorderdetail.FItemList(ix).Fitemno %></td>

				<% if (C_InspectorUser = False) then %>
				<!-- 소비자가 -->
                <td align="right"></td>

				<!-- 할인가 -->
            	<td align="right" style="padding-right:5px">
                	<% if (Not oorderdetail.FItemList(ix).IsOldJumun) then %>
                		<%= FormatNumber(oorderdetail.FItemList(ix).GetSalePrice,0) %>
                	<% else %>
                		----
                	<% end if %>
            	</td>

				<!-- 상품쿠폰적용가 -->
            	<td align="right" style="padding-right:5px">
                	<span title="배송비쿠폰 : <%= FormatNumber(oorderdetail.FItemList(ix).GetItemCouponDiscountPrice,0) %>원" style="cursor:hand">
                	<font color="<%= oorderdetail.FItemList(ix).GetItemCouponColor %>">
                		<%= FormatNumber(oorderdetail.FItemList(ix).GetItemCouponPrice,0) %>
                	</font>
                	</span>
            	</td>
				<% end if %>

				<!-- 보너스쿠폰적용가 -->
            	<td align="right" style="padding-right:5px">
                	<% if oorderdetail.FItemList(ix).FItemNo < 1 then %>
                		<font color="red"><%= FormatNumber(oorderdetail.FItemList(ix).GetBonusCouponPrice,0) %></font>
                	<% else %>
                    	<span title="<%= oorderdetail.FItemList(ix).GetBonusCouponText %>" style="cursor:hand">
                    	<font color="<%= oorderdetail.FItemList(ix).GetBonusCouponColor %>">
                    		<%= FormatNumber(oorderdetail.FItemList(ix).GetBonusCouponPrice,0) %>
                    	</font>
                    	</span>
                    <% end if %>
            	</td>

				<!-- 기타할인적용가 -->
				<td align="right" style="padding-right:5px">
                    <span title="<%= oorderdetail.FItemList(ix).GetEtcDiscountText %>" style="cursor:hand">
                    	<font color="<%= oorderdetail.FItemList(ix).GetEtcDiscountColor %>">
                    		<%= FormatNumber(oorderdetail.FItemList(ix).GetEtcDiscountPrice,0) %>
                    	</font>
                    </span>
				</td>

            	<td align="right"></td>

            	<td align="right" style="padding-right:5px">
                	<%= FormatNumber(oorderdetail.FItemList(ix).Fbuycash,0) %>
            	</td

				<td align="right"></td>
            	<td></td>
            </tr>
            <tr>
        		<td height="1" colspan="17" bgcolor="#BABABA"></td>
        	</tr>
	<% end if %>
<% next %>
<% for ix=0 to oorderdetail.FResultCount-1 %>
	<%
	'/포장비
	if oorderdetail.FItemList(ix).Fitemid = 100 then
	%>
		<% if oorderdetail.FItemList(ix).FCancelyn ="Y" then %>
			<tr align="center" height="25" bgcolor="#EEEEEE" class="gray">
        <% else %>
			<tr align="center" height="25">
        <% end if %>
                <td></td>
                <td></td>
				<td><%= oorderdetail.FItemList(ix).Fprdcode %></td>
            	<td><%= oorderdetail.FItemList(ix).Fitemid %></td>
               	<td></td>
            	<td align="left"><%= oorderdetail.FItemList(ix).FMakerid %></td>
            	<td align="left">
            		포장비
            		<a href="" onclick="pojangview('<%= orderserial %>'); return false;"><font color="blue">[포장상세보기]</font></a>
            	</td>
            	<td><%= oorderdetail.FItemList(ix).Fitemno %></td>

				<% if (C_InspectorUser = False) then %>
				<!-- 소비자가 -->
                <td align="right"></td>

				<!-- 할인가 -->
            	<td align="right" style="padding-right:5px">
                	<% if (Not oorderdetail.FItemList(ix).IsOldJumun) then %>
                		<%= FormatNumber(oorderdetail.FItemList(ix).GetSalePrice,0) %>
                	<% else %>
                		----
                	<% end if %>
            	</td>

				<!-- 상품쿠폰적용가 -->
            	<td align="right" style="padding-right:5px">
                	<span title="배송비쿠폰 : <%= FormatNumber(oorderdetail.FItemList(ix).GetItemCouponDiscountPrice,0) %>원" style="cursor:hand">
                	<font color="<%= oorderdetail.FItemList(ix).GetItemCouponColor %>">
                		<%= FormatNumber(oorderdetail.FItemList(ix).GetItemCouponPrice,0) %>
                	</font>
                	</span>
            	</td>
				<% end if %>

				<!-- 보너스쿠폰적용가 -->
            	<td align="right" style="padding-right:5px">
                	<% if oorderdetail.FItemList(ix).FItemNo < 1 then %>
                		<font color="red"><%= FormatNumber(oorderdetail.FItemList(ix).GetBonusCouponPrice,0) %></font>
                	<% else %>
                    	<span title="<%= oorderdetail.FItemList(ix).GetBonusCouponText %>" style="cursor:hand">
                    	<font color="<%= oorderdetail.FItemList(ix).GetBonusCouponColor %>">
                    		<%= FormatNumber(oorderdetail.FItemList(ix).GetBonusCouponPrice,0) %>
                    	</font>
                    	</span>
                    <% end if %>
            	</td>

            	<td align="right"></td>
				<td align="right"></td>
				<td align="right"></td>
				<td align="right"></td>
            	<td></td>
            </tr>
            <tr>
        		<td height="1" colspan="17" bgcolor="#BABABA"></td>
        	</tr>
	<% end if %>
<% next %>
<% for ix=0 to oorderdetail.FResultCount-1 %>
	<% if oorderdetail.FItemList(ix).Fitemid <> 0 and oorderdetail.FItemList(ix).Fitemid <> 100 then %>
		<% if oorderdetail.FItemList(ix).FCancelyn ="Y" then %>
			<tr align="center" height="35" bgcolor="#EEEEEE" class="gray">
        <% else %>
			<tr align="center" height="35">
        <% end if %>
                <td><font color="<%= oorderdetail.FItemList(ix).CancelStateColor %>"><%= oorderdetail.FItemList(ix).CancelStateStr %></font></td>
                <td><font color="<%= oorderdetail.FItemList(ix).GetStateColor %>"><%= oorderdetail.FItemList(ix).GetStateName %></font></td>
            	<td>
            	    <% if oorderdetail.FItemList(ix).Fisupchebeasong="Y" then %>
            	    	<% if oorderdetail.FItemList(ix).fodlvfixday="G" then %>
            	    		<a href="javascript:popOrderDetailEdit('<%= oorderdetail.FItemList(ix).Fidx %>');"><font color="red"><%= oorderdetail.FItemList(ix).Fprdcode %><br>(해외직구)</font></a>
            	    	<% else %>
							<a href="javascript:popOrderDetailEdit('<%= oorderdetail.FItemList(ix).Fidx %>');"><font color="red"><%= oorderdetail.FItemList(ix).Fprdcode %><br>(업체)</font></a>
						<% end if %>
					<% elseif oorderdetail.FItemList(ix).Fisupchebeasong="N" and InStr(",2,9,7,", CStr(oorderdetail.FItemList(ix).FodlvType)) > 0 then %>
						<a href="javascript:popOrderDetailEdit('<%= oorderdetail.FItemList(ix).Fidx %>');"><font color="red"><%= oorderdetail.FItemList(ix).Fprdcode %><br>(해외)</font></a>
                    <% else %>
						<a href="javascript:popOrderDetailEdit('<%= oorderdetail.FItemList(ix).Fidx %>');"><%= oorderdetail.FItemList(ix).Fprdcode %></a>
                    <% end if %>
                </td>
            	<td>
            	    <% if oorderdetail.FItemList(ix).Fisupchebeasong="Y" then %>
            	    	<% if oorderdetail.FItemList(ix).fodlvfixday="G" then %>
            	    		<a href="javascript:popOrderDetailEdit('<%= oorderdetail.FItemList(ix).Fidx %>');"><font color="red"><%= oorderdetail.FItemList(ix).Fitemid %><br>(해외직구)</font></a>
            	    	<% else %>
							<a href="javascript:popOrderDetailEdit('<%= oorderdetail.FItemList(ix).Fidx %>');"><font color="red"><%= oorderdetail.FItemList(ix).Fitemid %><br>(업체)</font></a>
						<% end if %>
					<% elseif oorderdetail.FItemList(ix).Fisupchebeasong="N" and InStr(",2,9,7,", CStr(oorderdetail.FItemList(ix).FodlvType)) > 0 then %>
						<a href="javascript:popOrderDetailEdit('<%= oorderdetail.FItemList(ix).Fidx %>');"><font color="red"><%= oorderdetail.FItemList(ix).Fitemid %><br>(해외)</font></a>
                    <% else %>
						<a href="javascript:popOrderDetailEdit('<%= oorderdetail.FItemList(ix).Fidx %>');"><%= oorderdetail.FItemList(ix).Fitemid %></a>
                    <% end if %>
                </td>
                <td align="center">
                    <% if onimage="on" then %>
                    <a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oorderdetail.FItemList(ix).Fitemid %>" target="_blank"><img src="<%= oorderdetail.FItemList(ix).FSmallImage %>" width="50" height="50" border="0"></a>
                    <% else %>
                    &nbsp;
                    <% end if %>
                </td>
                <td align="left">
                    <a href="javascript:popSimpleBrandInfo('<%= oorderdetail.FItemList(ix).Fmakerid %>');">
                    <acronym title="<%= oorderdetail.FItemList(ix).Fmakerid %>"><%= Left(oorderdetail.FItemList(ix).Fmakerid,12) %></acronym>
                    </a>
                </td>
            	<td align="left">
            	    <table width="100%" border="0" cellspacing=0 cellpadding=2 class=a style="table-layout:fixed">
            	    <tr>
            	    	<td>
	                	    <acronym title="<%= oorderdetail.FItemList(ix).FItemName %>"><a href="javascript:popOrderDetailEditItem('<%=oorderdetail.FItemList(ix).Fidx%>')"><%= oorderdetail.FItemList(ix).FItemName %></a></acronym><br>
							<% if oorderdetail.FItemList(ix).FItemoption<>"0000" then %>
		                	    <a href="javascript:popOrderDetailEditOption('<%=oorderdetail.FItemList(ix).Fidx%>');"><font color="blue"><%= oorderdetail.FItemList(ix).FItemoptionName %></font></a>
	                	    <% end if %>
	                	    <% if oorderdetail.FItemList(ix).IsRequireDetailExistsItem then %>
	                	    	<br>
	                	    	<a href="javascript:EditRequireDetail('<%= orderserial %>','<%= oorderdetail.FItemList(ix).Fidx%>')"><font color="red">[주문제작상품]</font>
	                	    	<br>
	                	    	<%= db2html(oorderdetail.FItemList(ix).getRequireDetailHtml) %>
	                	    	</a>
	                	    <% end if %>
            	    	</td>
            	    </tr>
            	    </table>
            	</td>

            	<% if oorderdetail.FItemList(ix).FItemNo > 1 then %>
            	<td><b><font color="blue"><%= oorderdetail.FItemList(ix).FItemNo %></font></b></td>
            	<% elseif oorderdetail.FItemList(ix).FItemNo < 1 then %>
            	<td><b><font color="red"><%= oorderdetail.FItemList(ix).FItemNo %></font></b></td>
            	<% else %>
            	<td><font color="blue"><%= oorderdetail.FItemList(ix).FItemNo %></font></td>
            	<% end if %>

				<% if (C_InspectorUser = False) then %>
				<!-- 소비자가 -->
                <td align="right" style="padding-right:5px">
                	<% if (Not oorderdetail.FItemList(ix).IsOldJumun) then %>
                    	<font color="<%= oorderdetail.FItemList(ix).GetOrgItemCostColor %>">
                    		<%= FormatNumber(oorderdetail.FItemList(ix).GetOrgItemCostPrice,0) %>
                    	</font>
                	<% else %>
                		----
                	<% end if %>
                </td>

				<!-- 할인가 -->
            	<td align="right" style="padding-right:5px">
                	<% if (Not oorderdetail.FItemList(ix).IsOldJumun) then %>
                    	<span title="<%= oorderdetail.FItemList(ix).GetSaleText %>" style="cursor:hand">
                    	<font color="<%= oorderdetail.FItemList(ix).GetSaleColor %>">
                    		<%= FormatNumber(oorderdetail.FItemList(ix).GetSalePrice,0) %>
                    	</font>
                    	</span>
                	<% else %>
                		----
                	<% end if %>
            	</td>

				<!-- 상품쿠폰적용가 -->
            	<td align="right" style="padding-right:5px">
                	<span title="<%= oorderdetail.FItemList(ix).GetItemCouponText %>" style="cursor:hand">
                	<font color="<%= oorderdetail.FItemList(ix).GetItemCouponColor %>">
                		<%= FormatNumber(oorderdetail.FItemList(ix).GetItemCouponPrice,0) %>
                	</font>
                	</span>
            	</td>
				<% end if %>

				<!-- 보너스쿠폰적용가 -->
            	<td align="right" style="padding-right:5px">
                	<% if oorderdetail.FItemList(ix).FItemNo < 1 then %>
                		<font color="red"><%= FormatNumber(oorderdetail.FItemList(ix).GetBonusCouponPrice,0) %></font>
                	<% else %>
                    	<span title="<%= oorderdetail.FItemList(ix).GetBonusCouponText %>" style="cursor:hand">
                    	<font color="<%= oorderdetail.FItemList(ix).GetBonusCouponColor %>">
                    		<%= FormatNumber(oorderdetail.FItemList(ix).GetBonusCouponPrice,0) %>
                    	</font>
                    	</span>
                    <% end if %>
            	</td>

				<!-- 기타할인적용가 -->
				<td align="right" style="padding-right:5px">
                    <span title="<%= oorderdetail.FItemList(ix).GetEtcDiscountText %>" style="cursor:hand">
                    	<font color="<%= oorderdetail.FItemList(ix).GetEtcDiscountColor %>">
                    		<%= FormatNumber(oorderdetail.FItemList(ix).GetEtcDiscountPrice,0) %>
                    	</font>
                    </span>
				</td>

				<!-- 구매마일리지 -->
            	<td align="right" style="padding-right:5px">
                	<%= FormatNumber(oorderdetail.FItemList(ix).Fmileage,0) %>
            	</td>

				<!-- 매입가 -->
            	<td align="right" style="padding-right:5px">
                	<%= FormatNumber(oorderdetail.FItemList(ix).Fbuycash,0) %>
            	</td>

				<td>
					<acronym title="<%= oordermaster.FOneItem.Fbaljudate %>"><%= Left(oordermaster.FOneItem.Fbaljudate,10) %></acronym><br>
					<acronym title="<%= oorderdetail.FItemList(ix).Fupcheconfirmdate %>"><%= Left(oorderdetail.FItemList(ix).Fupcheconfirmdate,10) %></acronym>
					<%
					''주문통보 전환버튼
					if IsShowSetStateBtnAvail and Not IsShowSetStateBtn then
						if (oorderdetail.FItemList(ix).Fisupchebeasong="Y") and (oorderdetail.FItemList(ix).FCancelyn <> "Y") and IsNull(oordermaster.FOneItem.Fbaljudate) and IsNull(oorderdetail.FItemList(ix).Fupcheconfirmdate) then
							IsShowSetStateBtn = True
							if oorderdetail.FItemList(ix).Fcurrstate = "0" then
					%>
					<input type="button" class="button" value="통보" onClick="jsSetCurrState()">
					<%
							end if
						end if
					end if
					%>
				</td>
            	<td>
            		<acronym title="<%= oorderdetail.FItemList(ix).Fbeasongdate %>"><%= Left(oorderdetail.FItemList(ix).Fbeasongdate,10) %></acronym><br>
            	    <%= oorderdetail.FItemList(ix).Fsongjangdivname %><br>
            		<% if (oorderdetail.FItemList(ix).FsongjangDiv="24") then %>
            		<a href="javascript:popDeliveryTrace('<%= oorderdetail.FItemList(ix).Ffindurl %>','<%= oorderdetail.FItemList(ix).Fsongjangno %>');"><%= oorderdetail.FItemList(ix).Fsongjangno %></a>
            	    <% else %>
            	    <a target="_blank" href="<%= oorderdetail.FItemList(ix).Ffindurl + oorderdetail.FItemList(ix).Fsongjangno %>"><%= oorderdetail.FItemList(ix).Fsongjangno %></a>
            	    <% end if %>
            	</td>
            </tr>
            <tr>
        		<td height="1" colspan="17" bgcolor="#BABABA"></td>
        	</tr>
	<% end if %>
<% next %>

<!--                <%= "CNT=" & oGift.FResultCount %>	-->
            <% for i=0 to oGift.FResultCount -1 %>
            <tr align="left" height="25">
            	<td colspan="17">
                    <table width="100%" border=0 cellspacing=0 cellpadding=0 class=a bgcolor="FFFFFF">
                	<tr>
	                    <td align="left">
	                    	<font color="blue">사은품</font>
	                    	&nbsp;&nbsp;
	                        <% if oGift.FItemList(i).Fisupchebeasong="Y" then %>
	                        <font color="red">업체</font>
	                        <% else %>
	                        <font color="blue">텐배</font>
	                        <% end if %>

	                        &nbsp;&nbsp;

		                    <% if (oGift.FItemList(i).Fevt_code<>0) then %>
		                    <a target="_blank" href="http://www.10x10.co.kr/event/eventmain.asp?eventid=<%= oGift.FItemList(i).Fevt_code %>"><font color="blue">[<%= oGift.FItemList(i).Fevt_code %>-<%= oGift.FItemList(i).Fgift_code %>]<%= oGift.FItemList(i).Fevt_name %></font></a>
		                    <% else %>
		                    [0-<%= oGift.FItemList(i).Fgift_code %>]<%= oGift.FItemList(i).Fgift_name %>
		                    <% end if %>

		                    &nbsp;&nbsp;
	                    	<%= oGift.FItemList(i).GetEventConditionStr %>
	                    </td>
	                </tr>
		            </table>
	        	</td>
            </tr>
            <tr>
        		<td colspan="17" height="1" bgcolor="#BABABA"></td>
        	</tr>
            <% next %>
            </form>
            </table>
        </td>
    </tr>
</table>
<script type="text/javascript">

function popDeliveryTrace(traceUrl, songjangNo){
	var f = document.popForm;
	f.traceUrl.value	= traceUrl;
	f.songjangNo.value	= songjangNo;
	f.submit();
}

</script>
<form name="popForm" action="popDeliveryTrace.asp" target="_blank" style="margin:0px;">
<input type="hidden" name="traceUrl">
<input type="hidden" name="songjangNo">
</form>

<form name="actFrm" action="orderdetail_process.asp">
	<input type="hidden" name="orderserial" value="<%= orderserial %>">
	<input type="hidden" name="mode">
</form>

<%
set oGift = Nothing
set oordermaster = Nothing
set oorderdetail = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
