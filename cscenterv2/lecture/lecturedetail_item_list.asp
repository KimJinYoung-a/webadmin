<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 강좌 cs센터 상품리스트
' Hieditor : 2015.05.27 이상구 생성
'			 2016.10.06 한용민 수정
'###########################################################
%>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/lecture/lecturecls.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/order/ordergiftcls.asp"-->
<%
dim orderserial, onimage, research, i, ix
	orderserial = RequestCheckvar(request("orderserial"),16)
	onimage     = RequestCheckvar(request("onimage"),2)
	research    = RequestCheckvar(request("research"),2)

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

''사은품정보 추가 : 상품 발주 시 생성됨.
dim oGift
set oGift = new COrderGift

if (oordermaster.FOneItem.Fipkumdiv>1) and (oordermaster.FOneItem.Fjumundiv<>9) then
    oGift.FRectOrderSerial = orderserial
    oGift.GetOneOrderGiftlist
end if
%>
<script type="text/javascript">

function popOrderDetailEdit(idx){
	alert('[차후 작업 예정] 주문 수정 & 저장하는 팝업 뜨는 자리');
	return;

	var popwin = window.open('/common/orderdetailedit.asp?idx=' + idx,'orderdetailedit','width=500,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popOrderDetailEditOption(idx){
	alert('[차후 작업 예정] 주문 수정 & 저장하는 팝업 뜨는 자리');
	return;

	var popwin = window.open('/cscenter/ordermaster/orderdetail_editoption.asp?idx=' + idx,'orderdetaileditoption','width=500,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popSimpleBrandInfo(makerid){
	var popwin = window.open('/common/popsimpleBrandInfo.asp?makerid=' + makerid,'popsimpleBrandInfo','width=500,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}
function popnamephone_update(idx){
	var popwin = window.open('/cscenterv2/cs/pop_namephone_update.asp?idx=' + idx,'pop_namephone_update','width=500,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<body topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>

<table width="100%" border="0" cellspacing=0 cellpadding=1 class=a bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="orderserial" value="<%= orderserial %>">
<input type="hidden" name="research" value="on">

<tr align="center" bgcolor="<%= adminColor("topbar") %>" style="padding:2">
	<td width="30">구분</td>
	<td width="80">CODE</td>
  	<td width="55">이미지</td>
    <td width="120">브랜드ID</td>
	<td>강좌명<font color="blue">[시간]</font></td>
	<td width="50">성명</td>
	<td width="100">연락처</td>
	<td width="70">강좌료</td>
	<td width="70">재료비</td>
	<td width="70">합계</td>
	<td width="70">할인액</td>
	<td width="70">기환불</td>
</tr>
<tr>
	<td height="1" colspan="12" bgcolor="#BABABA"></td>
</tr>

<% for ix=0 to oorderdetail.FResultCount-1 %>
<% if oorderdetail.FItemList(ix).Fitemid <>0 then %>

<% if oorderdetail.FItemList(ix).FCancelyn ="Y" then %>
<tr align="center" height="25" bgcolor="#EEEEEE" class="gray">
<% else %>
<tr align="center" height="25" bgcolor="#ffffff">
<% end if %>

    <td><font color="<%= oorderdetail.FItemList(ix).CancelStateColor %>"><%= oorderdetail.FItemList(ix).CancelStateStr %></font></td>
	<td>
	    <% if oorderdetail.FItemList(ix).Fisupchebeasong="Y" then %>
	    <a href="javascript:popOrderDetailEdit('<%= oorderdetail.FItemList(ix).Fidx %>');"><font color="red"><%= oorderdetail.FItemList(ix).Fitemid %><br>(업체)</font></a>
        <% else %>
        <a href="javascript:popOrderDetailEdit('<%= oorderdetail.FItemList(ix).Fidx %>');"><%= oorderdetail.FItemList(ix).Fitemid %></a>
        <% end if %>
    </td>
    <td align="center">
        <% if onimage="on" then %>
        <a href="http://www.thefingers.co.kr/lecture/lecturedetail.asp?lec_idx=<%= oorderdetail.FItemList(ix).Fitemid %>" target="_blank"><img src="<%= oorderdetail.FItemList(ix).FSmallImage %>" width="50" height="50" border="0"></a>
        <% else %>
        &nbsp;
        <% end if %>
    </td>
    <td>
        <acronym title="<%= oorderdetail.FItemList(ix).Fmakerid %>"><%= Left(oorderdetail.FItemList(ix).Fmakerid,12) %></acronym>
    </td>
	<td align="left">
	    <acronym title="<%= oorderdetail.FItemList(ix).FItemName %>"><%= Left(oorderdetail.FItemList(ix).FItemName,35) %></acronym>
    	    <br>
    	    <a href="javascript:popOrderDetailEditOption('<%=oorderdetail.FItemList(ix).Fidx%>');"><font color="blue"><%= oorderdetail.FItemList(ix).FItemoptionName %></font></a>
	    <% if oorderdetail.FItemList(ix).IsRequireDetailExistsItem then %>
	    	<br>
	    	<a href="javascript:EditRequireDetail('<%= orderserial %>','<%= oorderdetail.FItemList(ix).Fidx%>')"><font color="red">[주문제작상품]</font>
	    	<br>
	    	<%= db2html(oorderdetail.FItemList(ix).getRequireDetailHtml) %>
	    	</a>
	    <% end if %>
	</td>
	<td><%= db2html(oorderdetail.FItemList(ix).Fentryname) %></td>
<% If oorderdetail.FResultCount-1 >= 1 and ix <> 0 Then %>
	<td onclick="javascript:popnamephone_update('<%=oorderdetail.FItemList(ix).Fidx%>');" style="cursor:pointer;"><%= db2html(oorderdetail.FItemList(ix).Fentryhp) %></td>
<% Else %>
	<td><%= db2html(oorderdetail.FItemList(ix).Fentryhp) %></td>
<% End If %>

	<td align="right">
		<% if oorderdetail.FItemList(ix).Fmatinclude_yn = "C" then %>
		<%= FormatNumber((oorderdetail.FItemList(ix).FcouponNotAsigncost - oorderdetail.FItemList(ix).FmatcostAdded),0) %>
		<% else %>
		<%= FormatNumber(oorderdetail.FItemList(ix).FcouponNotAsigncost,0) %>
		<% end if %>
	</td>
	<td align="right">
		<%= FormatNumber(oorderdetail.FItemList(ix).FmatcostAdded,0) %><br>
		<% if oorderdetail.FItemList(ix).Fmatinclude_yn = "C" then %>
		[선납]
		<% else %>
		[현장]
		<% end if %>
	</td>
	<td align="right"><%= FormatNumber(oorderdetail.FItemList(ix).FcouponNotAsigncost,0) %></td>
   	<td align="right">
   		<% if (oorderdetail.FItemList(ix).FCancelyn="Y") then %>
   			<font color="gray"><%= FormatNumber(oorderdetail.FItemList(ix).FreducedPrice,0) %></font>
   		<% elseif oorderdetail.FItemList(ix).FItemNo < 1 then %>
			<font color="red"><%= FormatNumber(oorderdetail.FItemList(ix).FreducedPrice,0) %></font>
   		<% else %>
       	    <% if oorderdetail.FItemList(ix).Fissailitem="Y" then %>
           	    <span title="세일상품" style="cursor:hand"><font color="red"><b><%= FormatNumber(oorderdetail.FItemList(ix).FreducedPrice,0) %></b></font></span>
       	    <% elseif oorderdetail.FItemList(ix).Fissailitem="P" then %>
           	    <span title="플러스세일상품" style="cursor:hand"><font color="purple"><%= FormatNumber(oorderdetail.FItemList(ix).FreducedPrice,0) %></font></span>
       	    <% elseif oorderdetail.FItemList(ix).IsBonusCouponDiscountItem then %>
           	    <span title="비율보너스쿠폰할인상품" style="cursor:hand">
           	    <font color="blue">
           	        <%= FormatNumber(oorderdetail.FItemList(ix).FreducedPrice,0) %>
           	    </font>
           	    </span>
       	    <% elseif oorderdetail.FItemList(ix).IsItemCouponDiscountItem then %>
           	    <span title="상품보너스쿠폰할인상품" style="cursor:hand"><font color="green"><b><%= FormatNumber(oorderdetail.FItemList(ix).FreducedPrice,0) %></b></font></span>
       	    <% else %>
           	    <span title="정상가격" style="cursor:hand"><font color="#000000"><%= FormatNumber(oorderdetail.FItemList(ix).FreducedPrice,0) %></font></span>
       	    <% end if %>
   		<% end if %>
	</td>
    <td align="right">
    	<span title="<%= oorderdetail.FItemList(ix).GetRefundStateName %>" style="cursor:hand"><font color="<%= oorderdetail.FItemList(ix).GetRefundStateColor %>"><%= FormatNumber(oorderdetail.FItemList(ix).GetRefundPrice,0) %></font></span>
	</td>
</tr>
<tr>
	<td height="1" colspan="12" bgcolor="#BABABA"></td>
</tr>
<% end if %>
<% next %>

<!--                <%= "CNT=" & oGift.FResultCount %>	-->
<% for i=0 to oGift.FResultCount -1 %>
<tr align="left" height="25" bgcolor="#ffffff">
	<td colspan="13">
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
	<td colspan="12" height="1" bgcolor="#BABABA"></td>
</tr>
<% next %>
</form>
</table>

<script type="text/javascript">

function popDeliveryTrace(traceUrl, songjangNo){
	var f = document.popForm;
	f.traceUrl.value	= traceUrl;
	f.songjangNo.value	= songjangNo;
	f.submit();
}

</script>
<form name="popForm" action="popDeliveryTrace.asp" target="_blank">
<input type="hidden" name="traceUrl">
<input type="hidden" name="songjangNo">
</form>

<%
set oGift = Nothing
set oordermaster = Nothing
set oorderdetail = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->