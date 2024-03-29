<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 매장 고객센터
' Hieditor : 2012.03.20 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/shopcscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/shopcscenter_order_cls.asp"-->
<!-- #include virtual="/admin/offshop/shopcscenter/cscenter_Function_off.asp"-->
<%
dim masteridx, onimage, research ,itemcostcomment, itemcostcolor ,oordermaster, oorderdetail
dim  i, ix
	masteridx = requestCheckVar(request("masteridx"),10)
	onimage     = requestCheckVar(request("onimage"),2)
	research    = requestCheckVar(request("research"),2)

	if (onimage = "") and (research="") then  onimage = "on"

set oordermaster = new COrder
	oordermaster.FRectmasteridx = masteridx
	oordermaster.fQuickSearchOrderMaster

set oorderdetail = new COrder
	oorderdetail.FRectmasteridx = masteridx
	oorderdetail.fQuickSearchOrderDetail
%>

<script language='javascript'>

function popSimpleBrandInfo(makerid){
	var popwin = window.open('/common/popsimpleBrandInfo.asp?makerid=' + makerid,'popsimpleBrandInfo','width=500,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<table width="100%" border=0 cellspacing=0 cellpadding=0 class=a bgcolor="FFFFFF">
<tr>
    <td>
        <table width="100%" border="0" cellspacing=0 cellpadding=2 class=a bgcolor="FFFFFF">
            <form name="frm" method="get" action="">
            <input type="hidden" name="masteridx" value="<%= masteridx %>">
            <input type="hidden" name="research" value="on">

            <% for ix=0 to oorderdetail.FResultCount-1 %>

            <% if oorderdetail.FItemList(ix).Fitemid <>0 then %>
	            <% if oorderdetail.FItemList(ix).FCancelyn ="Y" then %>
					<tr align="center" height="25" bgcolor="#EEEEEE" class="gray">
	            <% else %>
					<tr align="center" height="25">
	            <% end if %>

                <td width="30"><font color="<%= oorderdetail.FItemList(ix).CancelStateColor %>"><%= oorderdetail.FItemList(ix).CancelYnName %></font></td>
            	<td width="80">
            	    <% if oorderdetail.FItemList(ix).Fisupchebeasong="Y" then %>
            	    	<a href="javascript:popOrderDetailEdit('<%= oorderdetail.FItemList(ix).fdetailidx %>');"><font color="red">
            	    	<%=oorderdetail.FItemList(ix).fitemgubun%>-<%=CHKIIF(oorderdetail.FItemList(ix).fitemid>=1000000,Format00(8,oorderdetail.FItemList(ix).fitemid),Format00(6,oorderdetail.FItemList(ix).fitemid))%>-<%=oorderdetail.FItemList(ix).fitemoption%>
            	    	<br>(업체)</font></a>
                    <% else %>
                    	<a href="javascript:popOrderDetailEdit('<%= oorderdetail.FItemList(ix).fdetailidx %>');">
                    	<%=oorderdetail.FItemList(ix).fitemgubun%>-<%=CHKIIF(oorderdetail.FItemList(ix).fitemid>=1000000,Format00(8,oorderdetail.FItemList(ix).fitemid),Format00(6,oorderdetail.FItemList(ix).fitemid))%>-<%=oorderdetail.FItemList(ix).fitemoption%>
                    	</a>
                    <% end if %>
                </td>
                <td width="120" align="left">
                    <a href="javascript:popSimpleBrandInfo('<%= oorderdetail.FItemList(ix).Fmakerid %>');">
                    <acronym title="<%= oorderdetail.FItemList(ix).Fmakerid %>"><%= Left(oorderdetail.FItemList(ix).Fmakerid,12) %></acronym>
                    </a>
                </td>
            	<td align="left">
            	    <acronym title="<%= oorderdetail.FItemList(ix).FItemName %>"><%= Left(oorderdetail.FItemList(ix).FItemName,64) %></acronym>
            	    <% if oorderdetail.FItemList(ix).FItemoption<>"0000" then %>
                	    <br>
                	    <font color="blue"><%= oorderdetail.FItemList(ix).FItemoptionName %></font>
            	    <% end if %>
            	</td>

            	<% if oorderdetail.FItemList(ix).FItemNo > 1 then %>
            		<td  width="30"><b><font color="blue"><%= oorderdetail.FItemList(ix).FItemNo %></font></b></td>
            	<% elseif oorderdetail.FItemList(ix).FItemNo < 1 then %>
            		<td  width="30"><b><font color="red"><%= oorderdetail.FItemList(ix).FItemNo %></font></b></td>
            	<% else %>
            		<td  width="30"><font color="blue"><%= oorderdetail.FItemList(ix).FItemNo %></font></td>
            	<% end if %>

                <td align="right" width="50"><%= FormatNumber(oorderdetail.FItemList(ix).fsellprice,0) %></td> <!-- 소비자가 -->
				<td align="center" width="50"><font color="red"><%= FormatNumber(oorderdetail.FItemList(ix).frealsellprice,0) %></font></td>
            </tr>
            <tr>
        		<td height="1" colspan="15" bgcolor="#BABABA"></td>
        	</tr>
            <% end if %>

            <% next %>

            </form>
        </table>
    </td>
</tr>
</table>

<%
set oordermaster = Nothing
set oorderdetail = Nothing
%>
<!-- #include virtual="/admin/offshop/shopcscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->