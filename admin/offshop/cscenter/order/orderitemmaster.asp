<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 고객센터
' Hieditor : 2011.03.08 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/order_cls.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->
<%
dim masteridx, onimage, research ,itemcostcomment, itemcostcolor ,oordermaster, oorderdetail
dim  i, ix
masteridx = requestCheckVar(request("masteridx"),10)
onimage     = requestCheckVar(request("onimage"),10)
research    = requestCheckVar(request("research"),2)

if (onimage = "") and (research="") then  onimage = "on"

set oordermaster = new COrder
	oordermaster.FRectmasteridx = masteridx
	oordermaster.fQuickSearchOrderMaster

set oorderdetail = new COrder
	oorderdetail.FRectmasteridx = masteridx
	oorderdetail.fQuickSearchOrderDetail
%>

<script type='text/javascript'>

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
	            <% if oorderdetail.FItemList(ix).Fitemid =0 then %>

	            <% if oorderdetail.FItemList(ix).FCancelyn ="Y" then %>
	            <tr align="center" height="25" bgcolor="#EEEEEE" class="gray">
	            <% else %>
	            <tr align="center" height="25">
	            <% end if %>
	                <td width="30"></td>
	                <td width="50"></td>
	            	<td width="80">0</td>
	            	<td width="120" align="left"><%= oorderdetail.FItemList(ix).FMakerid %></td>
	            	<td align="left">배송비<font color="blue">[<%= oorderdetail.BeasongCD2Name(oorderdetail.FItemList(ix).Fitemoption) %>]</font></td>
	            	<td width="30"></td>
	            	<td width="50"></td>
	            	<td width="50" align="right"><font color="blue"><%= FormatNumber(oorderdetail.FItemList(ix).Fitemcost,0) %></font></td>
	            	<td width="70"></td>
	            	<td width="70"></td>
	            	<td width="108"></td>
	            </tr>
	            <% end if %>
            <% next %>
            <tr>
        		<td height="1" colspan="12" bgcolor="#BABABA"></td>
        	</tr>
            <% for ix=0 to oorderdetail.FResultCount-1 %>
            <% if oorderdetail.FItemList(ix).Fitemid <>0 then %>

            <% if oorderdetail.FItemList(ix).FCancelyn ="Y" then %>
            <tr align="center" height="25" bgcolor="#EEEEEE" class="gray">
            <% else %>
            <tr align="center" height="25">
            <% end if %>

                <td width="30"><font color="<%= oorderdetail.FItemList(ix).CancelStateColor %>"><%= oorderdetail.FItemList(ix).CancelYnName %></font></td>
                <td width="50"><font color="<%= oorderdetail.FItemList(ix).GetStateColor %>"><%= oorderdetail.FItemList(ix).GetStateName %></font></td>
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
            	<td width="70"><acronym title="<%= oorderdetail.FItemList(ix).Fupcheconfirmdate %>"><%= Left(oorderdetail.FItemList(ix).Fupcheconfirmdate,10) %></acronym></td>
            	<td width="70"><acronym title="<%= oorderdetail.FItemList(ix).Fbeasongdate %>"><%= Left(oorderdetail.FItemList(ix).Fbeasongdate,10) %></acronym></td>
            	<td width="108">
            	    <%= oorderdetail.FItemList(ix).Fsongjangdivname %><br>
            		<% if (oorderdetail.FItemList(ix).FsongjangDiv="24") then %>
            			<a href="javascript:popDeliveryTrace('<%= oorderdetail.FItemList(ix).Ffindurl %>','<%= oorderdetail.FItemList(ix).Fsongjangno %>');"><%= oorderdetail.FItemList(ix).Fsongjangno %></a>
            	    <% else %>
            	    	<a target="_blank" href="<%= oorderdetail.FItemList(ix).Ffindurl + oorderdetail.FItemList(ix).Fsongjangno %>"><%= oorderdetail.FItemList(ix).Fsongjangno %></a>
            	    <% end if %>
            	</td>
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
<form name="popForm" action="popDeliveryTrace.asp" target="_blank">
	<input type="hidden" name="traceUrl">
	<input type="hidden" name="songjangNo">
</form>
</table>

<script language="javascript">

function popDeliveryTrace(traceUrl, songjangNo)
{
	var f = document.popForm;
	f.traceUrl.value	= traceUrl;
	f.songjangNo.value	= songjangNo;
	f.submit();
}

</script>

<%
set oordermaster = Nothing
set oorderdetail = Nothing
%>
<!-- #include virtual="/admin/offshop/cscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->