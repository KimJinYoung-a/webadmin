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
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/classes/offshop/order/shopcscenter_order_cls.asp"-->
<!-- #include virtual="/admin/offshop/shopcscenter/cscenter_Function_off.asp"-->
<%
dim masteridx ,ojumun , oaslist, totalascount ,ix , orderno ,shopid , oaslistmaejang , oaslistfinal
dim maejangascount , finalascount
	masteridx = RequestCheckVar(request("masteridx"),16)

totalascount = 0

set ojumun = new COrder
	if masteridx <> "" then
	    ojumun.FRectmasteridx = masteridx
	    ojumun.fQuickSearchOrderMaster
	end if

if ojumun.ftotalcount > 0 then
	orderno = ojumun.FOneItem.Forderno
	shopid = ojumun.FOneItem.Fshopid
end if

'/관련a/s 건수
set oaslist = new COrder
	if masteridx <> "" then
	    oaslist.FRectmasteridx = masteridx
	    oaslist.fGetCSASTotalCount
		
	    totalascount = oaslist.FResultCount
	end if

'/매장처리 대상건수
set oaslistmaejang = new COrder
	if masteridx <> "" then
	    oaslistmaejang.FRectmasteridx = masteridx
	    oaslistmaejang.frectcurrstate = "'B001','B004'"
	    oaslistmaejang.frectdeleteyn = "N"
	    oaslistmaejang.fGetCSASTotalCount
		
	    maejangascount = oaslistmaejang.FResultCount
	end if

'/최종완료처리 대상건수
set oaslistfinal = new COrder
	if masteridx <> "" then
	    oaslistfinal.FRectmasteridx = masteridx
	    oaslistfinal.frectcurrstate = "'B006','B008'"
	    oaslistfinal.frectdeleteyn = "N"
	    oaslistfinal.fGetCSASTotalCount
		
	    finalascount = oaslistfinal.FResultCount
	end if	 
	 

%>

<script language="javascript">
</script>

<% if (masteridx<>"") then %>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="FFFFFF">
<tr height="25">
	<td align="left">
		<input type="button" class="button" value="업체A/S신청" class="csbutton" onclick="javascript:PopOpenServiceItemas('<%= masteridx %>');">
		<input type="button" class="button" value="매장완료처리[<%=maejangascount%>]" class="csbutton" onclick="javascript:PopmaejangAction('<%=orderno%>','<%= shopid %>','','notfinish');">
		<input type="button" class="button" value="최종완료처리/수정[<%= finalascount %>]" class="csbutton" onclick="javascript:Cscenter_Action_List_off('<%= masteridx %>','<%=orderno%>','','notfinish','<%= shopid %>');">
    </td>
    <td align="right">
    	<input type="button" class="button" value="영수증재출력" style="width:90px;" onclick="javascript:popOrderReceipt('<%= orderno %>');">
	</td>
</tr>
</table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr valign="top">
	<td>
		<!-- 구매상품정보 -->
		<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="25" bgcolor="<%= adminColor("topbar") %>" style="padding:2 2 2 2">
		    <td colspan="10">
		    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		    		<tr>
		    			<td width="500">
		    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>구매상품정보</b>
					    	[<b><%= orderno %> 주문 관련 총A/S <%=totalascount%>건</b>]
    				    </td>
    				    <td align="right">
    				    </td>
    				</tr>
    			</table>
    		</td>
		</tr>
		<tr height="400" bgcolor="#FFFFFF">
		    <td valign="top">
		        <table height="25" width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#BABABA">
		            <tr align="center" bgcolor="<%= adminColor("topbar") %>" style="padding:2">
                    	<td width="30">구분</td>
                    	<td width="80">CODE</td>                      	
                        <td width="120">브랜드ID</td>
                    	<td>상품명<font color="blue">[옵션명]</font></td>
                    	<td width="30">수량</td>
                    	<td width="50">현재<br>소비자가</td>
                    	<td width="50">판매가</td>
                    </tr>
                    <tr>
                        <td height="1" colspan="15" bgcolor="#BABABA"></td>
                    </tr>
                 </table>
                 <table height="365" width="100%" border=0 cellspacing=0 cellpadding=0 class=a bgcolor="FFFFFF">
                    <tr height="100%">
                        <td colspan="12">
                	        <iframe name="orderdetail" src="/admin/offshop/shopcscenter/order/orderitemmaster.asp?masteridx=<%= masteridx %>" border=0 frameSpacing=0 frameborder="no" width="100%" height="100%" leftmargin="0"></iframe>
                        </td>
                    <tr>
                </table>
		    </td>
		</tr>
		</table>
		<!-- 구매상품정보 -->
	</td>
	<td width="5"></td>
	<td width="250" align="right">
		<!-- 구매자정보 -->
		<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<form name="frmbuyerinfo" onsubmit="return false;">
		<tr height="25" bgcolor="<%= adminColor("topbar") %>">
		    <td colspan="2">
		    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		    		<tr>
		    			<td width="100">
		    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>구매자 정보</b>
    				    </td>
    				    <td align="right">    				    	
    				    </td>
    				</tr>
    			</table>
    		</td>
		</tr>
		<tr height="23">
		    <td bgcolor="<%= adminColor("topbar") %>">IDX</td>
		    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.fmasteridx %></td>
		</tr>
		<tr height="23">
		    <td bgcolor="<%= adminColor("topbar") %>">주문번호</td>
		    <td bgcolor="#FFFFFF"><%= orderno %></td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">구매자명</td>
		    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.FBuyName %></td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">전화번호</td>
		    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.FBuyPhone %></td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">핸드폰</td>
		    <td bgcolor="#FFFFFF">
		        <%= ojumun.FOneItem.FBuyHp %>
		        <input type="button" name="buyhp" class="button" value="SMS" onclick="javascript:PopCSSMSSend_off('<%= ojumun.FOneItem.FBuyHp %>','<%= ojumun.FOneItem.Fmasteridx %>','<%= ojumun.FOneItem.forderno %>','');">
		    </td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">이메일</td>
		    <td bgcolor="#FFFFFF">
		        <%= ojumun.FOneItem.FBuyEmail %>
		    </td>
		</tr>
		</form>
		</table>
		<!-- 구매자정보 -->
		<Br>
	    <!-- 주문정보 -->
		<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">주문상태</td>
		    <td bgcolor="#FFFFFF">				        
		        <font color="<%= ojumun.FOneItem.CancelYnColor %>"><%= ojumun.FOneItem.CancelYnName %></font>
		    </td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">주문일시</td>
		    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.FRegDate %></td>
		</tr>
		<!-- 주문정보 -->		
		</table>	
	</td>
</tr>
</table>

<br>

<% else %>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr height="50">
    <td align="center"> [ 상세내역을 보시려면 주문번호를 선택 하세요 ]</td>
</tr>
</table>
<% end if %>

<%
set ojumun = Nothing
set oaslist = Nothing
set oaslistmaejang = Nothing
set oaslistfinal = Nothing
%>
<!-- #include virtual="/admin/offshop/shopcscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->