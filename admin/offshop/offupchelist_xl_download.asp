<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프샵
' History : 2016.06.28 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->
<%
dim shopid, designer, comm_cd, shopusing, partnerusing, page, research, diffCk, offupbea, i, menupos
dim hasContOnly, maeipdiv, vPurchaseType, isoffusing, adminopen
	page        = RequestCheckVar(request("page"),9)
	shopid      = RequestCheckVar(request("shopid"),32)
	designer    = RequestCheckVar(request("designer"),32)
	comm_cd     = RequestCheckVar(request("comm_cd"),9)
	shopusing   = RequestCheckVar(request("shopusing"),1)
	partnerusing  = RequestCheckVar(request("partnerusing"),1)
	research    = RequestCheckVar(request("research"),9)
	diffCk      = RequestCheckVar(request("diffCk"),9)
	offupbea    = RequestCheckVar(request("offupbea"),9)
	hasContOnly    = RequestCheckVar(request("hasContOnly"),9)
	maeipdiv = RequestCheckVar(request("maeipdiv"),1)
	vPurchaseType = requestCheckVar(request("purchasetype"),2)
	isoffusing = requestCheckVar(request("isoffusing"),1)
	adminopen = requestCheckVar(request("adminopen"),1)
	menupos = getNumeric(requestcheckvar(request("menupos"),10))

if page="" then page=1
if (research="") and (hasContOnly="") then hasContOnly="ON"
if (research="") then shopusing="Y"

dim ochargeuser

set ochargeuser = new COffShopChargeUser
	ochargeuser.FCurrPage = 1
	ochargeuser.FPageSize = 5000
	ochargeuser.FRectShopID     = shopid
	ochargeuser.FRectDesigner   = designer
	ochargeuser.FRectComm_cd    = comm_cd
	ochargeuser.FRectShopusing  = shopusing
	ochargeuser.FRectpartnerusing  	= partnerusing
	ochargeuser.FRectOffUpBea   	= offupbea
	ochargeuser.FRectHasContOnly  = hasContOnly
	ochargeuser.FRectmaeipdiv = maeipdiv
	ochargeuser.FRectBrandPurchaseType = vPurchaseType
	ochargeuser.FRectisoffusing = isoffusing
	ochargeuser.FRectadminopen = adminopen

	if (diffCk<>"") then
		ochargeuser.GetOffShopbrandcontractlisterror
	else
	    if (shopid="") and (designer="") then
	        if (offupbea<>"") then
	            ochargeuser.GetOffShopbrandcontractlist
	        end if
	    else
	        if (shopid<>"") then
	    		ochargeuser.GetOffShopDesignerList1
	    	else
	    		ochargeuser.GetOffShopbrandcontractlist
	    	end if
	    end if
	end if

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_입점업체마진관리_" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>

<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>

<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
		<% if (shopid="") and (designer="") and (offupbea="") then %>
		<div align="center"><font color="red">매장 또는 브랜드를 선택하세요.</font></div>
		<% end if %>
    </td>
    <td align="right">
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#DDDDDD" border=1>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="14">
		검색결과 : <b><%= ochargeuser.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td rowspan=2 width="100">ShopID</td>
	<td rowspan=2 width="100">Shop명</td>
	<td rowspan=2>브랜드ID</td>
	<td rowspan=2>브랜드명</td>
	<td rowspan=2 width="70">구매유형</td>
	<td colspan=3>OFF 계약</td>
	<td colspan=2>ON 계약</td>
	<td rowspan=2 width="50">OFF<br>브랜드<br>사용여부</td>
	<td rowspan=2 width="50">오프라인<br>어드민<br>오픈여부</td>
	<td rowspan=2 width="50">SCM<br>오픈여부</td>
	<td rowspan=2 width="50">정산일</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="90">계약구분</td>
	<td width="50">업체<br>매입마진</td>
	<td width="50">SHOP<br>출고마진</td>
	<td width="50">계약구분</td>
	<td width="50">마진</td>
</tr>
<% if ochargeuser.FresultCount >0 then %>
<% for i=0 to ochargeuser.FresultCount-1 %>

<% if ochargeuser.FItemList(i).FShopIsUsing="Y" then %>
	<tr align="center" bgcolor="#FFFFFF">
<% else %>
	<tr align="center" bgcolor="#DDDDDD">
<% end if %>

	<td class='txt'>
		<%if (ochargeuser.FItemList(i).FShopid="streetshop000") or (ochargeuser.FItemList(i).FShopid="streetshop800") or (ochargeuser.FItemList(i).FShopid="streetshop870") then %>
			<strong><%= ochargeuser.FItemList(i).FShopID %></strong>
		<% else %>
			<%= ochargeuser.FItemList(i).FShopID %>
		<% end if %>
	</td>
	<td>
		<%= ochargeuser.FItemList(i).FShopName %>
	</td>
	<td class='txt'><font color="<%= CHKIIF(ochargeuser.FItemList(i).IsContractExists,"#000000","#AAAAAA") %>"><%= ochargeuser.FItemList(i).FDesignerId %></font></td>
	<td><font color="<%= CHKIIF(ochargeuser.FItemList(i).IsContractExists,"#000000","#AAAAAA") %>"><%= ochargeuser.FItemList(i).FDesignerName %></font></td>
	<td>
		<%= getBrandPurchaseType(ochargeuser.FItemList(i).fpurchaseType) %>
	</td>

	<% if (ochargeuser.FItemList(i).IsContractExists) then %>
		<td><font color="<%= ochargeuser.FItemList(i).getChargeDivColor %>"><%= ochargeuser.FItemList(i).getChargeDivName %></font></td>
		<td><%= ochargeuser.FItemList(i).FDefaultMargin %></td>
		<td><%= ochargeuser.FItemList(i).FDefaultSuplyMargin %></td>
	<% else %>
		<td></td>
		<td></td>
		<td></td>
	<% end if %>
	<td>
		<font color="<%= CHKIIF(ochargeuser.FItemList(i).IsContractExists,"#000000","#AAAAAA") %>">
		<%= ochargeuser.FItemList(i).getMwName %></font>
	</td>
	<td>
		<%= ochargeuser.FItemList(i).Fonlinedefaultmargine %>
	</td>
	<td align="center">
		<% if (ochargeuser.FItemList(i).fisoffusing="Y") then  %>
			O
		<% else %>
			X
		<% end if %>
	</td>
	<td align="center">
		<% if (ochargeuser.FItemList(i).FAdminopen="Y") then  %>
			O
		<% else %>
			X
		<% end if %>
	</td>
	<td align="center">
		<% if (ochargeuser.FItemList(i).FPartnerisusing="Y") then  %>
			O
		<% else %>
			X
		<% end if %>
	</td>
	<td><%= ochargeuser.FItemList(i).Fjungsan_date_off %></td>	
</tr>
<% next %>

<% ELSE %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="14">등록된 내용이 없습니다.</td>
</tr>
<% end if %>
</table>

<%
set ochargeuser = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
