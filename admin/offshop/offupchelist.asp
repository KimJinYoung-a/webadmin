<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프샵
' History : 2009.04.07 서동석 생성
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->
<%
dim shopid, designer, comm_cd, shopusing, partnerusing, page, research, diffCk, offupbea, i
dim hasContOnly, maeipdiv, vPurchaseType, isoffusing, adminopen, diffshopdiv13
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
	diffshopdiv13      = RequestCheckVar(request("diffshopdiv13"),2)

if page="" then page=1
if (research="") and (hasContOnly="") then hasContOnly="ON"
if (research="") then shopusing="Y"
if diffshopdiv13="on" then
	hasContOnly="OFF"
	comm_cd=""
end if

dim ochargeuser

set ochargeuser = new COffShopChargeUser
	ochargeuser.FCurrPage = page

	if (designer<>"") then
		ochargeuser.FPageSize = 400
	else
		ochargeuser.FPageSize = 100
	end if

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
	elseif diffshopdiv13<>"" then
		ochargeuser.GetOffShopbrandcontractdiff
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
%>
<script type="text/javascript">

function editOffDesinger(shopid,designerid){
	var popwin = window.open("/admin/lib/popshopupcheinfo.asp?shopid=" + shopid + "&designer=" + designerid,"popshopupcheinfo","width=1280 height=768 scrollbars=yes resizable=yes");
	popwin.focus();
}

function popXL() {
	frm.action="/admin/offshop/offupchelist_xl_download.asp";
	frm.target='view';
	frm.submit();
	frm.action='';
	frm.target='';
}

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

</script>

<iframe id="view" name="view" src="" width="0" height="0" allowtransparency="true" frameborder="0" scrolling="no"></iframe>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" >
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		ShopID : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		&nbsp;
		브랜드ID : <% drawSelectBoxDesignerwithName "designer",designer  %>
     	&nbsp;
     	Shop운영여부 : <% drawSelectBoxUsingYN "shopusing",shopusing %>
		&nbsp;
		계약여부 :
     	<select name='hasContOnly'>
     		<option value=''>전체</option>
     		<option value='ON' <% if hasContOnly="ON" then response.write "selected" %>>계약Y</option>
     		<option value='OFF' <% if hasContOnly="OFF" then response.write "selected" %>>계약N</option>
     	</select>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		구매유형 : 
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
		&nbsp;
		OFF계약구분 : <% 'drawSelectBoxOFFChargeDiv "chargediv",chargediv %>
		<% drawSelectBoxOFFJungsanCommCD "comm_cd",comm_cd %>
		&nbsp;
		ON계약구분 :
		<% DrawBrandMWUCombo "maeipdiv",maeipdiv %>
		&nbsp;
		OFF브랜드사용여부 :
		<% drawSelectBoxUsingYN "isoffusing",isoffusing %>
		&nbsp;
		오프라인어드민사용여부 :
		<% drawSelectBoxUsingYN "adminopen",adminopen %>
     	&nbsp;
     	SCM오픈여부 : <% drawSelectBoxUsingYN "partnerusing",partnerusing %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="checkbox" name="diffCk" <%= ChkIIF(diffCk="on","checked","") %> >대표마진 불일치 검색
     	&nbsp;
     	<input type="checkbox" name="offupbea" <%= ChkIIF(offupbea="on","checked","") %> >오프 업체배송
     	&nbsp;
     	<input type="checkbox" name="diffshopdiv13" <%= ChkIIF(diffshopdiv13="on","checked","") %> >계약불일치(직영점대표Y,가맹점대표N)
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
		<% if (shopid="") and (designer="") and (offupbea="") then %>
		<div align="center"><font color="red">매장 또는 브랜드를 선택하세요.</font></div>
		<% end if %>
    </td>
    <td align="right">
		<input type="button" class="button" value="엑셀다운" onclick="popXL();">
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		검색결과 : <b><%= ochargeuser.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= ochargeuser.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
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
	<td rowspan=2 width="50">수정</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
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

	<td>
		<%if (ochargeuser.FItemList(i).FShopid="streetshop000") or (ochargeuser.FItemList(i).FShopid="streetshop800") or (ochargeuser.FItemList(i).FShopid="streetshop870") then %>
			<strong><%= ochargeuser.FItemList(i).FShopID %></strong>
		<% else %>
			<%= ochargeuser.FItemList(i).FShopID %>
		<% end if %>
	</td>
	<td>
		<%= ochargeuser.FItemList(i).FShopName %>
	</td>
	<td><font color="<%= CHKIIF(ochargeuser.FItemList(i).IsContractExists,"#000000","#AAAAAA") %>"><%= ochargeuser.FItemList(i).FDesignerId %></font></td>
	<td><font color="<%= CHKIIF(ochargeuser.FItemList(i).IsContractExists,"#000000","#AAAAAA") %>"><%= ochargeuser.FItemList(i).FDesignerName %></font></td>
	<td>
		<%= getBrandPurchaseType(ochargeuser.FItemList(i).fpurchaseType) %>
	</td>

	<% if (ochargeuser.FItemList(i).IsContractExists) then %>
		<td>
			<font color="<%= ochargeuser.FItemList(i).getChargeDivColor %>">
				<%= ochargeuser.FItemList(i).getChargeDivName %>
				<% if (ochargeuser.FItemList(i).Fjungsan_gubun = "간이과세") then %>
				<br />(간이)
				<% end if %>
			</font>
		</td>
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
	<td align="center"><input type="button" class="button" value="수정" onclick="editOffDesinger('<%= ochargeuser.FItemList(i).FShopid %>','<%= ochargeuser.FItemList(i).FDesignerId %>');"></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="25" align="center">
	<% if ochargeuser.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ochargeuser.StarScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + ochargeuser.StarScrollPage to ochargeuser.FScrollCount + ochargeuser.StarScrollPage - 1 %>
		<% if i>ochargeuser.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if ochargeuser.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% ELSE %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="20">등록된 내용이 없습니다.</td>
</tr>
<% end if %>
</table>

<%
set ochargeuser = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
