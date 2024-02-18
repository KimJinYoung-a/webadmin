<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프샵 리스트
' History : 2009.04.07 서동석 생성
'			2010.08.04 한용민 수정
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
dim ochargeuser ,i, shopdiv, isusing ,shopname, shopid, reloading, currencyUnit, loginsite, countrylangcd, vieworder
	menupos = requestCheckvar(request("menupos"),10)
	shopdiv = requestCheckvar(request("shopdiv"),32)
	isusing = requestCheckvar(request("isusing"),10)
	shopname = requestCheckvar(request("shopname"),64)
	shopid = requestCheckvar(request("shopid"),32)
	reloading = requestCheckvar(request("reloading"),2)
	currencyUnit = requestCheckvar(request("currencyUnit"),3)
	loginsite = requestCheckvar(request("loginsite"),32)
	countrylangcd = requestCheckvar(request("countrylangcd"),32)
	vieworder = requestCheckvar(request("vieworder"),1)

if reloading="" and isusing="" then isusing="Y"

set ochargeuser = new COffShopChargeUser
    ochargeuser.FRectShopDiv2 = shopdiv
    ochargeuser.FRectIsUsing = isusing
    ochargeuser.frectshopname = shopname
    ochargeuser.FRectShopID = shopid
    ochargeuser.FRectcurrencyUnit = currencyUnit
    ochargeuser.FRectloginsite = loginsite
    ochargeuser.FRectcountrylangcd = countrylangcd
	ochargeuser.FRectvieworder = vieworder
	ochargeuser.GetOffShopList
%>

<script language='javascript'>

function popShopInfo(ishopid){
	var popwin = window.open("/admin/lib/popoffshopinfo.asp?shopid=" + ishopid + "&menupos=<%=menupos%>","popoffshopinfo",'width=1024,height=768,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popShoplinkothersite(){
	var popShoplinkothersite = window.open("/admin/offshop/othersite/Shoplinkothersite.asp?menupos=<%=menupos%>","popShoplinkothersite",'width=1024,height=768,scrollbars=yes,resizable=yes');
	popShoplinkothersite.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="reloading" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* ShopID : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		&nbsp;&nbsp;
		* Shop명 : <input type="text" name="shopname" size=30 ,maxlength=30 value="<%=shopname%>">
		&nbsp;&nbsp;
		* Shop구분 : <% drawoffshop_commoncode "shopdiv", shopdiv, "shopdiv", "SUB", "", "" %>
		&nbsp;&nbsp;
		* Shop운영여부 : <% Call drawSelectBoxUsingYN("isusing",isusing) %>
		<br><br>
		* 대표화폐 : <% DrawexchangeRate "currencyUnit",currencyUnit,"" %>
		&nbsp;&nbsp;
		* 로그인사이트 : <% drawoffshop_commoncode "loginsite", loginsite, "loginsite", "MAIN", "", "" %>
		&nbsp;&nbsp;
		* 대표언어 : <% DrawexchangeRate_countrylangcd "countrylangcd",countrylangcd, "", "" %>
		&nbsp;&nbsp;
		* 프론트 노출 여부 : <select class="select" name="vieworder"><option value=""<% if vieworder="" then response.write " selected" %>>전체</option><option value="1"<% if vieworder="1" then response.write " selected"%>>노출</option><option value="0"<% if vieworder="0" then response.write " selected"%>>노출안함</option></select>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">

	</td>
	<td align="right">
		<% if C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH or C_logicsPowerUser then %>
			<input type="button" class="button" value="신규등록" onclick="popShopInfo('')">
		<% end if %>

		<!--<input type="button" class="button" value="외부매장매칭" onclick="popShoplinkothersite()">-->
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if ochargeuser.FresultCount > 0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ochargeuser.fresultcount %></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>ShopID</td>
	<td>Shop명</td>
	<td>Shop구분</td>
	<td>그룹코드<br>사업자번호</td>
	<td>회사명</td>
	<td>Shop<br>전화전호</td>
	<td>매니저</td>
	<td>매니저HP<br>매니저E-mail</td>
	<td>국가</td>
	<td>로그인<br>사이트</td>
	<td>대표<br>화폐</td>
	<td>대표<br>언어</td>
	<td>운영<br>여부</td>
	<td>화면<br>표시</td>
	<td>비고</td>
</tr>
<%
for i=0 to ochargeuser.FresultCount - 1
%>
<% if ochargeuser.FItemList(i).FIsUsing="N" then %>
	<tr bgcolor="#e1e1e1" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#e1e1e1';>
<% else %>
	<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
<% end if %>
	<td><%= ochargeuser.FItemList(i).Fuserid %></td>
	<td><%= ochargeuser.FItemList(i).Fshopname %></td>
	<td><%= ochargeuser.FItemList(i).FShopdivName %> (<%= ochargeuser.FItemList(i).fshopdiv %>)</td>
	<td><%= ochargeuser.FItemList(i).Fgroupid %><br><%= printUserId(ochargeuser.FItemList(i).fcompany_no, 2, "*") %></td>
	<td><%= ochargeuser.FItemList(i).fcompany_name %></td>
	<td><%= printtel(ochargeuser.FItemList(i).Fshopphone) %></td>
	<td><%= printUserId(ochargeuser.FItemList(i).Fmanname, 1, "*") %></td>
	<td>
		<%= printtel(ochargeuser.FItemList(i).Fmanhp) %>
		<br><%= printUserId(ochargeuser.FItemList(i).Fmanemail, 2, "*") %>
	</td>
	<td><%= ochargeuser.FItemList(i).FcountryNamekr %></td>
	<td><%= ochargeuser.FItemList(i).floginsite %></td>
	<td><%= ochargeuser.FItemList(i).fcurrencyUnit %></td>
	<td><%= ochargeuser.FItemList(i).fcountrylangcd %></td>
	<td><%= ochargeuser.FItemList(i).FIsUsing %></td>
	<td><%= ochargeuser.FItemList(i).Fvieworder %></td>
	<td>
		<input type="button" class="button" value="수정" onclick="popShopInfo('<%= ochargeuser.FItemList(i).Fuserid %>')">
	</td>
</tr>
<%
next
else
%>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan=10>검색 결과가 없습니다</td>
</tr>
<%
end if
%>
</table>

<%
set ochargeuser = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
