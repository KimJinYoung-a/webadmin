<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 온라인 & 오프라인 어드민 장바구니
' Hieditor : 2011.08.02 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/stock/shortagestock_cls.asp" -->
<!-- #include virtual="/lib/classes/items/adminshoppingbag/adminshoppingbag_cls.asp" -->
<%
dim IS_HIDE_BUYCASH : IS_HIDE_BUYCASH = False

'// 매입가 소스에서 제거, skyer9, 2018-02-14
dim PriceEditEnable : PriceEditEnable = False

dim itemgubunarr , itemidarr , itemoptionarr, itemnoarr, onoffgubun, shopid , i ,acURL ,research
dim obaginsert , userid , isusing ,makerid ,itemid ,itemname ,comm_cd ,cdl ,cdm ,cds ,obag, myorderyn
	itemgubunarr = request("itemgubunarr")
	itemidarr = request("itemidarr")
	itemoptionarr = request("itemoptionarr")
	itemnoarr = request("itemnoarr")
	onoffgubun = requestCheckVar(request("onoffgubun"),10)
	userid = session("ssBctId")
    isusing = requestCheckVar(request("isusing"),1)
    makerid = requestCheckVar(request("makerid"),32)
    itemid = requestCheckVar(request("itemid"),10)
    itemname = requestCheckVar(request("itemname"),64)
    comm_cd = requestCheckVar(request("comm_cd"),32)
    cdl = requestCheckVar(request("cdl"),3)
    cdm = requestCheckVar(request("cdm"),3)
    cds = requestCheckVar(request("cds"),3)
	shopid = requestCheckVar(request("shopid"),32)
    research = requestCheckVar(request("research"),2)
    myorderyn = requestcheckvar(request("myorderyn"),1)

if (research<>"on") and (isusing="") then
    isusing = "Y"
end if
if (research<>"on") and (myorderyn="") then myorderyn="Y"

if C_ADMIN_USER then

'/매장일경우 본인 매장만 사용가능
elseif (C_IS_SHOP) then
	IS_HIDE_BUYCASH = True
	myorderyn = "Y"

	'/어드민권한 점장 미만
	'if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	'end if
else
	myorderyn = "Y"

	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then

		else

		end if
	end if
end if

'/온오프 구분 페이지 초기값 오프라인(OFF)
if onoffgubun = "" then onoffgubun = "OFF"
if onoffgubun = "" then
	response.write "<script>alert('온라인 & 오프구분이 없습니다'); self.close();</script>"
	dbget.close() : response.end
end if

'//장바구니 추가
'response.write userid &"/"& onoffgubun&"/"& shopid&"/"& itemgubunarr&"/"& itemidarr&"/"& itemoptionarr&"/"& itemnoarr &"<Br>"
putadminshoppingbag_insert userid, onoffgubun, shopid, itemgubunarr, itemidarr, itemoptionarr, itemnoarr , "self" ,menupos

set obag  = new cadminshoppingbag_list
	obag.FPageSize = 300
	obag.FCurrPage = 1
    obag.frectcdl = cdl
    obag.frectcdm = cdm
    obag.frectcds = cds
    obag.Frectshopid = shopid
    obag.Frectisusing = isusing

    'if onoffgubun = "" and itemgubunarr = "" and itemidarr = "" and itemoptionarr = "" and itemnoarr = "" then
    	obag.Frectmakerid = makerid
    'end if

    obag.Frectitemid = itemid
    obag.Frectitemname = itemname
    obag.Frectcomm_cd = comm_cd
    obag.frectonoffgubun = onoffgubun

	if myorderyn="Y" then
		obag.frectuserid = userid
	end if

	'/온라인 장바구니 리스트
    if onoffgubun = "ON" then
    	obag.fadminshoppingbag_on

    '/오프라인 장바구니 리스트
    elseif onoffgubun = "OFF" then
        obag.fadminshoppingbag_off

	    'if shopid = "" then
	    '    response.write "<script language='javascript'>"
	    '    response.write "    alert('매장을 선택하셔야 주문이 가능 합니다');"
	    '    response.write "</script>"
	    'end if
    end if

'//신규상품 추가시 팝업으로 넘어갈 경로		'/공용팝업으로 액션 페이지를 통채로 넘긴다
acURL =Server.HTMLEncode("/common/item/adminshoppingbag_process.asp?onoffgubun="&onoffgubun)
%>

<font color="red">※ <%= userid %> 님의 <%= onoffgubun %>LINE 장바구니</font>

<%
'/온라인 장바구니
if onoffgubun = "ON" then
%>

<%
'/오프라인 장바구니
elseif onoffgubun = "OFF" then
%>
    <Br>&nbsp;&nbsp;&nbsp;- 물류센터 주문 : 정산구분(텐바이텐특정/출고매입/출고특정)
    <br>&nbsp;&nbsp;&nbsp;- 업체 주문 : 정산구분(업체특정/업체매입)
    <br>&nbsp;&nbsp;&nbsp;- 필요수량(7일) = (7일판매분 x 1) - (유효재고 + 기주문건)
	<!-- 검색 시작 -->
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="onoffgubun" value="<%= onoffgubun %>">
	<tr align="center" bgcolor="#FFFFFF" >
	    <td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	    <td align="left">
	        매장 :
	        <% if C_ADMIN_USER then %>
				<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
	        <% elseif (C_IS_SHOP) then %>
	    		<%= shopid %>
	    	<% else %>
				<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
	        <% end if %>

	        사용여부:<% drawSelectBoxUsingYN "isusing", isusing %>
	        <!-- #include virtual="/common/module/categoryselectbox.asp"-->
	    </td>
	    <td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
	        <input type="button" class="button_s" value="검색" onClick="javascript:reg(frm);">
	    </td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
	    <td align="left">
	        브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
	        &nbsp;
	        상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="9" onKeyPress="if (event.keyCode == 13) reg(frm);">
	        &nbsp;
	        상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32" onKeyPress="if (event.keyCode == 13) reg(frm);">
	        정산기준 : <% drawSelectBoxOFFJungsanCommCDmulti "comm_cd",comm_cd %>
	    </td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
	    <td align="left">
	    	<% 'if C_ADMIN_USER then %>
				나의장바구니만보기<input type="checkbox" name="myorderyn" value="Y" <% if myorderyn="Y" then response.write " checked" %>>
			<% 'end if %>
	    </td>
	</tr>
	</form>
	</table>
	<!-- 검색 끝 -->
	<br>

	<!-- 액션 시작 -->
	<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
	    <td align="left">
	        <input type="button" class="button" value="선택수정" onclick="bageditarr(frmbag)">
	        <input type="button" class="button" value="선택삭제" onclick="bagdelarr(frmbag)">
	    </td>
	    <td align="right">
			<% if True or (session("ssBctCname") = "이상구") then %>
			<input type="button" value="새상품추가" onclick="jsAddNewItemOFF(frm, '<%= shopid %>', '<%= acURL %>');" class="button">
			<% else %>
	    	<input type="button" value="새상품추가" onclick="addnewItem('<%=onoffgubun%>',frm,'<%=shopid%>','<%=acURL%>');" class="button">
			<% end if %>
	    	<%' if shopid <> "" then %>
		        <% if obag.FresultCount>0 then %>
		            <input type="button" class="button" value="선택주문작성(텐바이텐물류)" onclick="AddArr(frmArrupdate,'<%=C_IS_SHOP%>')">
		        <% end if %>
		        <% if obag.FresultCount>0 then %>
		        	<%' if makerid <> "" or comm_cd = "B012" or comm_cd = "B022" then %>
		            	<input type="button" class="button" value="선택주문작성(업체)" onclick="AddArr_upche(frmArrupdate,'<%=C_IS_SHOP%>')">
		            <%' end if %>
		        <% end if %>
		    <%' end if %>
	    </td>
	</tr>
	</table>
	<!-- 액션 끝 -->

	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
	    <td colspan="20">
	        검색결과 : <b><%= obag.FTotalcount %></b> ※최대 300건까지 노출 됩니다.
	    </td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td>매장</td>
	    <td>
	    	공급처
	    </td>
	    <td>브랜드</td>
	    <td>상품코드</td>
	    <td>이미지</td>
	    <td>상품명<br>[옵션명]</td>
	    <td>판매가</td>

	    <% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
	    	<td>업체<br>매입가</td>
	    <% end if %>

	    <td>매장<br>공급가</td>
    	<td>유효<br>재고</td>
	    <td>
	    	판매수량<br>(7일)
	    </td>
	    <td>
	        필요수량<br>(7일)
	    </td>
	    <td>수량</td>
	    <td>등록자</td>
	    <td>비고</td>
	</tr>
	<% if obag.FresultCount > 0 then %>
	<% for i=0 to obag.FresultCount -1 %>
	<form method="get" action="" name="frmBuyPrc<%=i%>">

	<% if obag.FItemlist(i).Fisusing="N" then %>
		<tr bgcolor="#EEEEEE" align="center">
	<% else %>
		<tr bgcolor="#FFFFFF" align="center">
	<% end if %>
	<input type="hidden" name="onlinebuycash" value="<%= obag.FItemlist(i).fonlinebuycash %>">
	<input type="hidden" name="onlinemwdiv" value="<%= obag.FItemlist(i).fonlinemwdiv %>">
	<input type="hidden" name="bagidx" value="<%= obag.FItemlist(i).fbagidx %>">
	<input type="hidden" name="itemgubun" value="<%= obag.FItemlist(i).fitemgubun %>">
	<input type="hidden" name="itemid" value="<%= obag.FItemlist(i).fitemid %>">
	<input type="hidden" name="itemoption" value="<%= obag.FItemlist(i).fitemoption %>">
	<input type="hidden" name="shopitemprice" value="<%= obag.FItemlist(i).fshopitemprice %>">
	<input type="hidden" name="itemname" value="<%= obag.FItemlist(i).fshopitemname %>">
	<input type="hidden" name="itemoptionname" value="<%= obag.FItemlist(i).fshopitemoptionname %>">
	<input type="hidden" name="makerid" value="<%= obag.FItemlist(i).fmakerid %>">
	<input type="hidden" name="comm_cd" value="<%= obag.FItemlist(i).fcomm_cd %>">
	<% if IS_HIDE_BUYCASH = True then %>
	<input type="hidden" name="shopsuplycash" value="-1">
	<% else %>
	<input type="hidden" name="shopsuplycash" value="<%= obag.FItemlist(i).fshopsuplycash %>">
	<% end if %>
	<input type="hidden" name="shopbuyprice" value="<%= obag.FItemlist(i).fshopbuyprice %>">
	<input type="hidden" name="shopid" value="<%= obag.FItemlist(i).fshopid %>">
	    <td width=20>
	        <input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
	    </td>
	    <td>
	    	<%= obag.FItemlist(i).fshopname %>
	        <Br><%= obag.FItemlist(i).fshopid %>
	    </td>
	    <td width=100>
	        <%= GetdeliverGubunName(obag.FItemlist(i).fcomm_cd) %><br>(<%= obag.FItemlist(i).fcomm_name %>)
	    </td>
	    <td>
	        <a href="javascript:searchmakerid('<%= obag.FItemlist(i).fmakerid %>',frm);" onfocus="this.blur()"><%= obag.FItemlist(i).fmakerid %></a>
	    </td>
	    <td width=80>
	        <%= obag.FItemlist(i).Fitemgubun %><%=  CHKIIF(obag.FItemlist(i).Fitemid>=1000000,Format00(8,obag.FItemlist(i).Fitemid),Format00(6,obag.FItemlist(i).Fitemid)) %><%= obag.FItemlist(i).Fitemoption %>
	        <% if obag.FItemlist(i).Fitemgubun="10" then %>
	        	<Br><a href="<%=wwwUrl%>/shopping/category_prd.asp?itemid=<%=obag.FItemlist(i).Fitemid%>" target="_blink" onfocus="this.blur()">[상세]</a>
	        <% end if %>
	    </td>
	    <td width=50>
	        <img src="<%= obag.FItemlist(i).GetImageSmall %>" width=50 height=50 border=0>
	    </td>
	    <td align="left">
	        <%= obag.FItemlist(i).fshopitemname %><Br>
	        <% if obag.FItemlist(i).fshopitemoptionname <> "" then %>
	            [<%=obag.FItemlist(i).fshopitemoptionname%>]
	        <% end if %>
	    </td>
	    <td align="right" width=60>
	        <%= FormatNumber(obag.FItemlist(i).fshopitemprice,0) %>
	    </td>

	    <% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		    <td align="right" width=80>
		        <%= FormatNumber(obag.FItemlist(i).fshopsuplycash,0) %>

		        <% if obag.FItemlist(i).fcentermwdiv="M" then %>
		        	<p>ON:<%= FormatNumber(obag.FItemlist(i).fonlinebuycash,0) %>
		        <% end if %>
		    </td>
		<% end if %>

	    <td align="right" width=60>
	        <%= FormatNumber(obag.FItemlist(i).fshopbuyprice,0) %>
	    </td>
	    <td width=60>
	        <%= obag.FItemlist(i).getAvailStock %>     <!--유효재고-->
	    </td>
	    <td width=60>
	        <%= obag.FItemlist(i).fsell7days %> (7일)      <!--판매수량-->
	    </td>
	    <td width=60>
	        <!-- 총필요수량 -->
	        <% if obag.FItemlist(i).frequire7daystock > 0 then %>
	            <a href="javascript:inputiteno('<%= obag.FItemlist(i).frequire7daystock %>',frmBuyPrc<%= i %>);" onfocus="this.blur()"><p>
	            <font color="red"><%= obag.FItemlist(i).frequire7daystock %> (7일)</font>
	            </a>
	        <% else %>
	           0 (7일)
	        <% end if %>
	    </td>
	    <td width=60>
	        <input type="text" class="text" name="itemno" value="<%= obag.FItemlist(i).fitemno %>" size="4" maxlength="4" onKeyDown="CheckThis(frmBuyPrc<%= i %>);">
	    </td>
	    <td width=90>
	        <%= obag.FItemlist(i).fuserid %>
	    </td>
	    <td>
	        <% if obag.FItemList(i).Fpreorderno>0 then %>
	        	기주문:
	            <% if obag.FItemList(i).Fpreorderno<>obag.FItemList(i).Fpreordernofix then response.write CStr(obag.FItemList(i).Fpreorderno) + " -> " %>
	        	<%= obag.FItemList(i).Fpreordernofix %><br>
	        <% end if %>
	    </td>
	</tr>
	</form>
	<% next %>

	<% else %>

	<tr bgcolor="#FFFFFF">
	    <td colspan="20" align="center">[장바구니에 상품이 없습니다.]</td>
	</tr>
	<% end if %>
	<form name="frmArrupdate" method="post" action="">
	    <input type="hidden" name="mode" value="arrins">
	    <input type="hidden" name="itemgubunarr2" value="">
	    <input type="hidden" name="itemidadd2" value="">
	    <input type="hidden" name="itemoptionarr2" value="">
	    <input type="hidden" name="sellcasharr2" value="">
	    <input type="hidden" name="buycasharr2" value="">
	    <input type="hidden" name="suplycasharr2" value="">
	    <input type="hidden" name="itemnoarr2" value="">
	    <input type="hidden" name="itemnamearr2" value="">
	    <input type="hidden" name="itemoptionnamearr2" value="">
	    <input type="hidden" name="designerarr2" value="">
	    <input type="hidden" name="shopid" value="<%=shopid%>">
	    <input type="hidden" name="suplyer" value="10x10">
	    <input type="hidden" name="idx" value="0">
	    <input type="hidden" name="chargeid" value="<%=makerid%>">
	    <input type="hidden" name="shopbuypricearr2" value="">
	    <input type="hidden" name="isreq" value="Y">
	    <input type="hidden" name="bagidxarr">
	    <input type="hidden" name="cwflag">
	</form>
	<form name="frmbag" method="post" action="">
		<input type="hidden" name="mode">
		<input type="hidden" name="bagidxarr">
	    <input type="hidden" name="onoffgubun">
	    <input type="hidden" name="itemgubunarr">
	    <input type="hidden" name="itemidarr">
	    <input type="hidden" name="itemoptionarr">
	    <input type="hidden" name="itemnoarr">
	    <input type="hidden" name="makerid">
	    <input type="hidden" name="shopid" >
	</form>
	</table>

	<!-- 액션 시작 -->
	<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
	    <td align="left">
	        <input type="button" class="button" value="선택수정" onclick="bageditarr(frmbag)">
	        <input type="button" class="button" value="선택삭제" onclick="bagdelarr(frmbag)">
	    </td>
	    <td align="right">
	    	<input type="button" value="새상품추가" onclick="addnewItem('<%=onoffgubun%>',frm,'<%=shopid%>','<%=acURL%>');" class="button">
	    	<%' if shopid <> "" then %>
		        <% if obag.FresultCount>0 then %>
		            <input type="button" class="button" value="선택주문작성(텐바이텐물류)" onclick="AddArr(frmArrupdate,'<%=C_IS_SHOP%>')">
		        <% end if %>
		        <% if obag.FresultCount>0 then %>
		        	<%' if makerid <> "" or comm_cd = "B012" or comm_cd = "B022" then %>
		            	<input type="button" class="button" value="선택주문작성(업체)" onclick="AddArr_upche(frmArrupdate,'<%=C_IS_SHOP%>')">
		            <%' end if %>
		        <% end if %>
		    <%' end if %>
	    </td>
	</tr>
	</table>
	<!-- 액션 끝 -->
<% end if %>
<iframe id="view" name="view" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<%
set obag = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
