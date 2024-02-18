<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs센터
' Hieditor : 2009.04.17 이상구 생성
'			 2016.07.21 한용민 수정
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_giftcard_ordercls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/sp_tenGiftCardCls.asp" -->
<%

'// 매장에서도 접속가능하도록 수정, skyer9, 2018-01-15

dim searchfield, userid, giftorderserial, username, userhp, etcfield, etcstring
dim checkYYYYMMDD
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim jumundiv, jumunsite
dim research
dim AlertMsg
dim shopid, showShopID, fixShopID, cpnreguserid

'// ============================================================================
showShopID = False
fixShopID = False

if C_ADMIN_USER then

'// 직영/가맹점
elseif (C_IS_SHOP) then
	showShopID = True
	'// 어드민권한 점장 미만
	if getlevel_sn("",session("ssBctId")) > 3 then
		fixShopID = True
		shopid = C_STREETSHOPID
	else
		shopid 	= requestCheckvar(request("shopid"),32)
		if (shopid = "") then
			shopid = C_STREETSHOPID
		end if
	end if
end if

if C_InspectorUser then
	showShopID = True
	fixShopID = False
	shopid = "streetshop011"
elseif session("ssBctBigo") <> "" then
	showShopID = True
	fixShopID = True
	shopid = session("ssBctBigo")
end if


'==============================================================================
searchfield = request("searchfield")
userid 		= requestCheckvar(request("userid"),32)
giftorderserial = requestCheckvar(request("giftorderserial"),32)
username 	= requestCheckvar(request("username"),32)
userhp 		= requestCheckvar(request("userhp"),32)
etcfield 	= requestCheckvar(request("etcfield"),32)
etcstring 	= requestCheckvar(request("etcstring"),32)
cpnreguserid 		= requestCheckvar(request("cpnreguserid"),32)
checkYYYYMMDD = request("checkYYYYMMDD")

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

jumundiv = request("jumundiv")
jumunsite = request("jumunsite")
research = request("research")


if (research="") and (checkYYYYMMDD="") then checkYYYYMMDD="Y"

if (research="") and (shopid <> "") then
	searchfield = "shopid"
end if
'==============================================================================
dim nowdate, searchnextdate


''기본 N달. 디폴트 체크
if (yyyy1="") then
    nowdate = Left(CStr(dateadd("m",-1,now())),10)
	yyyy1   = Left(nowdate,4)
	mm1     = Mid(nowdate,6,2)
	dd1     = Mid(nowdate,9,2)

	nowdate = Left(CStr(now()),10)
	yyyy2   = Left(nowdate,4)
	mm2     = Mid(nowdate,6,2)
	dd2     = Mid(nowdate,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2,mm2,dd2),1)),10)


'==============================================================================
dim page
dim oGiftOrder

page = request("page")
if (page="") then page=1

set oGiftOrder = new cGiftCardOrder

oGiftOrder.FPageSize = 10
oGiftOrder.FCurrPage = page

if (checkYYYYMMDD="Y") then
	oGiftOrder.FRectRegStart = Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
	oGiftOrder.FRectRegEnd = searchnextdate
end if

if (searchfield = "giftorderserial") then
        '주문번호
        oGiftOrder.FRectGiftOrderSerial = giftorderserial
elseif (searchfield = "userid") then
        '고객아이디
        oGiftOrder.FRectUserID = userid
elseif (searchfield = "cpnreguserid") then
        '수령인아이디
        oGiftOrder.FRectcpnreguserid = cpnreguserid
elseif (searchfield = "shopid") then
        '매장아이디
        oGiftOrder.FRectUserID = shopid
elseif (searchfield = "username") then
        '구매자명
        oGiftOrder.FRectBuyname = username
elseif (searchfield = "userhp") then
        '구매자핸드폰
        oGiftOrder.FRectBuyHp = userhp
elseif (searchfield = "etcfield") then
        '기타조건
        if etcfield="04" then
        	oGiftOrder.FRectIpkumName = etcstring
        elseif etcfield="07" then
        	oGiftOrder.FRectBuyPhone = etcstring
        elseif etcfield="08" then
        	oGiftOrder.FRectReqHp = etcstring
        end if
end if

dim ix,iy

oGiftOrder.getCSGiftcardOrderList

'' 검색결과가 1개일대 디테일 자동으로 뿌림
dim ResultOneOrderserial
ResultOneOrderserial = ""
if (oGiftOrder.FResultCount=1) then
    ResultOneOrderserial = oGiftOrder.FItemList(0).FgiftOrderSerial
end if

%>
<script language='javascript'>
function copyClipBoard(itxt) {
	var posSpliter = itxt.indexOf("|");

	try{
	    parent.callring.frm.giftorderserial.value=itxt.substring(0,posSpliter);
	    parent.callring.frm.userid.value=itxt.substring(posSpliter+1,255);
	}catch(ignore){

	}
}

function GotoOrderDetail(giftorderserial) {
        parent.detailFrame.location.href = "cs_giftcard_OrderDetail.asp?giftorderserial=" + giftorderserial;
}

function ChangeCheckbox(frmname, frmvalue) {
        for (var i = 0; i < frm.elements.length; i++) {
                if (frm.elements[i].type == "radio") {
                        if ((frm.elements[i].name == frmname) && (frm.elements[i].value == frmvalue)) {
                                frm.elements[i].checked = true;
                        }
                }
        }
}

function FocusAndSelect(frm, obj){
        ChangeFormBgColor(frm);

        obj.focus();
        obj.select();
}

function ChangeFormBgColor(frm) {
        // style='background-color:#DDDDFF'
        var radioselected = false;
        var checkboxchecked = false;
        var ischecked = false;

        for (var i = 0; i < frm.elements.length; i++) {
                if (frm.elements[i].type == "radio") {
                        ischecked = frm.elements[i].checked;
                }

                if (frm.elements[i].type == "checkbox") {
                        ischecked = frm.elements[i].checked;
                }

                if (frm.elements[i].type == "text") {
                        if (ischecked == true) {
                                frm.elements[i].style.background = "FFFFCC";
                        } else {
                                frm.elements[i].style.background = "EEEEEE";
                        }
                }

                if (frm.elements[i].type == "select-one") {
                        if (ischecked == true) {
                                frm.elements[i].style.background = "FFFFCC";
                        } else {
                                frm.elements[i].style.background = "EEEEEE";
                        }
                }
        }
}

// tr 색상변경
var pre_selected_row = null;
var pre_selected_row_color = null;

function ChangeColor(e, selcolor, defcolor){
	if (pre_selected_row_color != null) {
	        pre_selected_row.bgColor = pre_selected_row_color;
        }
        pre_selected_row = e;
        pre_selected_row_color = defcolor;
        e.bgColor = selcolor;
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function popCardSellReg() {
	var popwin = window.open('pop_off_giftcard_sell_reg.asp?shopid=<%= shopid %>','popCardSellReg','width=1300, height=720, scrollbars=yes, resizable=yes');
	popwin.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#F4F4F4">
	    <td width="50" bgcolor="#EEEEEE">검색<br>조건</td>
		<td align="left">
			<%
			if (showShopID = True) then
				%>
			<table width="100%" align="center" cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<td align="left">
						<%
				if (fixShopID = True) then
					%>
					샵아이디 : <%= shopid %><input type="hidden" name="shopid" value="<%= shopid %>">
					<%
				else
					%>
			<input type="radio" name="searchfield" value="shopid" checked onClick="FocusAndSelect(frm, frm.shopid)"> 샵아이디
    		<input type="text" class="text" name="shopid" value="<%= shopid %>" size="13" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'shopid'); FocusAndSelect(frm, frm.shopid);">
					<%
				end if
						%>
					</td>
					<td align="right">
						<input type="button" class="button_s" value="판매등록" onClick="popCardSellReg();">
					</td>
				</tr>
			</table>
			<% else %>
			<input type="radio" name="searchfield" value="giftorderserial" <% if searchfield="giftorderserial" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.giftorderserial)"> 주문번호
    		<input type="text" class="text" name="giftorderserial" value="<%= giftorderserial %>" size="13" maxlength="11" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'giftorderserial'); FocusAndSelect(frm, frm.giftorderserial);">

    		<input type="radio" name="searchfield" value="userid" <% if searchfield="userid" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.userid)"> 아이디
    		<input type="text" class="text" name="userid" value="<%= userid %>" size="12" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'userid'); FocusAndSelect(frm, frm.userid);">

    		<input type="radio" name="searchfield" value="cpnreguserid" <% if searchfield="cpnreguserid" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.cpnreguserid)"> 수령인아이디
    		<input type="text" class="text" name="cpnreguserid" value="<%= cpnreguserid %>" size="12" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'cpnreguserid'); FocusAndSelect(frm, frm.cpnreguserid);">

    		<input type="radio" name="searchfield" value="username" <% if searchfield="username" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.username)"> 구매자명
    		<input type="text" class="text" name="username" value="<%= username %>" size="8" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'username'); FocusAndSelect(frm, frm.username);">

    		<input type="radio" name="searchfield" value="userhp" <% if searchfield="userhp" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.userhp)"> 구매자핸드폰
    		<input type="text" class="text" name="userhp" value="<%= userhp %>" size="14" maxlength="14" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'userhp'); FocusAndSelect(frm, frm.userhp);">
            <input type="radio" name="searchfield" value="etcfield" <% if searchfield="etcfield" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.etcstring)"> 기타조건

    		<select name="etcfield" class="select">
    			  <option value="">선택</option>
                  <option value="04" <% if etcfield="04" then response.write "selected" %> >입금자명</option>
                  <option value="07" <% if etcfield="07" then response.write "selected" %> >구매자 전화</option>
                  <option value="08" <% if etcfield="08" then response.write "selected" %> >수령인 핸드폰</option>
            </select>
    		<input type="text" class="text" name="etcstring" value="<%= etcstring %>" size="14" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'etcfield'); FocusAndSelect(frm, frm.etcstring);">
			<br />
					<%
			end if
			%>
    		<input type="checkbox" name="checkYYYYMMDD" value="Y" <% if checkYYYYMMDD="Y" then response.write "checked" %> onClick="ChangeFormBgColor(frm)">
    		주문일 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<% 'if C_ADMIN_AUTH then %>
		<input type="button" class="button_s" value="판매등록" onClick="popCardSellReg();">
		<% 'end if %>
	    </td>
	    <td width="50" bgcolor="#EEEEEE">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->


<p>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="30">구분</td>
    	<td width="70">주문번호</td>
    	<td width="100">UserID</td>
    	<td width="70">구매자</td>
    	<td width="100">기프티콘Pin</td>
    	<td width="100">교환UserID</td>
    	<td>카드명</td>

    	<td width="60">판매가</td>

    	<td width="60"><b>실결제액</b></td>

    	<td width="100">결제방법</td>
    	<td width="50">주문상태</td>
    	<td width="50">입금구분</td>
    	<td width="50">카드상태</td>
    	<td width="70">주문일</td>
    	<td width="70">입금확인일</td>
    </tr>
    <% if oGiftOrder.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="15" align="center">[검색결과가 없습니다.]</td>
    </tr>
    <% else %>

	<% for ix=0 to oGiftOrder.FresultCount-1 %>

	<% if oGiftOrder.FItemList(ix).IsValidOrder then %>
	<tr align="center" bgcolor="#FFFFFF" class="a" onclick="ChangeColor(this,'#AFEEEE','FFFFFF'); copyClipBoard('<%= oGiftOrder.FItemList(ix).FgiftOrderSerial %>|<%= oGiftOrder.FItemList(ix).FUserID %>'); GotoOrderDetail('<%= oGiftOrder.FItemList(ix).FgiftOrderSerial %>'); " style="cursor:hand">
	<% else %>
	<tr align="center" bgcolor="#EEEEEE" class="gray" onclick="ChangeColor(this,'#AFEEEE','EEEEEE'); copyClipBoard('<%= oGiftOrder.FItemList(ix).FgiftOrderSerial %>|<%= oGiftOrder.FItemList(ix).FUserID %>'); GotoOrderDetail('<%= oGiftOrder.FItemList(ix).FgiftOrderSerial %>'); " style="cursor:hand">
	<% end if %>
		<td><font color="<%= oGiftOrder.FItemList(ix).CancelYnColor %>"><%= oGiftOrder.FItemList(ix).CancelYnName %></font></td>
		<td><%= oGiftOrder.FItemList(ix).FgiftOrderSerial %></td>
		<td align="left">
		    <!--<a href="?searchfield=userid&userid=<%'= oGiftOrder.FItemList(ix).FUserID %>">-->
		    	<font color="<%= getUserLevelColorByDate(oGiftOrder.FItemList(ix).FUserLevel, Left(oGiftOrder.FItemList(ix).FRegDate,10)) %>">
		    	<b><%= printUserId(oGiftOrder.FItemList(ix).FUserID, 2, "*") %></b></font>
		    <!--</a>-->
		</td>
		<td><%= oGiftOrder.FItemList(ix).FBuyName %></td>
        <td><%= oGiftOrder.FItemList(ix).FcouponNo %></td>
		<td><%= oGiftOrder.FItemList(ix).FcpnRegUserID %></td>
		<td><%= oGiftOrder.FItemList(ix).FCarditemname %></td>

		<td align="right">
			<%= FormatNumber(oGiftOrder.FItemList(ix).Ftotalsum,0) %>
		</td>

		<td align="right"><b><%= FormatNumber((oGiftOrder.FItemList(ix).FSubTotalPrice),0) %></b></td>


		<td><%= oGiftOrder.FItemList(ix).GetAccountdivName %></td>
		<% if (oGiftOrder.FItemList(ix).Fipkumdiv="0") or oGiftOrder.FItemList(ix).Fipkumdiv="1"  then %>
		<td><font color="<%= oGiftOrder.FItemList(ix).GetJumunDivColor %>"><acronym title="<%= oGiftOrder.FItemList(ix).Fresultmsg %>"><%= oGiftOrder.FItemList(ix).GetJumunDivName %></acronym></font></td>
		<% else %>
		<td><font color="<%= oGiftOrder.FItemList(ix).GetJumunDivColor %>"><%= oGiftOrder.FItemList(ix).GetJumunDivName %></font></td>
		<% end if %>
		<td><font color="<%= oGiftOrder.FItemList(ix).IpkumDivColor %>"><%= oGiftOrder.FItemList(ix).GetIpkumDivName %></font></td>
			<td><font color="<%= oGiftOrder.FItemList(ix).GetCardStatusColor %>"><%= oGiftOrder.FItemList(ix).GetCardStatusName %></font></td>
		<td><acronym title="<%= oGiftOrder.FItemList(ix).FRegDate %>"><%= Left(oGiftOrder.FItemList(ix).FRegDate,10) %></acronym></td>
		<td><acronym title="<%= oGiftOrder.FItemList(ix).Fipkumdate %>"><%= Left(oGiftOrder.FItemList(ix).Fipkumdate,10) %></acronym></td>
	</tr>
	<% next %>

<% end if %>

    <tr align="center" bgcolor="#FFFFFF">
        <td colspan="15">
            <% if oGiftOrder.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oGiftOrder.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for ix=0 + oGiftOrder.StartScrollPage to oGiftOrder.FScrollCount + oGiftOrder.StartScrollPage - 1 %>
    			<% if ix>oGiftOrder.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(ix) then %>
    			<font color="red">[<%= ix %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
    			<% end if %>
    		<% next %>

    		<% if oGiftOrder.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
        </td>
    </tr>
</table>
<script language='javascript'>
    ChangeFormBgColor(frm);

    <% if ResultOneOrderserial<>"" then %>
    GotoOrderDetail('<%= ResultOneOrderserial %>')
    <% end if %>
</script>
<!-- 표 하단바 끝-->
<%
set oGiftOrder = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
