<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 고객센터
' Hieditor : 2011.03.07 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<% '<!-- #include virtual="/lib/checkAllowIPWithLog.asp" --> %>
<!-- #include virtual="/lib/classes/offshop/order/order_cls.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->
<%
dim searchfield, Orderno, username, userhp, etcfield, etcstring ,checkYYYYMMDD
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2 , AlertMsg , research ,nowdate, searchnextdate
dim page ,ojumun , ix,iy , shopid ,ResultOnemasteridx
	shopid = requestCheckVar(request("shopid"),32)
	searchfield = requestCheckVar(request("searchfield"),32)
	Orderno = requestCheckvar(request("Orderno"),16)
	username 	= requestCheckvar(request("username"),32)
	userhp 		= requestCheckvar(request("userhp"),16)
	etcfield 	= requestCheckvar(request("etcfield"),10)
	etcstring 	= requestCheckvar(request("etcstring"),32)
	checkYYYYMMDD = requestCheckVar(request("checkYYYYMMDD"),1)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	research = requestCheckVar(request("research"),2)
	page = requestCheckVar(request("page"),10)

if (page="") then page=1
	
if (research="") and (checkYYYYMMDD="") then checkYYYYMMDD="Y"
ResultOnemasteridx = ""
	
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

set ojumun = new COrder
	ojumun.FPageSize = 10
	ojumun.FCurrPage = page
	ojumun.frectshopid = shopid
	
	if (checkYYYYMMDD="Y") then
		ojumun.FRectRegStart = Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
		ojumun.FRectRegEnd = searchnextdate
	end if
	
	if (searchfield = "Orderno") then
	    '주문번호
	    ojumun.FRectOrderno = Orderno
	elseif (searchfield = "username") then
	    '구매자명
	    ojumun.FRectBuyname = username
	elseif (searchfield = "userhp") then
	    '구매자핸드폰
	    ojumun.FRectBuyHp = userhp	    
	elseif (searchfield = "etcfield") then
	    '기타조건
	    if etcfield="02" then
	    	ojumun.FRectReqName = etcstring
	    elseif etcfield="07" then
	    	ojumun.FRectBuyPhone = etcstring
	    elseif etcfield="08" then
	    	ojumun.FRectReqHp = etcstring
	    elseif etcfield="10" then
	    	ojumun.FRectReqPhone = etcstring
	    end if
	end if
	
	''검색조건 없을때 최근 N건 검색
	ojumun.fQuickSearchOrderList

'' 검색결과가 1개일대 디테일 자동으로 뿌림
if (ojumun.FResultCount=1) then
    ResultOnemasteridx = ojumun.FItemList(0).Fmasteridx
end if
%>

<script language='javascript'>

function copyClipBoard(itxt) {
	var posSpliter = itxt.indexOf("|");

	try{
	    parent.callring.frm.Orderno.value=itxt.substring(0,posSpliter);
	}catch(ignore){
	}
}

function GotoOrderDetail(masteridx) {
	parent.detailFrame.location.href = "ordermaster_detail.asp?masteridx=" + masteridx;
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
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

</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="F4F4F4">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr>
    <td>
		<input type="radio" name="searchfield" value="Orderno" <% if searchfield="Orderno" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.Orderno)"> 주문번호
		<input type="text" class="text" name="Orderno" value="<%= Orderno %>" size="16" maxlength="16" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'Orderno'); FocusAndSelect(frm, frm.Orderno);">
		<input type="radio" name="searchfield" value="username" <% if searchfield="username" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.username)"> 수령인명
		<input type="text" class="text" name="username" value="<%= username %>" size="8" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'username'); FocusAndSelect(frm, frm.username);">
		<input type="radio" name="searchfield" value="userhp" <% if searchfield="userhp" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.userhp)"> 구매자핸드폰
		<input type="text" class="text" name="userhp" value="<%= userhp %>" size="14" maxlength="14" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'userhp'); FocusAndSelect(frm, frm.userhp);">
        <input type="radio" name="searchfield" value="etcfield" <% if searchfield="etcfield" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.etcstring)"> 기타조건
		<select name="etcfield" class="select">
			  <option value="">선택</option>
              <option value="02" <% if etcfield="02" then response.write "selected" %> >수령인명</option>                          
              <option value="07" <% if etcfield="07" then response.write "selected" %> >구매자 전화</option>
              <option value="10" <% if etcfield="10" then response.write "selected" %> >수령인 전화</option>
              <option value="08" <% if etcfield="08" then response.write "selected" %> >수령인 핸드폰</option>              
            </select>
		<input type="text" class="text" name="etcstring" value="<%= etcstring %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'etcfield'); FocusAndSelect(frm, frm.etcstring);">
		<br>
		<input type="checkbox" name="checkYYYYMMDD" value="Y" <% if checkYYYYMMDD="Y" then response.write "checked" %> onClick="ChangeFormBgColor(frm)">
		주문일 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>            
		ShopID : 
		<% 'drawSelectBoxOffShop "shopid",shopid %>
		<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
    </td>
    <td align="right" valign="top">
        <input type="button" class="button_s" value="새로고침" onclick="document.location.reload();">
        &nbsp;
        <input type="button" class="button_s" value="검색하기" onclick="document.frm.submit();">
    </td>
</tr>
</form>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>idx</td>
	<td>구분</td>
	<td>주문매장</td>
	<td>주문번호</td>
	<td>구매자</td>
	<td>수령인</td>
	<td>거래상태</td>
	<td>주문일</td>
	<td>발주일</td>
</tr>
<% if ojumun.FresultCount > 0 then %>
<% for ix=0 to ojumun.FresultCount-1 %>
<% if ojumun.FItemList(ix).IsAvailJumun then %>
<tr align="center" bgcolor="#FFFFFF" class="a" onclick="ChangeColor(this,'#AFEEEE','FFFFFF'); copyClipBoard('<%= ojumun.FItemList(ix).FOrderno %>'); GotoOrderDetail('<%= ojumun.FItemList(ix).Fmasteridx %>'); " style="cursor:hand">
<% else %>
<tr align="center" bgcolor="#EEEEEE" class="gray" onclick="ChangeColor(this,'#AFEEEE','EEEEEE'); copyClipBoard('<%= ojumun.FItemList(ix).FOrderno %>'); GotoOrderDetail('<%= ojumun.FItemList(ix).Fmasteridx %>'); " style="cursor:hand">
<% end if %>
	<td>
	    <%= ojumun.FItemList(ix).fmasteridx %>
	</td>
	<td><font color="<%= ojumun.FItemList(ix).CancelYnColor %>"><%= ojumun.FItemList(ix).CancelYnName %></font></td>
	<td>
	    <%= ojumun.FItemList(ix).fshopname %>
	</td>
	<td><%= ojumun.FItemList(ix).FOrderno %></td>		
	<td>
		<%= ojumun.FItemList(ix).FBuyName %>
	</td>
	<td>
		<%= ojumun.FItemList(ix).FReqName %>
	</td>
	<td><font color="<%= ojumun.FItemList(ix).shopIpkumDivColor %>"><%= ojumun.FItemList(ix).shopIpkumDivName %></font></td>
	<td><acronym title="<%= ojumun.FItemList(ix).FRegDate %>"><%= Left(ojumun.FItemList(ix).FRegDate,10) %></acronym></td>
	<td><acronym title="<%= ojumun.FItemList(ix).Fbaljudate %>"><%= Left(ojumun.FItemList(ix).Fbaljudate,10) %></acronym></td>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="20">
        <% if ojumun.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for ix=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
			<% if ix>ojumun.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(ix) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if ojumun.HasNextScroll then %>
			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
    </td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
</table>
<!-- 표 하단바 끝-->

<script language='javascript'>
    ChangeFormBgColor(frm);

    <% if ResultOnemasteridx<>"" then %>
    GotoOrderDetail('<%= ResultOnemasteridx %>')
    // top.detailFrame.location.href = "ordermaster_detail.asp?Orderno=<%= ResultOnemasteridx %>";
    <% end if %>

    <% if (AlertMsg<>"") then %>
        alert('<%= AlertMsg %>');
    <% end if %>
</script>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/offshop/cscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->