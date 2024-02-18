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
dim searchfield, Orderno, etcfield, etcstring ,checkYYYYMMDD ,oaslistmaejang ,oaslistfinal
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2 , AlertMsg , research ,nowdate, searchnextdate ,maejangascount, finalascount
dim page ,ojumun , ix,iy , shopid ,ResultOnemasteridx ,datefg ,onlineuserid
	shopid = requestCheckVar(request("shopid"),32)
	searchfield = requestCheckVar(request("searchfield"),32)
	Orderno = requestCheckvar(request("Orderno"),16)
	etcfield 	= requestCheckvar(request("etcfield"),32)
	etcstring 	= requestCheckvar(request("etcstring"),32)
	checkYYYYMMDD = requestCheckVar(request("checkYYYYMMDD"),10)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)	
	research = requestCheckVar(request("research"),2)
	page = requestCheckVar(request("page"),10)
	datefg = requestCheckVar(request("datefg"),10)
	onlineuserid 	= requestCheckvar(request("onlineuserid"),32)

if datefg = "" then datefg = "maechul"
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

if (C_IS_SHOP) then
	'직영/가맹점
	shopid = C_STREETSHOPID
else
	if (C_IS_Maker_Upche) then
		'makerid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
		else
		end if
	end if
end if

set ojumun = new COrder
	ojumun.FPageSize = 10
	ojumun.FCurrPage = page
	ojumun.frectshopid = shopid
	ojumun.frectdatefg = datefg
	
	if (checkYYYYMMDD="Y") then
		ojumun.FRectRegStart = Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
		ojumun.FRectRegEnd = searchnextdate
	end if
	
	if (searchfield = "Orderno") then
	    '주문번호
	    ojumun.FRectOrderno = Orderno
	elseif (searchfield = "onlineuserid") then
	    '온라인iD
	    ojumun.frectuserid = onlineuserid	      
	elseif (searchfield = "etcfield") then
	    '기타조건
	    if etcfield="01" then
	    	ojumun.FRectBuyname = etcstring
	    elseif etcfield="02" then
	    	ojumun.FRectBuyHp = etcstring
	    elseif etcfield="03" then
	    	ojumun.FRectmail = etcstring
	    elseif etcfield="04" then
	    	ojumun.FrectCardNo = etcstring	    	
	    end if
	end if
	
	''검색조건 없을때 최근 N건 검색
	ojumun.fQuickSearchOrderList

'' 검색결과가 1개일대 디테일 자동으로 뿌림
if (ojumun.FResultCount=1) then
    ResultOnemasteridx = ojumun.FItemList(0).Fmasteridx
end if

'/매장처리 대상건수
set oaslistmaejang = new COrder
    oaslistmaejang.frectcurrstate = "'B001','B004'"
    oaslistmaejang.frectdeleteyn = "N"
	oaslistmaejang.frectshopid = shopid    
    oaslistmaejang.fGetCSASTotalCount
	
    maejangascount = oaslistmaejang.FResultCount

'/최종완료처리 대상건수
set oaslistfinal = new COrder
    oaslistfinal.frectcurrstate = "'B006','B008'"
    oaslistfinal.frectdeleteyn = "N"
	oaslistfinal.frectshopid = shopid    
    oaslistfinal.fGetCSASTotalCount
	
    finalascount = oaslistfinal.FResultCount
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
		<input type="radio" name="searchfield" value="onlineuserid" <% if searchfield="onlineuserid" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.onlineuserid)"> 온라인ID
		<input type="text" class="text" name="onlineuserid" value="<%= onlineuserid %>" size="16" maxlength="16" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'onlineuserid'); FocusAndSelect(frm, frm.onlineuserid);">		
        <input type="radio" name="searchfield" value="etcfield" <% if searchfield="etcfield" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.etcstring)"> 포인트카드회원검색
		<select name="etcfield" class="select">
			<option value="">선택</option>
			<option value="01" <% if etcfield="01" then response.write "selected" %> >이름</option>                          
			<option value="02" <% if etcfield="02" then response.write "selected" %> >휴대폰</option>
			<option value="03" <% if etcfield="03" then response.write "selected" %> >이메일</option>  
			<option value="04" <% if etcfield="04" then response.write "selected" %> >카드번호</option>              	         
		</select>
		<input type="text" class="text" name="etcstring" value="<%= etcstring %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'etcfield'); FocusAndSelect(frm, frm.etcstring);">
		<br>
		<input type="checkbox" name="checkYYYYMMDD" value="Y" <% if checkYYYYMMDD="Y" then response.write "checked" %> onClick="ChangeFormBgColor(frm)">
		<% drawmaechuldatefg "datefg" ,datefg ,""%><% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>            
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
<tr bgcolor="#FFFFFF">
	<td colspan="13">검색결과 : <b><%=ojumun.FTotalCount%></b>&nbsp;&nbsp;페이지 : <b><%= page %>/ <%= ojumun.FTotalPage %></b></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>idx</td>
	<td>구분</td>
	<td>주문<br>매장</td>
	<td>주문<br>번호</td>
	<td>총금액</td>
	<td>총결제<br>금액</td>
	<td>현금<br>결제</td>
	<td>카드<br>결제</td>
	<td>주문일</td>
	<td>
		포인트카드
		<br>고객정보
	</td>
	<td>온라인ID</td>
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
	<td><%= FormatNumber(ojumun.FItemList(ix).ftotalsum,0) %></td>
	<td><%= FormatNumber(ojumun.FItemList(ix).frealsum,0) %></td>
	<td><%= FormatNumber(ojumun.FItemList(ix).fcashsum,0) %></td>
	<td><%= FormatNumber(ojumun.FItemList(ix).fcardsum,0) %></td>
	<td><acronym title="<%= ojumun.FItemList(ix).FRegDate %>"><%= Left(ojumun.FItemList(ix).FRegDate,10) %></acronym></td>
	<td>
		<%= ojumun.FItemList(ix).fpointuserno %>
		<% if ojumun.FItemList(ix).Fbuyname <> "" then %>
			(<%= ojumun.FItemList(ix).Fbuyname %>)
		<% end if %>
	</td>
	<td>
		<%= printUserId(ojumun.FItemList(ix).fonlineuserid, 2, "*") %>
	</td>
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

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:5;">
<tr>
	<td align="left">
		<% if shopid <> "" then %>
			<%= shopid %>매장 주문중 :
		<% else %>
			전체매장 주문중 :
		<% end if %>
		<a href="javascript:PopmaejangAction('','<%= shopid %>','','notfinish');" onfocus="this.blur()">매장완료처리대상 : <font color="red"><%=maejangascount%>건</font></a> 
		/ <a href="javascript:Cscenter_Action_List_off('','','','notfinal','<%= shopid %>');" onfocus="this.blur()">최종완료처리대상 : <font color="red"><%= finalascount %>건</font></a>
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->		

<script language='javascript'>
    ChangeFormBgColor(frm);

    <% if ResultOnemasteridx<>"" then %>
    	GotoOrderDetail('<%= ResultOnemasteridx %>')
    <% end if %>

    <% if (AlertMsg<>"") then %>
        alert('<%= AlertMsg %>');
    <% end if %>
</script>

<%
set ojumun = Nothing
set oaslistmaejang = Nothing
set oaslistfinal = Nothing
%>
<!-- #include virtual="/admin/offshop/shopcscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->