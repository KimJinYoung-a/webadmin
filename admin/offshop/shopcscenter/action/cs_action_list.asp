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
dim i, username, masteridx, makerid, searchfield, searchstring ,ix , orderno
dim searchtype, divcd, currstate
Dim writeUser ,delYN, toDate ,page ,ocsaslist ,ResultOneCsID ,shopid
	delYN	= requestCheckVar(req("delYN",""),1)
	username = requestCheckVar(request("username"),32)
	masteridx = requestCheckVar(request("masteridx"),10)
	orderno = requestCheckVar(request("orderno"),16)
	searchfield = requestCheckVar(request("searchfield"),32)
	searchstring = requestCheckVar(request("searchstring"),128)
	searchtype = requestCheckVar(request("searchtype"),32)
	divcd = requestCheckVar(request("divcd"),4)
	currstate = requestCheckVar(request("currstate"),4)
	page = requestCheckVar(request("page"),10)
	shopid = requestCheckVar(request("shopid"),32)

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

	if page="" then page=1		
	if searchtype="searchfield" and searchfield="" then searchstring="" end if
	if searchtype="" then searchtype="searchfield"

'orderno 에 데이타가 파라미터가 왔을때는 해당 파라미터로 세팅하고
'없을경우는 searchstring 에 데이타가 있는지를 확인하여 세팅한다.
'다른 페이지에서 링크를 걸어 팝업을 열었을때에 대한 처리.

if (orderno <> "") then
    searchtype = "searchfield"
    username = ""
    searchfield = "orderno"
    searchstring = orderno
    divcd = ""
    currstate = ""
else
    if (searchstring <> "") then
        if (searchfield = "orderno") then
                username = ""
                orderno = searchstring
                makerid = ""
        elseif (searchfield = "makerid") then
                username = ""
                orderno = ""
                makerid = searchstring

		elseif (searchfield = "writeUser") then
                writeUser = searchstring
		else
                username = searchstring
                orderno = ""
                makerid = ""
        end If       		
    else
        username = ""        
        searchfield = ""
        searchstring = ""
    end if
end if
	
set ocsaslist = New COrder
	ocsaslist.FPageSize = 10
	ocsaslist.FCurrPage = page
	
	if (searchtype = "searchfield") then
		ocsaslist.FRectSearchType = searchtype
	    ocsaslist.FRectUserName = username
	    ocsaslist.FRectmasteridx = masteridx
	    ocsaslist.FRectorderno = orderno
	    ocsaslist.FRectMakerid  = makerid	
	    ocsaslist.FRectDivcd = divcd
	    ocsaslist.FRectCurrstate = currstate
	    ocsaslist.FRectWriteUser = writeUser	
	    ocsaslist.FRectDeleteYN	= delYN
	    ocsaslist.FRectshopid	= shopid
	end if

	ocsaslist.fGetCSASMasterList()

if ocsaslist.FResultCount=1 then
    ResultOneCsID = ocsaslist.FItemList(0).fmasteridx
end if
%>

<script language='javascript'>

var pre_selected_row = null;

function ChangeColor(e, selcolor, defcolor){
	if (pre_selected_row != null) {
	        pre_selected_row.bgColor = defcolor;
        }
        pre_selected_row = e;
        e.bgColor = selcolor;
}

function searchDetail(masteridx){
    buffrm.masteridx.value = masteridx;
    buffrm.submit();
}

function NextPage(page){
    frm.page.value = page;
    frm.submit();
}

function reSearch(){
    frm.page.value="1";
    frm.submit();
}

function reSearchByorderno(iorderno){
    frm.searchfield[4].selected = true;
    frm.searchstring.value = iorderno;
    frm.page.value="1";
    frm.submit();
}

function reSearchByMakerid(imakerid){
    frm.searchfield[3].selected = true;
    frm.searchstring.value = imakerid;
    frm.page.value="1";
    frm.submit();
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

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="cs_action_list.asp" >
<input type="hidden" name="page" value="1">
<input type="hidden" name="masteridx" value="<%=masteridx%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
        <select class="select" name="searchfield">
        	<option value="" <% if (searchfield = "") then %>selected<% end if %>>전체</option>
			<option value="masteridx" <% if (searchfield = "masteridx") then %>selected<% end if %>>일렬변호</option>
			<option value="username" <% if (searchfield = "username") then %>selected<% end if %>>고객명</option>
			<option value="makerid" <% if (searchfield = "makerid") then %>selected<% end if %>>업체처리아이디</option>
			<option value="orderno" <% if (searchfield = "orderno") then %>selected<% end if %>>주문번호</option>
			<option value="username" <% if (searchfield = "username") then %>selected<% end if %>>고객명</option>
			<option value="makerid" <% if (searchfield = "makerid") then %>selected<% end if %>>업체처리아이디</option>
			<option value="writeUser" <% if (searchfield = "writeUser") then %>selected<% end if %>>접수자아이디</option>
        </select>
        <input type="text" class="text" name="searchstring" value="<%= searchstring %>" size="20" onFocus="ChangeCheckbox('searchtype', 'searchfield'); this.style.background = '#FFFFFF'">
        &nbsp;
        구분:
        <select class="select" name="divcd">
        	<option value="">전체</option>
			<option value="A030" <% if (divcd = "A030") then response.write "selected" end if %>>업체A/S</option>
			<option value="A031" <% if (divcd = "A031") then response.write "selected" end if %>>업체A/S(매장회수)</option>
        </select>
        &nbsp;
        진행상태: <% drawcurrstate "currstate" ,currstate ,"" %>
        <Br>
        매장 : <% drawSelectBoxOffShop "shopid",shopid %>
        <input type="checkbox" name="delYN" value="N" <%if (delYN="N") then %>checked<% end if %>>삭제(취소)제외     
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="reSearch();">
		<Br><input type="button" class="button_s" value="새로고침" onclick="document.location.reload();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">       	
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
	
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">			
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" >
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>Idx</td>
    <td>구분</td>
    <td>관련주문번호</td>    
    <td>업체ID</td>
    <td>고객명<br>(매장명)</td>    
    <td>제목</td>
    <td>상태</td>    
    <td>등록일</td>
    <td>처리일</td>
    <td>삭제</td>
</tr>
<% if ocsaslist.FResultCount > 0 then %>
<% for i = 0 to (ocsaslist.FResultCount - 1) %>
<% if (ocsaslist.FItemList(i).Fdeleteyn <> "N") then %>
<tr bgcolor="#EEEEEE" style="color:gray" align="center" onclick="ChangeColor(this,'AFEEEE','FFFFFF'); searchDetail('<%= ocsaslist.FItemList(i).fmasteridx %>');" style="cursor:hand">
<% else %>
<tr bgcolor="#FFFFFF" align="center" onclick="ChangeColor(this,'AFEEEE','FFFFFF'); searchDetail('<%= ocsaslist.FItemList(i).fmasteridx %>');" style="cursor:hand">
<% end if %> 
    <td height="20" ><%= ocsaslist.FItemList(i).fmasteridx %></td>
    <td align="left"><acronym title="<%= ocsaslist.FItemList(i).shopGetAsDivCDName %>"><font color="<%= ocsaslist.FItemList(i).shopGetAsDivCDColor %>"><%= ocsaslist.FItemList(i).shopGetAsDivCDName %></font></acronym></td>
    <td><a href="javascript:reSearchByorderno('<%= ocsaslist.FItemList(i).forderno %>');" ><%= ocsaslist.FItemList(i).forderno %></a></td>    
    <td align="left">
        <acronym title="<%= ocsaslist.FItemList(i).Fmakerid %>"><a href="javascript:reSearchByMakerid('<%= ocsaslist.FItemList(i).Fmakerid %>');" ><%= Left(ocsaslist.FItemList(i).Fmakerid,32) %></a></acronym>
	</td>
    <td><%= ocsaslist.FItemList(i).Fcustomername %></td>    
    <td align="left"><acronym title="<%= ocsaslist.FItemList(i).Ftitle %>"><%= ocsaslist.FItemList(i).Ftitle %></acronym></td>
    <td><font color="<%= ocsaslist.FItemList(i).shopGetCurrstateColor %>"><%= ocsaslist.FItemList(i).shopGetCurrstateName %></font></td>    
    <td><acronym title="<%= ocsaslist.FItemList(i).Fregdate %>"><%= Left(ocsaslist.FItemList(i).Fregdate,10) %></acronym></td>
    <td><acronym title="<%= ocsaslist.FItemList(i).Ffinishdate %>"><%= Left(ocsaslist.FItemList(i).Ffinishdate,10) %></acronym></td>
    <td>
	    <% if ocsaslist.FItemList(i).Fdeleteyn="Y" then %>
	    	<font color="red">삭제</font>
	    <% elseif ocsaslist.FItemList(i).Fdeleteyn="C" then %>
	    	<font color="red"><strong>취소</strong></font>
	    <% end if %>
    </td>
</tr>
<% next %>

<tr bgcolor="#FFFFFF" >
    <td colspan="13" align="center">
        <% if ocsaslist.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ocsaslist.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for ix=0 + ocsaslist.StartScrollPage to ocsaslist.FScrollCount + ocsaslist.StartScrollPage - 1 %>
			<% if ix>ocsaslist.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(ix) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if ocsaslist.HasNextScroll then %>
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
<form name="buffrm" method="get" target="detailFrame" action="/admin/offshop/shopcscenter/action/cs_action_detail.asp" >
	<input type="hidden" name="masteridx" value="">
</form>
</table>

<script language='javascript'>

<% if ResultOneCsID<>"" then %>
    if (top.detailFrame!=undefined){
        top.detailFrame.location.href = "cs_action_detail.asp?id=<%= ResultOneCsID %>";
    }
<% end if %>
    
</script>

<%
set ocsaslist = Nothing
%>
<!-- #include virtual="/admin/offshop/shopcscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->