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
dim i, username, masteridx, makerid, searchfield, searchstring ,ix
dim searchtype, divcd, currstate ,yyyy1,yyyy2,mm1,mm2,dd1,dd2, yyyymmdd1
Dim writeUser ,delYN ,fromDate, toDate ,page ,ocsaslist ,ResultOneCsID
	delYN	= requestCheckVar(request("delYN"),10)
	username = requestCheckVar(request("username"),32)
	masteridx = requestCheckVar(request("masteridx"),10)
	searchfield = requestCheckVar(request("searchfield"),32)
	searchstring = requestCheckVar(request("searchstring"),32)
	searchtype = requestCheckVar(request("searchtype"),32)
	divcd = requestCheckVar(request("divcd"),4)
	currstate = requestCheckVar(request("currstate"),4)
	yyyy1   = requestCheckVar(request("yyyy1"),4)
	yyyy2   = requestCheckVar(request("yyyy2"),4)
	mm1     = requestCheckVar(request("mm1"),2)
	mm2     = requestCheckVar(request("mm2"),2)
	dd1     = requestCheckVar(request("dd1"),2)
	dd2     = requestCheckVar(request("dd2"),2)
	page = requestCheckVar(request("page"),10)

if page="" then page=1		
if searchtype="searchfield" and searchfield="" then searchstring="" end if
if searchtype="" then searchtype="searchfield"
	
if (yyyy1="") then 
    yyyymmdd1 = dateAdd("m",-2,now())
    yyyy1 = Cstr(Year(yyyymmdd1))
    mm1 = Cstr(Month(yyyymmdd1))
    dd1 = Cstr(day(yyyymmdd1))
end if

if (yyyy2="")   then yyyy2 = Cstr(Year(now()))
if (mm2="")     then mm2 = Cstr(Month(now()))
if (dd2="")     then dd2 = Cstr(day(now()))

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

'masteridx 에 데이타가 파라미터가 왔을때는 해당 파라미터로 세팅하고
'없을경우는 searchstring 에 데이타가 있는지를 확인하여 세팅한다.
'다른 페이지에서 링크를 걸어 팝업을 열었을때에 대한 처리.

if (masteridx <> "") then
    searchtype = "searchfield"
    username = ""
    searchfield = "masteridx"
    searchstring = masteridx
    divcd = ""
    currstate = ""
else
    if (searchstring <> "") then
        if (searchfield = "masteridx") then
                username = ""
                masteridx = searchstring
                makerid = ""
        elseif (searchfield = "makerid") then
                username = ""
                masteridx = ""
                makerid = searchstring

		elseif (searchfield = "writeUser") then
                writeUser = searchstring
		else
                username = searchstring
                masteridx = ""
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
	    ocsaslist.FRectMakerid  = makerid	
	    ocsaslist.FRectDivcd = divcd
	    ocsaslist.FRectCurrstate = currstate	
	    ocsaslist.FRectWriteUser = writeUser	
	    ocsaslist.FRectDeleteYN	= delYN
	else
	    ocsaslist.FRectStartDate = fromDate
	    ocsaslist.FRectEndDate = toDate
	    ocsaslist.FRectSearchType = searchtype
	    
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
    frm.searchtype[0].checked =true;
    frm.searchfield[4].selected = true;
    frm.searchstring.value = iorderno;
    frm.divcd.value = "";
    frm.currstate.value = "";
    frm.page.value="1";
    frm.submit();
}

function reSearchByMakerid(imakerid){
    frm.searchtype[0].checked =true;
    frm.searchfield[3].selected = true;
    frm.searchstring.value = imakerid;
    frm.page.value="1";
    frm.divcd.value = "";
    frm.currstate.value = "";
    frm.submit();
}

function SetComp(comp) {
    if (comp.value=="searchfield"){
        document.frm.dummy.checked = false;
        frm.searchstring.style.background = "#FFFFFF";
        
        frm.searchstring.focus();
        frm.searchstring.select();
    }else{
        document.frm.dummy.checked = true;
        frm.searchstring.style.background = "#EEEEEE";
        
        comp.focus();
        comp.select();
    }        
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
        <input type="radio" name="searchtype" onClick="SetComp(this);" value="searchfield" <% if (searchtype = "searchfield") then %>checked<% end if %>>조건검색
        [
        1.조건:
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
        <input type="text" class="text" name="searchstring" value="<%= searchstring %>" size="14" onFocus="ChangeCheckbox('searchtype', 'searchfield'); this.style.background = '#FFFFFF'">
        &nbsp;
        2.구분:
        <select class="select" name="divcd">
        	<option value="">전체</option>
			<option value="A000" <% if (divcd = "A000") then response.write "selected" end if %>>맞교환출고</option>
			<option value="A001" <% if (divcd = "A001") then response.write "selected" end if %>>누락재발송</option>
			<option value="A002" <% if (divcd = "A002") then response.write "selected" end if %>>서비스발송</option>			
			<option value="A004" <% if (divcd = "A004") then response.write "selected" end if %>>반품접수(업체배송)</option>
			<option value="A006" <% if (divcd = "A006") then response.write "selected" end if %>>출고시유의사항</option>
			<option value="A900" <% if (divcd = "A900") then response.write "selected" end if %>>주문내역변경</option>
        </select>
        &nbsp;
        3.진행상태:
        <select class="select" name="currstate">
        	<option value="">전체</option>
			<option value="B001" <% if (currstate = "B001") then response.write "selected" end if %>>접수</option>
			<option value="notfinish" <% if (currstate = "notfinish") then response.write "selected" end if %>>미처리전체</option> <!-- 6단계이하 -->
			<option value="B003" <% if (currstate = "B003") then response.write "selected" end if %>>택배사전송</option>
			<option value="B004" <% if (currstate = "B004") then response.write "selected" end if %>>운송장입력</option>
			<option value="B005" <% if (currstate = "B005") then response.write "selected" end if %>>확인요청</option>
			<option value="B006" <% if (currstate = "B006") then response.write "selected" end if %>>업체처리완료</option>
			<option value="B007" <% if (currstate = "B007") then response.write "selected" end if %>>완료</option>
        </select>
        <Br><input type="checkbox" name="dummy" value="" disabled <% if (searchfield="") then %>checked<% end if %>>
        <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
        ]
        <input type="checkbox" name="delYN" value="N" <%if (delYN="N") then %>checked<% end if %>>삭제(취소)제외     
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="reSearch();">
		<Br><input type="button" class="button_s" value="새로고침" onclick="document.location.reload();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
        <input type="radio" name="searchtype" onClick="SetComp(this)" value="notfinish" <% if (searchtype = "notfinish") then %>checked<% end if %>> 미처리전체                             
        <input type="radio" name="searchtype" onClick="SetComp(this)" value="beasongnocheck" <% if (searchtype = "beasongnocheck") then %>checked<% end if %>> 출고시유의사항
        <input type="radio" name="searchtype" onClick="SetComp(this)" value="upchemifinish" <% if (searchtype = "upchemifinish") then %>checked<% end if %>> 업체미처리
        <input type="radio" name="searchtype" onClick="SetComp(this)" value="upchefinish" <% if (searchtype = "upchefinish") then %>checked<% end if %>> 업체처리완료
        <input type="radio" name="searchtype" onClick="SetComp(this)" value="returnmifinish" <% if (searchtype = "returnmifinish") then %>checked<% end if %>> 회수요청미처리
        <input type="radio" name="searchtype" onClick="SetComp(this)" value="confirm" <% if (searchtype = "confirm") then %>checked<% end if %>> 확인요청        	
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
    <td>고객명</td>    
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
<form name="buffrm" method="get" target="detailFrame" action="/admin/offshop/cscenter/action/cs_action_detail.asp" >
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
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->