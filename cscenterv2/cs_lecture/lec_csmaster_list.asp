<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 핑거스 고객센터 강좌CS처리 리스트
' Hieditor : 2015.05.27 이상구 생성
'			 2017.07.07 한용민 수정
'###########################################################
%>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog_ACA.asp" -->
<!-- #include virtual="/cscenterv2/lib/classes/cs_lecture/lec_cs_aslistcls.asp"-->
<!--
<11111111111!-111- #inc111lude virtual="/lib/util/datelib.asp" -111-11111111111>
-->
<%
Dim delYN	: delYN	= req("delYN","")


dim i, userid, username, orderserial, makerid, searchfield, searchstring
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, yyyymmdd1
dim fromDate, toDate
dim searchtype, divcd, currstate
Dim writeUser

userid = RequestCheckvar(request("userid"),32)
username = RequestCheckvar(request("username"),32)
orderserial = RequestCheckvar(request("orderserial"),16)
searchfield = RequestCheckvar(request("searchfield"),16)
searchstring = RequestCheckvar(request("searchstring"),32)
searchtype = RequestCheckvar(request("searchtype"),16)
divcd = RequestCheckvar(request("divcd"),4)
currstate = RequestCheckvar(request("currstate"),4)

if searchtype="searchfield" and searchfield="" then searchstring="" end if

if searchtype="" then searchtype="searchfield"

'userid/orderserial 에 데이타가 파라미터가 왔을때는 해당 파라미터로 세팅하고
'없을경우는 searchstring 에 데이타가 있는지를 확인하여 세팅한다.
'다른 페이지에서 링크를 걸어 팝업을 열었을때에 대한 처리.
if (userid <> "") then
    searchtype = "searchfield"
    username = ""
    orderserial = ""
    searchfield = "userid"
    searchstring = userid
    divcd = ""
    currstate = ""

elseif (orderserial <> "") then
    searchtype = "searchfield"
    username = ""
    userid = ""
    searchfield = "orderserial"
    searchstring = orderserial
    divcd = ""
    currstate = ""
else
    if (searchstring <> "") then
        if (searchfield = "userid") then
                userid = searchstring
                username = ""
                orderserial = ""
                makerid = ""
        elseif (searchfield = "orderserial") then
                userid = ""
                username = ""
                orderserial = searchstring
                makerid = ""
        elseif (searchfield = "makerid") then
                userid = ""
                username = ""
                orderserial = ""
                makerid = searchstring

		elseif (searchfield = "writeUser") then
                writeUser = searchstring
		else
                userid = ""
                username = searchstring
                orderserial = ""
                makerid = ""
        end If


    else
        userid = ""
        username = ""
        orderserial = ""
        searchfield = ""
        searchstring = ""
    end if
end if


yyyy1   = RequestCheckvar(request("yyyy1"),4)
yyyy2   = RequestCheckvar(request("yyyy2"),4)
mm1     = RequestCheckvar(request("mm1"),2)
mm2     = RequestCheckvar(request("mm2"),2)
dd1     = RequestCheckvar(request("dd1"),2)
dd2     = RequestCheckvar(request("dd2"),2)

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




'==============================================================================

dim page
page = RequestCheckvar(request("page"),10)
if page="" then page=1

dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FPageSize = 10
ocsaslist.FCurrPage = page
if (searchtype = "searchfield") then
	ocsaslist.FRectSearchType = searchtype
    ocsaslist.FRectUserID = userid
    ocsaslist.FRectUserName = username
    ocsaslist.FRectOrderSerial = orderserial
    ocsaslist.FRectMakerid  = makerid

    ocsaslist.FRectDivcd = divcd
    ocsaslist.FRectCurrstate = currstate

    ocsaslist.FRectWriteUser = writeUser

    ocsaslist.FRectDeleteYN	= delYN

'    ocsaslist.FRectStartDate = fromDate
'    ocsaslist.FRectEndDate = toDate
else
    ocsaslist.FRectStartDate = fromDate
    ocsaslist.FRectEndDate = toDate
    ocsaslist.FRectSearchType = searchtype

end if

ocsaslist.GetCSASMasterList


dim ResultOneCsID
if ocsaslist.FResultCount=1 then
    ResultOneCsID = ocsaslist.FItemList(0).FId
end if

dim ix
%>

<script language='javascript'>
// tr 색상변경
var pre_selected_row = null;

function ChangeColor(e, selcolor, defcolor){
	if (pre_selected_row != null) {
	        pre_selected_row.bgColor = defcolor;
        }
        pre_selected_row = e;
        e.bgColor = selcolor;
}

function searchDetail(idx){
    buffrm.id.value = idx;
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

function reSearchByOrderserial(iorderserial){
    frm.searchtype[0].checked =true;
    frm.searchfield[1].selected = true;
    frm.searchstring.value = iorderserial;
    frm.divcd.value = "";
    frm.currstate.value = "";
    frm.page.value="1";
    frm.submit();
}

function reSearchByUserid(iuserid){
    frm.searchtype[0].checked =true;
    frm.searchfield[4].selected = true;
    frm.searchstring.value = iuserid;
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


function pop_modal_repay(id){
	if (id == "") {
	        alert("먼저 CS요청을 선택하세요.");
	        return;
        }
	var popwin = window.open("pop_modal_repay.asp?id=" + id,"pop_modal_repay","width=350 height=350 scrollbars=no resizable=no");
	popwin.focus();
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



<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="F4F4F4">
   	<form name="frm" method="get">
   	<input type="hidden" name="page" value="1">
   	<input type="hidden" name="id" value="">
	<tr height="50">
    	<td>

            <input type="radio" name="searchtype" onClick="SetComp(this);" value="searchfield" <% if (searchtype = "searchfield") then %>checked<% end if %>>조건검색
            [
            1.조건:
            <select class="select" name="searchfield">
            	<option value="" <% if (searchfield = "") then %>selected<% end if %>>전체</option>
				<option value="orderserial" <% if (searchfield = "orderserial") then %>selected<% end if %>>주문번호</option>
				<option value="username" <% if (searchfield = "username") then %>selected<% end if %>>고객명</option>
				<option value="makerid" <% if (searchfield = "makerid") then %>selected<% end if %>>업체처리아이디</option>
				<option value="userid" <% if (searchfield = "userid") then %>selected<% end if %>>아이디</option>
				<option value="orderserial" <% if (searchfield = "orderserial") then %>selected<% end if %>>주문번호</option>
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
				<option value="A003" <% if (divcd = "A003") then response.write "selected" end if %>>환불요청</option>
				<option value="A004" <% if (divcd = "A004") then response.write "selected" end if %>>반품접수(업체배송)</option>
				<option value="A005" <% if (divcd = "A005") then response.write "selected" end if %>>외부몰환불요청</option>
				<option value="A006" <% if (divcd = "A006") then response.write "selected" end if %>>출고시유의사항</option>
				<option value="A007" <% if (divcd = "A007") then response.write "selected" end if %>>신용카드/이체취소요청</option>
				<option value="A008" <% if (divcd = "A008") then response.write "selected" end if %>>주문취소</option>
				<option value="A009" <% if (divcd = "A009") then response.write "selected" end if %>>기타내역(메모)</option>
				<option value="A010" <% if (divcd = "A010") then response.write "selected" end if %>>회수신청(텐바이텐배송)</option>
				<option value="A011" <% if (divcd = "A011") then response.write "selected" end if %>>맞교환회수(텐바이텐배송)</option>
				<option value="A700" <% if (divcd = "A700") then response.write "selected" end if %>>업체기타정산</option>
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

            <input type="checkbox" name="dummy" value="" disabled <% if (searchfield="") then %>checked<% end if %>>
            <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
            ]
            <input type="checkbox" name="delYN" value="N" <%if (delYN="N") then %>checked<% end if %>>삭제(취소)제외
            <br>

            <input type="radio" name="searchtype" onClick="SetComp(this)" value="notfinish" <% if (searchtype = "notfinish") then %>checked<% end if %>> 미처리전체
            <input type="radio" name="searchtype" onClick="SetComp(this)" value="norefund" <% if (searchtype = "norefund") then %>checked<% end if %>> 환불미처리
            <input type="radio" name="searchtype" onClick="SetComp(this)" value="norefundmile" <% if (searchtype = "norefundmile") then %>checked<% end if %>> 마일리지환불미처리
            <input type="radio" name="searchtype" onClick="SetComp(this)" value="cardnocheck" <% if (searchtype = "cardnocheck") then %>checked<% end if %>> 카드취소미처리
            <input type="radio" name="searchtype" onClick="SetComp(this)" value="cancelnofinish" <% if (searchtype = "cancelnofinish") then %>checked<% end if %>> 주문취소미처리
            <input type="radio" name="searchtype" onClick="SetComp(this)" value="beasongnocheck" <% if (searchtype = "beasongnocheck") then %>checked<% end if %>> 출고시유의사항
            <input type="radio" name="searchtype" onClick="SetComp(this)" value="upchemifinish" <% if (searchtype = "upchemifinish") then %>checked<% end if %>> 업체미처리
            <input type="radio" name="searchtype" onClick="SetComp(this)" value="upchefinish" <% if (searchtype = "upchefinish") then %>checked<% end if %>> 업체처리완료
            <input type="radio" name="searchtype" onClick="SetComp(this)" value="returnmifinish" <% if (searchtype = "returnmifinish") then %>checked<% end if %>> 회수요청미처리
            <input type="radio" name="searchtype" onClick="SetComp(this)" value="confirm" <% if (searchtype = "confirm") then %>checked<% end if %>> 확인요청
            <input type="radio" name="searchtype" onClick="SetComp(this)" value="norefundetc" <% if (searchtype = "norefundetc") then %>checked<% end if %>> 외부몰환불미처리
            &nbsp;

        </td>
        <td width="80" align="right" valign="top">
            <input type="button" class="button_s" value="새로고침" onclick="document.location.reload();">
            &nbsp;
            <input type="button" class="button_s" value="검색하기" onclick="reSearch();">
        </td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" style="table-layout:fixed">
    <tr bgcolor="<%= adminColor("tabletop") %>">
        <td width="50" align="center">Idx</td>
        <td width="100" align="center">구분</td>
        <td width="90" align="center">관련주문번호</td>
        <td width="90" align="center">Site</td>
        <td width="110" align="center">업체ID</td>
        <td width="50" align="center">고객명</td>
        <td width="80" align="center">아이디</td>
        <td align="center">제목</td>
        <td width="75" align="center">상태</td>
        <td width="70" align="center">환불금액</td>
        <td width="80" align="center">등록일</td>
        <td width="80" align="center">처리일</td>
        <td width="30" align="center">삭제</td>
    </tr>

<% for i = 0 to (ocsaslist.FResultCount - 1) %>
    <% if (ocsaslist.FItemList(i).Fdeleteyn <> "N") then %>
    <tr bgcolor="#EEEEEE" style="color:gray" align="center" onclick="ChangeColor(this,'AFEEEE','FFFFFF'); searchDetail('<%= ocsaslist.FItemList(i).Fid %>');" style="cursor:hand">
    <% else %>
    <tr bgcolor="#FFFFFF" align="center" onclick="ChangeColor(this,'AFEEEE','FFFFFF'); searchDetail('<%= ocsaslist.FItemList(i).Fid %>');" style="cursor:hand">
    <% end if %>
        <td height="20" nowrap><%= ocsaslist.FItemList(i).Fid %></td>
        <td nowrap align="left"><acronym title="<%= ocsaslist.FItemList(i).GetAsDivCDName %>"><font color="<%= ocsaslist.FItemList(i).GetAsDivCDColor %>"><%= ocsaslist.FItemList(i).GetAsDivCDName %></font></acronym></td>
        <td nowrap><a href="javascript:reSearchByOrderserial('<%= ocsaslist.FItemList(i).Forderserial %>');" ><%= ocsaslist.FItemList(i).Forderserial %></a></td>
        <td nowrap><%= ocsaslist.FItemList(i).FExtsitename %></td>
        <td nowrap align="left">
            <% if ocsaslist.FItemList(i).FExtsitename<>"10x10" then %>
            <%''= ocsaslist.FItemList(i).FAuthCode %>
            <% end if %>
            <acronym title="<%= ocsaslist.FItemList(i).Fmakerid %>"><a href="javascript:reSearchByMakerid('<%= ocsaslist.FItemList(i).Fmakerid %>');" ><%= Left(ocsaslist.FItemList(i).Fmakerid,32) %></a></acronym></td>
        <td nowrap><%= ocsaslist.FItemList(i).Fcustomername %></td>
        <td nowrap align="left">
        	<!--<acronym title="<%'= ocsaslist.FItemList(i).Fuserid %>">-->
        	<!--<a href="javascript:reSearchByUserid('<%'= ocsaslist.FItemList(i).Fuserid %>');" >-->
        	<%= printUserId(ocsaslist.FItemList(i).Fuserid, 2, "*") %>
        	<!--</a>-->
        	<!--</acronym>-->
        </td>
        <td nowrap align="left">
        	<acronym title="<%= ocsaslist.FItemList(i).Ftitle %>"><%= ocsaslist.FItemList(i).Ftitle %></acronym></td>
        <td nowrap><font color="<%= ocsaslist.FItemList(i).GetCurrstateColor %>"><%= ocsaslist.FItemList(i).GetCurrstateName %></font></td>
        <td nowrap align="right"><%= FormatNumber(ocsaslist.FItemList(i).Frefundrequire,0) %></td>
        <td nowrap><acronym title="<%= ocsaslist.FItemList(i).Fregdate %>"><%= Left(ocsaslist.FItemList(i).Fregdate,10) %></acronym></td>
        <td nowrap><acronym title="<%= ocsaslist.FItemList(i).Ffinishdate %>"><%= Left(ocsaslist.FItemList(i).Ffinishdate,10) %></acronym></td>
        <td nowrap>
        <% if ocsaslist.FItemList(i).Fdeleteyn="Y" then %>
        <font color="red">삭제</font>
        <% elseif ocsaslist.FItemList(i).Fdeleteyn="C" then %>
        <font color="red"><strong>취소</strong></font>
        <% end if %>
        </td>
    </tr>
<% next %>
<% if (ocsaslist.FResultCount < 9) then %>
        <% for i = 0 to (9 - (ocsaslist.FResultCount mod 9)) %>
    <tr bgcolor="#FFFFFF" align="center">
        <td height="20"></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
    </tr>
        <% next %>
<% end if %>
    <tr bgcolor="#FFFFFF" >
        <td colspan="13" align="center">
            <% if ocsaslist.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ocsaslist.StarScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for ix=0 + ocsaslist.StarScrollPage to ocsaslist.FScrollCount + ocsaslist.StarScrollPage - 1 %>
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

</table>

<form name="buffrm" method="get" target="detailFrame" action="lec_csdetail_view.asp" >
<input type="hidden" name="id" value="">
</form>

<script language='javascript'>
    <% if ResultOneCsID<>"" then %>
    if (top.detailFrame!=undefined){
        top.detailFrame.location.href = "lec_csdetail_view.asp?id=<%= ResultOneCsID %>";
    }
    <% end if %>
</script>
<%

set ocsaslist = Nothing

%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->