<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : cs센터 cs처리리스트
' History	:  2007.06.01 이상구 생성
'              2022.08.16 한용민 수정(isms보안조치)
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<%
'' 물류 팀장님 요청으로 물류팀원 모두 조회가능하게 수정, skyer9, 2017-03-09
''if (session("ssAdminPsn") = 9) then
''	'// 물류
''	if (session("ssBctId") <> "josin222") and (session("ssBctId") <> "jjh") and (session("ssBctId") <> "sunna0822") then
''		response.write "<br><br>권한이 없습니다."
''		response.end
''	end if
''end if

Dim delYN		: delYN	 = requestCheckvar(request("delYN"),1)
Dim periodYN	: periodYN = requestCheckvar(request("periodYN"),1)
Dim notfinishYN	: notfinishYN = requestCheckvar(request("notfinishYN"),1)
Dim research	: research = requestCheckvar(request("research"),2)
dim i, userid, username, orderserial, makerid, searchfield, searchstring, asid, writeUser, extsitename, checkExtSite, finishuser
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, yyyymmdd1, fromDate, toDate, notfinishtype, divcd, currstate
Dim onlycustomerjupsu, onlycsservicerefund, searchtype, dateType, upreturnmifinishBaseDate, tmpSql, page
dim ResultOneCsID, ix
	userid      	= requestCheckvar(request("userid"),32)
	username    	= requestCheckvar(request("username"),32)
	orderserial 	= requestCheckvar(request("orderserial"),32)
	asid 			= requestCheckvar(request("asid"),32)
	searchfield 	= requestCheckvar(request("searchfield"),32)
	searchstring 	= requestCheckvar(request("searchstring"),32)
	notfinishtype  	= requestCheckvar(request("notfinishtype"),32)
	divcd       	= requestCheckvar(request("divcd"),32)
	currstate   	= requestCheckvar(request("currstate"),32)
	extsitename 	= requestCheckvar(request("extsitename"),32)
	checkExtSite	= requestCheckvar(request("checkExtSite"),32)
	onlycustomerjupsu	= requestCheckvar(request("onlycustomerjupsu"),32)
	onlycsservicerefund	= requestCheckvar(request("onlycsservicerefund"),32)
	searchtype		= requestCheckvar(request("searchtype"),32)			'// 호환성을 위해 남겨둔다.(예를들면 [CS]고객센터>>[CS]메인 에서 오는 경우)
	dateType		= requestCheckvar(request("dateType"),32)
	yyyy1   = requestcheckvar(request("yyyy1"),4)
	yyyy2   = requestcheckvar(request("yyyy2"),4)
	mm1     = requestcheckvar(request("mm1"),2)
	mm2     = requestcheckvar(request("mm2"),2)
	dd1     = requestcheckvar(request("dd1"),2)
	dd2     = requestcheckvar(request("dd2"),2)
	page = requestcheckvar(getNumeric(request("page")),10)

if page="" then page=1

if (searchtype <> "") then
	if (searchtype = "searchfield") then
		'
	else
		notfinishYN = "Y"
		notfinishtype = searchtype
	end if
end if

if (research = "") then

	delYN = "N"

	if (searchtype <> "upchefinish") then
		periodYN = "Y"
	end if

	'// userid/orderserial 파라미터가 왔을때는 해당 파라미터로 세팅
	'// (다른 페이지에서 링크를 걸어 팝업을 열었을때에 대한 처리.)
	if (userid <> "") then
	    searchfield = "userid"
	    searchstring = userid
	elseif (orderserial <> "") then
	    searchfield = "orderserial"
	    searchstring = orderserial
	end if

    if (notfinishtype = "confirm") then
        divcd = "A003"
        currstate = "B005"
    elseif (notfinishtype = "cardnocheckdp1") then
        divcd = "A007"
        currstate = "notfinish"
    elseif (notfinishtype = "norefund") then
        divcd = "A003"
        currstate = "B001"
    end if
end if

if (searchfield <> "") and (searchstring <> "") then

    if (searchfield = "userid") then

            userid = searchstring

    elseif (searchfield = "orderserial") then

            orderserial = searchstring

    elseif (searchfield = "username") then

            username = searchstring

    elseif (searchfield = "makerid") then

            makerid = searchstring

	elseif (searchfield = "writeUser") then

            writeUser = searchstring

	elseif (searchfield = "finishuser") then

            finishuser = searchstring

	elseif (searchfield = "asid") then

			asid = searchstring

    end If

end if

if (searchfield = "") and (searchstring <> "") then

	if IsNumeric(searchstring) and Len(searchstring) >= 11 then
		'// 주문번호 검색
		searchfield = "orderserial"
		orderserial = searchstring
	end if

end if

if (yyyy1="") then
    yyyymmdd1 = dateAdd("m",-3,now())			'// [CS]고객센터>>[CS]메인 참조
    yyyy1 = Cstr(Year(yyyymmdd1))
    mm1 = Cstr(Month(yyyymmdd1))
    dd1 = Cstr(day(yyyymmdd1))
end if

if (yyyy2 = "") then
	if (notfinishtype = "upreturnmifinish") then
		'// 업체반품미처리의 경우 기본값 = D+7 일
		tmpSql = " exec [db_cs].[dbo].[usp_getDayMinusWorkday] '" & Left(now(), 10) & "', 7 " & VbCRLF
		rsget.CursorLocation = adUseClient
		rsget.Open tmpSql, dbget, adOpenForwardOnly
		if Not rsget.Eof then
		    '// 근무일수 기준 D+7 일
		    upreturnmifinishBaseDate = rsget("minusworkday")

		    yyyy2 = Cstr(Year(upreturnmifinishBaseDate))
		    mm2 = Cstr(Month(upreturnmifinishBaseDate))
		    dd2 = Cstr(day(upreturnmifinishBaseDate))
		end if
		rsget.close
	end if

    if (notfinishtype = "cardnocheckdp1") then
        toDate = DateAdd("d", -1, Now())
        yyyy2 = Cstr(Year(toDate))
        mm2 = Cstr(Month(toDate))
        dd2 = Cstr(day(toDate))
        notfinishtype = "cardnocheck"
    end if

	if (yyyy2="")   then yyyy2 = Cstr(Year(now()))
	if (mm2="")     then mm2 = Cstr(Month(now()))
	if (dd2="")     then dd2 = Cstr(day(now()))
end if

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FPageSize = 10
ocsaslist.FCurrPage = page

if (searchfield <> "") and (searchstring <> "") then
    ocsaslist.FRectUserID = userid
    ocsaslist.FRectUserName = username
    ocsaslist.FRectOrderSerial = orderserial
    ocsaslist.FRectMakerid  = makerid
    ocsaslist.FRectWriteUser = writeUser
	ocsaslist.FRectFinishUser = finishuser
	ocsaslist.FRectCsAsID = asid
end if

ocsaslist.FRectDivcd = divcd
ocsaslist.FRectCurrstate = currstate

if (orderserial = "") and (userid = "") then
	'// 주문번호 또는 아이디 검색하면 삭제내역 포함 표시
	ocsaslist.FRectDeleteYN	= delYN
end if

if (notfinishYN = "Y") then
	ocsaslist.FRectSearchType = notfinishtype
end if

If (periodYN = "Y") and (orderserial = "") Then
	'// 주문번호 입력하면 기간제한 없음
	ocsaslist.FRectDateType = dateType
	ocsaslist.FRectStartDate = fromDate
	ocsaslist.FRectEndDate = toDate
End If

IF (checkExtSite<>"") then                      '''2011-06 추가
    ocsaslist.FRectExtSitename = ExtSitename
ENd IF

ocsaslist.FRectOnlyCustomerJupsu = onlycustomerjupsu
ocsaslist.FRectOnlyCSServiceRefund = onlycsservicerefund
''ocsaslist.GetCSASMasterListNew
ocsaslist.GetCSASMasterListByProcedure

if ocsaslist.FResultCount=1 then
    ResultOneCsID = ocsaslist.FItemList(0).FId
end if

%>
<link rel="stylesheet" href="/cscenter/css/cs.css" type="text/css">
<style>
.csH15 { line-height: 15px; }
</style>
<script type='text/javascript'>
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
	frm.target = "";
	frm.action = "cs_action_list.asp"
    frm.page.value = page;
    frm.submit();
}

function reSearch(){
	if (frm.searchfield.value=="asid"){
		if (frm.searchstring.value!=""){
			if (!IsDouble(frm.searchstring.value)){
				alert('cs번호는 숫자만 가능합니다.');
				frm.searchstring.focus();
				return;
			}
		}
	}

	frm.target = "";
	frm.action = "cs_action_list.asp"
    frm.page.value="1";
    frm.submit();
}

<% ' 기능은 있으나, isms 결함조치로 막음 %>
//function reSearchExcelDown(){
//	frm.target = "exceldown";
<% '	frm.action = "/cscenter/action/cs_action_list_excel.asp" %>
//    frm.submit();
//	frm.target = "";
//	frm.action = ""
//}

function reSearchByOrderserial(iorderserial){
    frm.searchfield.value = "orderserial";
    frm.searchstring.value = iorderserial;

    frm.divcd.value = "";
    frm.currstate.value = "";

	// frm.notfinishYN.checked = false;
	// frm.periodYN.checked = false;
	// frm.checkExtSite.checked = false;
	// frm.delYN.checked = false;

    frm.page.value="1";
	frm.target = "";
	frm.action = "cs_action_list.asp"
    frm.submit();
}

function reSearchByUserid(iuserid){
    frm.searchfield.value = "userid";
    frm.searchstring.value = iuserid;

    frm.divcd.value = "";
    frm.currstate.value = "";

	// frm.notfinishYN.checked = false;
	// frm.periodYN.checked = false;
	// frm.checkExtSite.checked = false;
	// frm.delYN.checked = false;

    frm.page.value="1";
	frm.target = "";
	frm.action = "cs_action_list.asp"
    frm.submit();
}

function reSearchByMakerid(imakerid){
    frm.searchfield.value = "makerid";
    frm.searchstring.value = imakerid;

    frm.divcd.value = "";
    frm.currstate.value = "";

	// frm.notfinishYN.checked = false;
	// frm.periodYN.checked = false;
	// frm.checkExtSite.checked = false;
	// frm.delYN.checked = false;

    frm.page.value="1";
	frm.target = "";
	frm.action = "cs_action_list.asp"
    frm.submit();
}

function SetComp(comp) {
	frm.notfinishYN.checked = true;
}

function SetExtCheck(comp) {
    if (comp.name=="checkExtSite"){
        if (comp.checked){
            frm.extsitename.style.background = "#FFFFFF";
        }else{
            frm.extsitename.style.background = "#EEEEEE";
        }
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

<form name="frm" method="get" action="/cscenter/action/cs_action_list.asp" style="margin:0px;" >
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="F4F4F4">
	<tr>
    	<td>
			&nbsp;
            검색 :
            <select class="select" name="searchfield">
            	<option value="" <% if (searchfield = "") then %>selected<% end if %>>전체</option>
				<option value="orderserial" <% if (searchfield = "orderserial") then %>selected<% end if %>>주문번호</option>
				<option value="username" <% if (searchfield = "username") then %>selected<% end if %>>고객명</option>
				<option value="userid" <% if (searchfield = "userid") then %>selected<% end if %>>아이디</option>
				<option value="makerid" <% if (searchfield = "makerid") then %>selected<% end if %>>업체처리아이디</option>
				<option value="writeUser" <% if (searchfield = "writeUser") then %>selected<% end if %>>접수자아이디</option>
				<option value="finishuser" <% if (searchfield = "finishuser") then %>selected<% end if %>>처리자아이디</option>
				<option value="asid" <% if (searchfield = "asid") then %>selected<% end if %>>CSidx</option>
            </select>
            <input type="text" class="text" name="searchstring" value="<%= searchstring %>" size="14">
            &nbsp;&nbsp;
            구분:
            <select class="select" name="divcd">
            	<option value="">전체</option>
            	<option value="">-------------------------</option>
				<option value="A000" <% if (divcd = "A000") then response.write "selected" end if %>>교환출고</option>
				<option value="A100" <% if (divcd = "A100") then response.write "selected" end if %>>교환출고(상품변경)</option>
				<option value="">-------------------------</option>
				<option value="A004" <% if (divcd = "A004") then response.write "selected" end if %>>반품접수(업배)</option>
				<option value="">-------------------------</option>
				<option value="A001" <% if (divcd = "A001") then response.write "selected" end if %>>누락재발송</option>
				<option value="A002" <% if (divcd = "A002") then response.write "selected" end if %>>서비스발송</option>
				<option value="A200" <% if (divcd = "A200") then response.write "selected" end if %>>기타회수</option>
				<option value="">-------------------------</option>
				<option value="A003" <% if (divcd = "A003") then response.write "selected" end if %>>환불요청</option>
				<option value="A005" <% if (divcd = "A005") then response.write "selected" end if %>>외부몰환불요청</option>
				<option value="A007" <% if (divcd = "A007") then response.write "selected" end if %>>신용카드/이체취소요청</option>
				<option value="A700" <% if (divcd = "A700") then response.write "selected" end if %>>업체기타정산</option>
				<option value="A999" <% if (divcd = "A999") then response.write "selected" end if %>>고객추가결제</option>
				<option value="">-------------------------</option>
				<option value="A060" <% if (divcd = "A060") then response.write "selected" end if %>>업체긴급문의</option>
				<option value="A006" <% if (divcd = "A006") then response.write "selected" end if %>>출고시유의사항</option>
				<option value="A008" <% if (divcd = "A008") then response.write "selected" end if %>>주문취소</option>
				<option value="A009" <% if (divcd = "A009") then response.write "selected" end if %>>기타내역(메모)</option>
				<option value="A900" <% if (divcd = "A900") then response.write "selected" end if %>>주문내역변경</option>
				<option value="">-------------------------</option>
				<option value="A010" <% if (divcd = "A010") then response.write "selected" end if %>>회수신청(텐배)</option>
				<option value="">-------------------------</option>
				<option value="A011" <% if (divcd = "A011") then response.write "selected" end if %>>교환회수(텐배)</option>
				<option value="A012" <% if (divcd = "A012") then response.write "selected" end if %>>교환회수(업배)</option>
				<option value="A111" <% if (divcd = "A111") then response.write "selected" end if %>>교환회수(상품변경,텐배)</option>
				<option value="A112" <% if (divcd = "A112") then response.write "selected" end if %>>교환회수(상품변경,업배)</option>
            </select>
            &nbsp;&nbsp;
            상태:
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
            &nbsp;&nbsp;
			<input type="checkbox" name="delYN" value="N" <%if (delYN="N") then %>checked<% end if %>>삭제(취소)제외
        </td>
        <td width="100" align="right" valign="top" rowspan="3">
			<% '<input type="button" class="button_s" value="엑셀다운로드" onclick="reSearchExcelDown();"><br /> %>
            <input type="button" class="button_s" value="새로고침" onclick="document.location.reload();"><br />
			<div style="height: 5px;"></div>
            <input type="button" class="button_s" value="검색하기" onclick="reSearch();">
        </td>
	</tr>
	<tr>
    	<td>
    		&nbsp;
    		<input type="checkbox" name="notfinishYN" value="Y" <%=CHKIIF(notfinishYN="Y","checked","")%>>
    		미처리CS :
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="notfinish" <% if (notfinishtype = "notfinish") then %>checked<% end if %>> 전체
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="norefundmile" <% if (notfinishtype = "norefundmile") then %>checked<% end if %>> 마일리지/예치금 환불
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="cardnocheck" <% if (notfinishtype = "cardnocheck") then %>checked<% end if %>> 카드취소
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="cancelnofinish" <% if (notfinishtype = "cancelnofinish") then %>checked<% end if %>> 주문취소
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="beasongnocheck" <% if (notfinishtype = "beasongnocheck") then %>checked<% end if %>> 출고시유의사항
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="upchemifinish" <% if (notfinishtype = "upchemifinish") then %>checked<% end if %>> 업체미처리전체
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="upreturnmifinish" <% if (notfinishtype = "upreturnmifinish") then %>checked<% end if %>> 업체반품
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="upchefinish" <% if (notfinishtype = "upchefinish") then %>checked<% end if %>> 업체처리완료
			<input type="radio" name="notfinishtype" onClick="SetComp(this)" value="logicsfinish" <% if (notfinishtype = "logicsfinish") then %>checked<% end if %>> 물류처리완료
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="chulgofinishnotreceive" <% if (notfinishtype = "chulgofinishnotreceive") then %>checked<% end if %>> 교환출고후미회수
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="returnmifinish" <% if (notfinishtype = "returnmifinish") then %>checked<% end if %>> 회수요청
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="confirm" <% if (notfinishtype = "confirm") then %>checked<% end if %>> 확인요청(환불)
			<input type="radio" name="notfinishtype" onClick="SetComp(this)" value="upcheconfirm" <% if (notfinishtype = "upcheconfirm") then %>checked<% end if %>> 확인요청(업체)
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="norefundetc" <% if (notfinishtype = "norefundetc") then %>checked<% end if %>> 외부몰환불
			<input type="radio" name="notfinishtype" onClick="SetComp(this)" value="customeraddpay" <% if (notfinishtype = "customeraddpay") then %>checked<% end if %>> 고객추가입금
        </td>
	</tr>
	<tr>
    	<td>
    		&nbsp;
            (Total : <%= ocsaslist.FTotalCount%> 건)
            &nbsp;
            <input type="checkbox" name="periodYN" value="Y" <%=CHKIIF(periodYN="Y","checked","")%>>
			<select class="select" name="dateType">
				<option value="regdate" <%= CHKIIF(dateType="regdate", "selected", "") %> >접수일</option>
				<option value="finishdate" <%= CHKIIF(dateType="finishdate", "selected", "") %> >처리일</option>
			</select>
             :
            <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;&nbsp;
            <input type="checkbox" name="checkExtSite" value="Y" <% if checkExtSite="Y" then response.write "checked" %> onClick="SetExtCheck(this)">
            특정사이트 : <% DrawSelectExtSiteName "extsitename", extsitename %>
			&nbsp;&nbsp;
			<input type="checkbox" name="onlycustomerjupsu" value="Y" <%if (onlycustomerjupsu="Y") then %>checked<% end if %>>고객 직접접수만
			&nbsp;&nbsp;
			<input type="checkbox" name="onlycsservicerefund" value="Y" <%if (onlycsservicerefund="Y") then %>checked<% end if %>>CS서비스 환불만
        </td>
	</tr>
</table>
</form>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a csH15" cellpadding="2" cellspacing="1" bgcolor="#BABABA" style="table-layout:fixed">
    <tr bgcolor="<%= adminColor("tabletop") %>">
        <td width="70" align="center">Idx</td>
        <td width="100" align="center">구분</td>
        <td width="90" align="center">원주문번호</td>
        <td width="90" align="center">Site</td>
        <td width="110" align="center">업체ID</td>
        <td width="50" align="center">고객명</td>
        <td width="80" align="center">아이디</td>
        <td align="center">제목</td>
        <td width="75" align="center">상태</td>
		<td width="75" align="center">접수자</td>
		<td width="75" align="center">처리자</td>
        <td width="70" align="center">환불금액</td>
        <td width="80" align="center">등록일</td>
        <td width="80" align="center">업체확인</td>
        <td width="80" align="center">처리일</td>
        <td width="30" align="center">삭제</td>
    </tr>

<% for i = 0 to (ocsaslist.FResultCount - 1) %>
    <% if (ocsaslist.FItemList(i).Fdeleteyn <> "N") then %>
    <tr bgcolor="#EEEEEE" class="csH15 csMp" style="color:gray" align="center" onclick="ChangeColor(this,'AFEEEE','FFFFFF'); searchDetail('<%= ocsaslist.FItemList(i).Fid %>');">
    <% else %>
	<tr bgcolor="#FFFFFF" class="csH15 csMp" align="center" onclick="ChangeColor(this,'AFEEEE','FFFFFF'); searchDetail('<%= ocsaslist.FItemList(i).Fid %>');">
    <% end if %>
        <td class="csNoWrap"><%= ocsaslist.FItemList(i).Fid %></td>
        <td class="csNoWrap" align="left"><font color="<%= ocsaslist.FItemList(i).GetAsDivCDColor %>"><%= ocsaslist.FItemList(i).GetAsDivCDName %></font></td>
        <td class="csNoWrap">
        	<a href="javascript:reSearchByOrderserial('<%= ocsaslist.FItemList(i).Forgorderserial %>');" >
        		<%= ocsaslist.FItemList(i).Forgorderserial %>
        		<% if (ocsaslist.FItemList(i).Forderserial <> ocsaslist.FItemList(i).Forgorderserial) then %>
        			+
        		<% end if %>
        	</a>
        </td>
        <td class="csNoWrap"><%= ocsaslist.FItemList(i).FExtsitename %></td>
        <td class="csNoWrap" align="left">
            <acronym title="<%= ocsaslist.FItemList(i).Fmakerid %>"><a href="javascript:reSearchByMakerid('<%= ocsaslist.FItemList(i).Fmakerid %>');" ><%= Left(ocsaslist.FItemList(i).Fmakerid,32) %></a></acronym>
		</td>
        <td class="csNoWrap">
			<%= AstarUserName(ocsaslist.FItemList(i).Fcustomername) %>
        </td>
        <td class="csNoWrap" align="left">
        	<!--<acronym title="<%'= ocsaslist.FItemList(i).Fuserid %>">-->
        	<!--<a href="javascript:reSearchByUserid('<%'= ocsaslist.FItemList(i).Fuserid %>');" >-->
			<%= AstarUserid(ocsaslist.FItemList(i).Fuserid) %>
        	<!--</a>-->
        	<!--</acronym>-->
        </td>
        <td class="csNoWrap" align="left">
			<%= ocsaslist.FItemList(i).Ftitle %>
			<% if ocsaslist.FItemList(i).FExtsitename<>"10x10" then %>(<%= ocsaslist.FItemList(i).FAuthCode %>)<% end if %>
		</td>
        <td class="csNoWrap"><font color="<%= ocsaslist.FItemList(i).GetCurrstateColor %>"><%= ocsaslist.FItemList(i).GetCurrstateName %></font></td>
		<td class="csNoWrap"><%= ocsaslist.FItemList(i).Fwriteuser %></td>
		<td class="csNoWrap"><%= ocsaslist.FItemList(i).Ffinishuser %></td>
        <td class="csNoWrap" align="right"><%= FormatNumber(ocsaslist.FItemList(i).Frefundrequire,0) %></td>
        <td class="csNoWrap"><acronym title="<%= ocsaslist.FItemList(i).Fregdate %>"><%= Left(ocsaslist.FItemList(i).Fregdate,10) %></acronym></td>
		<td class="csNoWrap"><acronym title="<%= ocsaslist.FItemList(i).Fconfirmdate %>"><%= Left(ocsaslist.FItemList(i).Fconfirmdate,10) %></acronym></td>
        <td class="csNoWrap"><acronym title="<%= ocsaslist.FItemList(i).Ffinishdate %>"><%= Left(ocsaslist.FItemList(i).Ffinishdate,10) %></acronym></td>
        <td class="csNoWrap">
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
    <tr bgcolor="#FFFFFF" class="csH15 csMp" align="center">
        <td class="csNoWrap">-</td>
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
        <td></td>
        <td></td>
        <td></td>
    </tr>
        <% next %>
<% end if %>
    <tr bgcolor="#FFFFFF" >
        <td colspan="16" align="center">
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

<form name="buffrm" method="get" target="detailFrame" action="/cscenter/action/cs_action_detail.asp" style="margin:0px;" >
<input type="hidden" name="id" value="">
</form>

<script type='text/javascript'>

    <% if ResultOneCsID<>"" then %>
    if (top.detailFrame!=undefined){
        top.detailFrame.location.href = "cs_action_detail.asp?id=<%= ResultOneCsID %>";
    }
    <% end if %>

</script>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe src="about:blank" name="exceldown" border="0" width="100%" height="300"></iframe>
<% else %>
	<iframe src="about:blank" name="exceldown" border="0" width="100%" height="0"></iframe>
<% end if %>

<%
set ocsaslist = Nothing
%>
<script type='text/javascript'>

function getOnload(){
SetExtCheck(frm.checkExtSite);
}

window.onload=getOnload;

</script>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
