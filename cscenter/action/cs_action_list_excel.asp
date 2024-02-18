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
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
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
dim ResultOneCsID, ix, arrlist, menupos
	menupos      	= requestCheckvar(request("menupos"),10)
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
ocsaslist.FPageSize = 100000
ocsaslist.FCurrPage = 1

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
ocsaslist.GetCSASMasterListByProcedure_notpaging
arrlist = ocsaslist.farrlist

if isarray(arrlist) then
    if ocsaslist.FResultCount=1 then
        ResultOneCsID = arrlist(0,0)
    end if
end if

Response.Buffer=true
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_CS처리리스트_" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
downFilemenupos=menupos
downPersonalInformation_rowcnt=ocsaslist.ftotalcount
%>
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 12">
<style type="text/css">
 td {font-size:8.0pt;}
 .txt {mso-number-format:"\@";}
 .num {mso-number-format:"0";}
 .prc {mso-number-format:"\#\,\#\#0";}
</style>
</head>
<body>
<!--[if !excel]>　　<![endif]-->
<div align=center x:publishsource="Excel">

<table width="100%" border="1" align="center" class="a csH15" cellpadding="2" cellspacing="1" bgcolor="#BABABA" style="table-layout:fixed">
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
		<td width="60" align="center">상품코드</td>
		<td width="50" align="center">옵션코드</td>
		<td width="80" align="center">상품명</td>
		<td width="80" align="center">옵션명</td>
		<td width="60" align="center">수량</td>
		<td width="90" align="center">제휴주문번호</td>
    </tr>
<% if isarray(arrlist) then %>
<% for i = 0 to ubound(arrlist,2) %>
    <% if (arrlist(16,i) <> "N") then %>
    <tr bgcolor="#EEEEEE" style="color:gray" align="center">
    <% else %>
	<tr bgcolor="#FFFFFF" align="center" >
    <% end if %>
        <td><%= arrlist(0,i) %></td>
        <td align="left"><%= arrlist(26,i) %></td>
        <td>
        		<%= arrlist(31,i) %>
        		<% if (arrlist(4,i) <> arrlist(31,i)) then %>
        			+
        		<% end if %>
        </td>
        <td><%= arrlist(24,i) %></td>
        <td align="left">
            <%= Left(arrlist(18,i),32) %>
		</td>
        <td>
			<%= AstarUserName(arrlist(5,i)) %>
        </td>
        <td class="txt" align="left">
			<%= AstarUserid(arrlist(6,i)) %>
        </td>
        <td align="left">
			<%= arrlist(9,i) %>
			<% if arrlist(24,i)<>"10x10" then %>(<%= arrlist(25,i) %>)<% end if %>
		</td>
        <td><%= GetCurrStateName(arrlist(10,i)) %></td>
		<td><%= arrlist(8,i) %></td>
		<td><%= arrlist(7,i) %></td>
        <td align="right"><%= FormatNumber(arrlist(22,i),0) %></td>
        <td><%= Left(arrlist(13,i),10) %></td>
		<td><%= Left(arrlist(15,i),10) %></td>
        <td><%= Left(arrlist(14,i),10) %></td>
        <td>
            <% if arrlist(16,i)="Y" then %>
                <font color="red">삭제</font>
            <% elseif arrlist(16,i)="C" then %>
                <font color="red"><strong>취소</strong></font>
            <% end if %>
        </td>
		<td>
			<%= arrlist(33,i) %>
		</td>
		<td class="txt">
			<%= arrlist(34,i) %>
		</td>
		<td>
			<%= arrlist(35,i) %>
		</td>
		<td>
			<%= arrlist(36,i) %>
		</td>
		<td>
			<%= arrlist(37,i) %>
		</td>
		<td class="txt">
			<% if arrlist(24,i) <> "10x10" then %>
				<%= arrlist(25,i) %>
			<% end if %>
		</td>
    </tr>
<%
if i mod 1000 = 0 then
    Response.Flush		' 버퍼리플래쉬
end if
next
end if
%>

</table>
</div>
</body>
</html>
<%
set ocsaslist = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
