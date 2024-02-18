<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/lecture/lecturecls.asp"-->
<%

dim searchfield, userid, orderserial, username, userhp, etcfield, etcstring
dim checkYYYYMMDD, checkJumunDiv, checkJumunSite
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim jumundiv, jumunsite
dim research
dim AlertMsg

'==============================================================================
searchfield = RequestCheckvar(request("searchfield"),16)
userid 		= requestCheckvar(request("userid"),32)
orderserial = requestCheckvar(request("orderserial"),32)
username 	= requestCheckvar(request("username"),32)
userhp 		= requestCheckvar(request("userhp"),32)
etcfield 	= requestCheckvar(request("etcfield"),32)
etcstring 	= requestCheckvar(request("etcstring"),32)

checkYYYYMMDD = RequestCheckvar(request("checkYYYYMMDD"),1)
checkJumunDiv = RequestCheckvar(request("checkJumunDiv"),1)
checkJumunSite = RequestCheckvar(request("checkJumunSite"),1)

yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1 = RequestCheckvar(request("mm1"),2)
dd1 = RequestCheckvar(request("dd1"),2)
yyyy2 = RequestCheckvar(request("yyyy2"),4)
mm2 = RequestCheckvar(request("mm2"),2)
dd2 = RequestCheckvar(request("dd2"),2)

jumundiv = RequestCheckvar(request("jumundiv"),16)
jumunsite = RequestCheckvar(request("jumunsite"),16)
research = RequestCheckvar(request("research"),2)

'강좌는 수년전 강좌신청 내역도 검색한다.
if (research="") and (checkYYYYMMDD="") then checkYYYYMMDD=""
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
dim ojumun

page = RequestCheckvar(request("page"),10)
if (page="") then page=1

set ojumun = new COrderMaster
ojumun.FPageSize = 10000
ojumun.FCurrPage = page

if (checkYYYYMMDD="Y") then
	ojumun.FRectRegStart = Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
	ojumun.FRectRegEnd = searchnextdate
end if

if (checkJumunDiv = "Y") then
        if (jumundiv="flowers") then
        	ojumun.FRectIsFlower = "Y"
        elseif (jumundiv="minus") then
                ojumun.FRectIsMinus = "Y"
        elseif (jumundiv="foreign") then
                ojumun.FRectIsForeign = "Y"
        elseif (jumundiv="weclass") then
                ojumun.FRectIsWeClass = "Y"
        end if
end if

if (checkJumunSite = "Y") then
	ojumun.FRectExtSiteName = jumunsite
end if


if (searchfield = "orderserial") then
        '주문번호
        ojumun.FRectOrderSerial = orderserial
elseif (searchfield = "userid") then
        '고객아이디
        ojumun.FRectUserID = userid
elseif (searchfield = "username") then
        '구매자명
        ojumun.FRectBuyname = username
elseif (searchfield = "userhp") then
        '구매자핸드폰
        ojumun.FRectBuyHp = userhp
elseif (searchfield = "etcfield") then
        '기타조건
        if etcfield="01" then
        	ojumun.FRectBuyname = etcstring
        elseif etcfield="02" then
        	ojumun.FRectReqName = etcstring
        elseif etcfield="03" then
        	ojumun.FRectUserID = etcstring
        elseif etcfield="04" then
        	ojumun.FRectIpkumName = etcstring
        elseif etcfield="06" then
        	ojumun.FRectSubTotalPrice = etcstring
        elseif etcfield="07" then
        	ojumun.FRectBuyPhone = etcstring
        elseif etcfield="08" then
        	ojumun.FRectReqHp = etcstring
        elseif etcfield="09" then
        	ojumun.FRectReqSongjangNo = etcstring
        elseif etcfield="10" then
        	ojumun.FRectReqPhone = etcstring
        end if
end if

''검색조건 없을때 최근 N건 검색
ojumun.QuickSearchOrderList

'' 과거 6개월 이전 내역 검색
if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderList

    if (ojumun.FResultCount>0) then
        AlertMsg = "6개월 이전 주문입니다."
    end if
end if

dim ix,iy


'' 검색결과가 1개일대 디테일 자동으로 뿌림
dim ResultOneOrderserial
ResultOneOrderserial = ""
if (ojumun.FResultCount=1) then
    ResultOneOrderserial = ojumun.FItemList(0).FOrderSerial
end if

'<td style='mso-number-format:"\@"'></td>
%>
<%	'엑셀 출력시작
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=강의신청내역(" & yyyy1 & "-" & mm1 & "-" & dd1 & " ~ " & yyyy2 & "-" & mm2 & "-" & dd2 & ").xls"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body>
<table border="1" style="font-size:10pt;">
<tr>
	<td>구분</td>
	<td>단체</td>
	<td>주문번호</td>
	<td>UserID</td>
	<td>강좌명</td>
	<td>신청인</td>
	<td>구매총액</td>
	<td>쿠폰</td>
	<td>마일리지</td>
	<td>기타할인</td>
	<td><b>결제금액</b></td>
	<td>결제방법</td>
	<td>거래상태</td>
	<td>주문일</td>
	<td>입금확인일</td>
</tr>
<% if ojumun.FresultCount<1 then %>
<tr>
	<td colspan="15" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
	<% for ix=0 to ojumun.FresultCount-1 %>
		<% if ojumun.FItemList(ix).IsAvailJumun then %>
		<tr align="center">
		<% else %>
		<tr align="center" bgcolor="#EEEEEE">
		<% end if %>
			<td><font color="<%= ojumun.FItemList(ix).CancelYnColor %>"><%= ojumun.FItemList(ix).CancelYnName %></font></td>
			<td><% if ojumun.FItemList(ix).isWeClass then %><font color=blue>단체</font><% end if %></td>
			<td style='mso-number-format:"\@"'><%= ojumun.FItemList(ix).FOrderSerial %></td>
			<td align="left" style='mso-number-format:"\@"'>
			    <% if (ojumun.FItemList(ix).FSitename<>MAIN_SITENAME1 and ojumun.FItemList(ix).FSitename<>MAIN_SITENAME2) then %>
			    <%= ojumun.FItemList(ix).FAuthCode %>
			    <% else %>
			    <font color="<%= ojumun.FItemList(ix).GetUserLevelColor %>"><%= ojumun.FItemList(ix).FUserID %></font>
			    <% end if %>
			</td>
			<td align="left"><%= ojumun.FItemList(ix).Fgoodsname %></td>
			<td><%= ojumun.FItemList(ix).FBuyName %> <% if (ojumun.FItemList(ix).Fusercnt > 1) then %> 외 <%= (ojumun.FItemList(ix).Fusercnt - 1) %>인<% end if %></td>
			<% if (C_InspectorUser = False) then %>
			<td align="right"><%= FormatNumber(ojumun.FItemList(ix).FTotalSum,0) %></td>
			<td align="right"><%= FormatNumber(ojumun.FItemList(ix).Ftencardspend,0) %></td>
			<td align="right"><%= FormatNumber(ojumun.FItemList(ix).Fmiletotalprice,0) %></td>
			<td align="right">
			    <% if ojumun.FItemList(ix).Fallatdiscountprice<>0 then %>
			    <acronym title="<%= CHKIIF(ojumun.FItemList(ix).FAccountDiv="80","올엣할인","국민카드할인") %>"><%= FormatNumber(ojumun.FItemList(ix).Fallatdiscountprice+ ojumun.FItemList(ix).Fspendmembership,0) %></acronym>
			    <% else %>
			    <%= FormatNumber(ojumun.FItemList(ix).Fallatdiscountprice+ ojumun.FItemList(ix).Fspendmembership,0) %>
			    <% end if %>
			</td>
			<% end if %>
			<td align="right"><font color="<%= ojumun.FItemList(ix).SubTotalColor%>" ><b><%= FormatNumber(ojumun.FItemList(ix).FSubTotalPrice,0) %></b></font></td>
		
			<td><%= ojumun.FItemList(ix).JumunMethodName %></td>
			<% if ojumun.FItemList(ix).FIpkumdiv="1" then %>
			<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><acronym title="<%= ojumun.FItemList(ix).Fresultmsg %>"><%= ojumun.FItemList(ix).IpkumDivName %></acronym></font></td>
			<% else %>
			<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><%= ojumun.FItemList(ix).IpkumDivName %></font></td>
			<% end if %>
			<td><acronym title="<%= ojumun.FItemList(ix).FRegDate %>"><%= Left(ojumun.FItemList(ix).FRegDate,10) %></acronym></td>
			<td><acronym title="<%= ojumun.FItemList(ix).Fipkumdate %>"><%= Left(ojumun.FItemList(ix).Fipkumdate,10) %></acronym></td>
		</tr>
	<% next %>
<% end if %>
</table>
</body>
</html>
<% set ojumun = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
