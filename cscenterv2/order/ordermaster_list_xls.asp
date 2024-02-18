<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/order/ordercls.asp"-->
<%

'// 사용불가(에러내자), skyer9, 2016-09-20
1

dim searchfield, userid, orderserial, username, userhp, etcfield, etcstring
dim checkYYYYMMDD, checkJumunDiv, checkJumunSite
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim jumundiv, jumunsite
dim research
dim AlertMsg

'==============================================================================
searchfield = requestCheckvar(request("searchfield"),16)
userid 		= requestCheckvar(request("userid"),32)
orderserial = requestCheckvar(request("orderserial"),32)
username 	= requestCheckvar(request("username"),32)
userhp 		= requestCheckvar(request("userhp"),32)
etcfield 	= requestCheckvar(request("etcfield"),32)
etcstring 	= requestCheckvar(request("etcstring"),32)

checkYYYYMMDD = requestCheckvar(request("checkYYYYMMDD"),1)
checkJumunDiv = requestCheckvar(request("checkJumunDiv"),1)
checkJumunSite = requestCheckvar(request("checkJumunSite"),1)

yyyy1 = requestCheckvar(request("yyyy1"),4)
mm1 = requestCheckvar(request("mm1"),2)
dd1 = requestCheckvar(request("dd1"),2)
yyyy2 = requestCheckvar(request("yyyy2"),4)
mm2 = requestCheckvar(request("mm2"),2)
dd2 = requestCheckvar(request("dd2"),2)

jumundiv = requestCheckvar(request("jumundiv"),16)
jumunsite = requestCheckvar(request("jumunsite"),16)
research = requestCheckvar(request("research"),2)

if (research="") and (checkYYYYMMDD="") then checkYYYYMMDD="Y"
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

page = requestCheckvar(request("page"),10)
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
%>
<%	'엑셀 출력시작
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=DIY주문내역(" & yyyy1 & "-" & mm1 & "-" & dd1 & " ~ " & yyyy2 & "-" & mm2 & "-" & dd2 & ").xls"
%>
<html>
<body>
<table border="1" style="font-size:10pt;">
<tr>
	<td>구분</td>
	<td>주문구분</td>
	<td>주문번호</td>
	<td>Site</td>
	<td>UserID</td>
	<td>구매자</td>
	<td>수령인</td>
	<td>구매총액</td>
	<td>쿠폰</td>
	<td>마일리지</td>
	<td>기타할인</td>
	<td><b>결제금액</b></td>
	<td>결제방법</td>
	<td>거래상태</td>
	<td>주문일</td>
	<td>입금확인일</td>
	<td>발주일</td>
</tr>

<% if ojumun.FresultCount<1 then %>
<tr>
	<td colspan="20" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>

<% for ix=0 to ojumun.FresultCount-1 %>

<% if ojumun.FItemList(ix).IsAvailJumun then %>
<tr align="center" bgcolor="#FFFFFF">
<% else %>
<tr align="center" bgcolor="#EEEEEE">
<% end if %>
	<td><font color="<%= ojumun.FItemList(ix).CancelYnColor %>"><%= ojumun.FItemList(ix).CancelYnName %></font></td>
	<td>
	    <% if (ojumun.FItemList(ix).IsForeignDeliver) then %>
	    <strong>해외</strong>
	    <% elseif (ojumun.FItemList(ix).IsArmiDeliver) then %>
	    <strong>군부대</strong>
	    <% else %>
	    <%= ojumun.FItemList(ix).GetJumunDivName %>
	    <% end if %>
	</td>
	<td style='mso-number-format:"\@"'><%= ojumun.FItemList(ix).FOrderSerial %></td>
	<td><font color="<%= ojumun.FItemList(ix).SiteNameColor %>"><%= ojumun.FItemList(ix).FSitename %></font></td>
	<td align="left" style='mso-number-format:"\@"'>
	    <% if (ojumun.FItemList(ix).FSitename<>MAIN_SITENAME1 and ojumun.FItemList(ix).FSitename<>MAIN_SITENAME2) then %>
	    <%= ojumun.FItemList(ix).FAuthCode %>
	    <% else %>
	    <a href="?searchfield=userid&userid=<%= ojumun.FItemList(ix).FUserID %>"><font color="<%= ojumun.FItemList(ix).GetUserLevelColor %>"><%= ojumun.FItemList(ix).FUserID %></font></a>
	    <% end if %>
	</td>
	<td><%= ojumun.FItemList(ix).FBuyName %></td>
	<td><%= ojumun.FItemList(ix).FReqName %></td>
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
	<td align="right"><font color="<%= ojumun.FItemList(ix).SubTotalColor%>" ><b><%= FormatNumber(ojumun.FItemList(ix).FSubTotalPrice,0) %></b></font></td>

	<td><%= ojumun.FItemList(ix).JumunMethodName %></td>
	<% if ojumun.FItemList(ix).FIpkumdiv="1" then %>
	<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><acronym title="<%= ojumun.FItemList(ix).Fresultmsg %>"><%= ojumun.FItemList(ix).IpkumDivName %></acronym></font></td>
	<% else %>
	<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><%= ojumun.FItemList(ix).IpkumDivName %></font></td>
	<% end if %>
	<td><acronym title="<%= ojumun.FItemList(ix).FRegDate %>"><%= Left(ojumun.FItemList(ix).FRegDate,10) %></acronym></td>
	<td><acronym title="<%= ojumun.FItemList(ix).Fipkumdate %>"><%= Left(ojumun.FItemList(ix).Fipkumdate,10) %></acronym></td>
	<td><acronym title="<%= ojumun.FItemList(ix).Fbaljudate %>"><%= Left(ojumun.FItemList(ix).Fbaljudate,10) %></acronym></td>
</tr>
<% next %>

<% end if %>
</table>
</body>
</html>
<% set ojumun = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
