<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 주문 목록 엑셀 다운로드
' Hieditor : 2017.07.11 허진원 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<%

''사용안함. (-> 관리자 허용)
if Not(C_ManagerUpJob or C_ADMIN_AUTH) then
	response.end
end if

Response.ContentType = "application/vnd.ms-excel" 
Response.AddHeader  "Content-Disposition" , "attachment; filename=orderlist"& replace(date&hour(now)&minute(now),"-","") &".xls"

dim orderserial, searchtype, searchrect, yyyy1,yyyy2,mm1,mm2,dd1,dd2, page, ojumun, ix,iy, PageSize
dim nowdate,searchnextdate,research, jumundiv, sellchnl, cknodate,ckdelsearch,ckipkumdiv4,ckipkumdiv2, not3pl, ipkumdiv
	searchtype  = requestCheckVar(request("searchtype"),32)
	searchrect  = requestCheckVar(request("searchrect"),32)
	yyyy1       = requestCheckVar(request("yyyy1"),4)
	mm1         = requestCheckVar(request("mm1"),2)
	dd1         = requestCheckVar(request("dd1"),2)
	yyyy2       = requestCheckVar(request("yyyy2"),4)
	mm2         = requestCheckVar(request("mm2"),2)
	dd2         = requestCheckVar(request("dd2"),2)
	jumundiv    = requestCheckVar(request("jumundiv"),10)
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	cknodate    = request("cknodate")
	ckdelsearch = request("ckdelsearch")
	ckipkumdiv4 = request("ckipkumdiv4")
	orderserial = request("orderserial")
	ckipkumdiv2 = request("ckipkumdiv2")
	ipkumdiv	= requestCheckVar(request("ipkumdiv"),1)
	research    = request("research")
	not3pl = request("not3pl")

'// 엑셀로 저장할 최대 행수 (너무 많으면 타임아웃 걸림)
page=1
PageSize = 50000

nowdate = Left(CStr(now()),10)

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)

	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1
end if

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

if research="" then ckipkumdiv2="on"
if research="" then not3pl="on"
    
set ojumun = new CJumunMaster
	if (jumundiv="flowers") then
		ojumun.FRectIsFlower = "Y"
	elseif (jumundiv="minus") then
	    ojumun.FRectIsMinus = "Y"
	elseif (jumundiv="foreign") then
	    ojumun.FRectIsForeign = "Y"
	elseif (jumundiv="military") then
	    ojumun.FRectIsMilitary = "Y"
	elseif (jumundiv="pojang") then
	    ojumun.FRectPojangOrder = "Y"
    elseif (jumundiv="sendGift") then
        ojumun.FRectIsSendGift = "Y"
	end if
	
	if cknodate="" then
		ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
		ojumun.FRectRegEnd = searchnextdate
	end if
	
	if ckdelsearch<>"on" then
		ojumun.FRectDelNoSearch="on"
	end if

	if searchtype="01" then
		ojumun.FRectBuyname = searchrect
	elseif searchtype="02" then
		ojumun.FRectReqName = searchrect
	elseif searchtype="03" then
		ojumun.FRectUserID = searchrect
	elseif searchtype="04" then
		ojumun.FRectIpkumName = searchrect
	elseif searchtype="06" then
		ojumun.FRectSubTotalPrice = searchrect
	end if
	
	ojumun.FPageSize = 30
	ojumun.FRectIpkumDiv4 = ckipkumdiv4
	ojumun.FRectIpkumDiv2 = ckipkumdiv2
	ojumun.FRectIpkumDiv = ipkumdiv
	ojumun.FRectOrderSerial = orderserial
	ojumun.FCurrPage = page
	ojumun.FPageSize = PageSize
	ojumun.FRectSellChannelDiv = sellchnl
	ojumun.FRectExcept3pl = not3pl  ''2017/03/29 추가
	ojumun.SearchJumunList

%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<style type="text/css">
 body {font-family:tahoma;font-size:12px}
 table {padding:2px;border-spacing:0px;font-family:tahoma;font-size:12px;border-collapse:collapse}
 td {text-align:center}
 .titbg {background-color:#FEE;}
</style>
</head>
<body>
<table>
<tr>
	<td class="titbg">주문번호</td>
	<td class="titbg">국가</td>
	<td class="titbg">채널</td>
	<td class="titbg">Site</td>
	<td class="titbg">RdSite</td>
	<td class="titbg">UserID</td>

	<% if (C_InspectorUser = False) then %>
		<td class="titbg">등급</td>
	<% end if %>

	<% if (FALSE) then %>
		<td class="titbg">구매자</td>
		<td class="titbg">수령인</td>
    <% end if %>

	<% if (C_InspectorUser = False) then %>
		<td class="titbg">주문총액</td>
		<td class="titbg">보너스쿠폰</td>
		<td class="titbg">상품쿠폰</td>
		<td class="titbg">기타할인</td>
		<td class="titbg">마일리지</td>
	<% end if %>

	<td class="titbg">(실)결제액</td>
	<td class="titbg">결제방법</td>
	<td class="titbg">거래상태</td>
	<td class="titbg">삭제여부</td>
	<td class="titbg">주문일</td>
</tr>
<% if ojumun.FresultCount>0 then %>
	<% for ix=0 to ojumun.FresultCount-1 %>
	<tr <%=chkIIF(ojumun.FMasterItemList(ix).IsAvailJumun,"","style=""background-color:#EEE;""")%>>
		<td><%= ojumun.FMasterItemList(ix).FOrderSerial %></td>
		<td><%= CHKIIF(ojumun.FMasterItemList(ix).IsForeignDeliver,ojumun.FMasterItemList(ix).FDlvcountryCode,"") %></td>
		<td><%= getSellChannelDivName(ojumun.FMasterItemList(ix).Fbeadaldiv) %> </td>
		<td><font color="<%= ojumun.FMasterItemList(ix).SiteNameColor %>"><%= ojumun.FMasterItemList(ix).FSitename %></font></td>
		<td><%= ojumun.FMasterItemList(ix).FRdSite %></td>

		<% if ojumun.FMasterItemList(ix).UserIDName<>"&nbsp;" then %>
			<td><%= printUserId(ojumun.FMasterItemList(ix).UserIDName,2,"*") %></td>
		<% else %>
			<td></td>
		<% end if %>

		<% if (C_InspectorUser = False) then %>
			<td>
			    <% if ojumun.FMasterItemList(ix).FUserID="" then %>
	
			    <% else %>
					<font color="<%= getUserLevelColor(ojumun.FMasterItemList(ix).fUserLevel) %>"><%= getUserLevelStr(ojumun.FMasterItemList(ix).fUserLevel) %></font>
			    <% end if %>
			</td>
		<% end if %>

		<% if (FALSE) then %>
			<td><%= ojumun.FMasterItemList(ix).FBuyName %></td>
			<td><%= ojumun.FMasterItemList(ix).FReqName %></td>
		<% end if %>

		<% if (C_InspectorUser = False) then %>
			<td><%= FormatNumber(ojumun.FMasterItemList(ix).FTotalSum,0) %></td>
			<td><%= FormatNumber(ojumun.FMasterItemList(ix).Fcouponpay,0) %></td>
			<td><%= FormatNumber(ojumun.FMasterItemList(ix).getMayItemCouponDiscount,0) %></td>
			<td><%= FormatNumber(ojumun.FMasterItemList(ix).Fallatdiscountprice,0) %></td>
			<td><%= FormatNumber(ojumun.FMasterItemList(ix).Fmiletotalprice,0) %></td>
		<% end if %>

		<td><font color="<%= ojumun.FMasterItemList(ix).SubTotalColor%>"><%= FormatNumber(ojumun.FMasterItemList(ix).FSubTotalPrice,0) %></font></td>
		<td><%= ojumun.FMasterItemList(ix).JumunMethodName %></td>
		<td><font color="<%= ojumun.FMasterItemList(ix).IpkumDivColor %>"><%= ojumun.FMasterItemList(ix).IpkumDivName %></font></td>
		<td><font color="<%= ojumun.FMasterItemList(ix).CancelYnColor %>"><%= ojumun.FMasterItemList(ix).CancelYnName %></font></td>
		<td><%= Left(ojumun.FMasterItemList(ix).GetRegDate,16) %></td>
	</tr>
	<%
			if (ix mod 500)=0 then
				Response.Flush
			end if
		next
	%>
<% end if %>
</table>
</body>
</html>
<%
set ojumun = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
