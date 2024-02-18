<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_analisys.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp" -->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%

dim vSYear, vSMonth, YYYYMM, makerid
dim xl, vPurchasetype, page, pageSize

page		= NullFillWith(request("page"),1)
vSYear		= NullFillWith(request("syear"),Year(date))
vSMonth		= NullFillWith(request("smonth"),NUM2STR(Month(date),2,"0","R"))
YYYYMM		= vSYear & "-" & vSMonth
makerid		= requestCheckvar(request("makerid"),32)
xl			= NullFillWith(request("xl"),"")
vPurchasetype = requestCheckvar(request("purchasetype"),2)

if xl="Y" then
	pageSize = 20000
else
	pageSize = 100
end if

dim i, j, k, arrCate(20), cateCount
dim totHTML, newHTML, rateHTML, sumTitleHTML
dim curMakerid, found
dim totReducedPriceSUM, totProfitSUM, totReducedNoSUM, newTotReducedPriceSUM, newTotProfitSUM, newTotReducedNoSUM

dim cStatistic
Set cStatistic = New cStaticTotalClass_list
cStatistic.FCurrPage = page
cStatistic.FPageSize = pageSize
cStatistic.FRectYYYYMM = YYYYMM
cStatistic.FRectMakerID = makerid
cStatistic.FRectPurchasetype = vPurchasetype

'// 총계 취합
cStatistic.fStatistic_Monthly_brand_NewItemMeachul_Total()

cateCount = cStatistic.FResultCount
if cStatistic.FResultCount>0 then
	For i = 0 To cStatistic.FResultCount - 1
		totHTML = totHTML + "<td align='center'>" & FormatNumber(cStatistic.FList(i).FtotReducedPrice,0) & "</td><td align='center'>" & FormatNumber(cStatistic.FList(i).FtotProfit,0) & "</td><td align='center'>" & FormatNumber(cStatistic.FList(i).FtotReducedNo,0) & "</td>"
		newHTML = newHTML + "<td align='center'>" & FormatNumber(cStatistic.FList(i).FnewTotReducedPrice,0) & "</td><td align='center'>" & FormatNumber(cStatistic.FList(i).FnewTotProfit,0) & "</td><td align='center'>" & FormatNumber(cStatistic.FList(i).FnewTotReducedNo,0) & "</td>"
		if cStatistic.FList(i).FtotReducedPrice <> 0 and cStatistic.FList(i).FtotReducedNo <> 0 then
			if (100.0*cStatistic.FList(i).FnewTotReducedPrice/cStatistic.FList(i).FtotReducedPrice >= 5.0) then
				rateHTML = rateHTML + "<td align='center'><b><font color='red'>" & FormatNumber(100.0*cStatistic.FList(i).FnewTotReducedPrice/cStatistic.FList(i).FtotReducedPrice,2) & "%</font></b></td>"
			else
				rateHTML = rateHTML + "<td align='center'>" & FormatNumber(100.0*cStatistic.FList(i).FnewTotReducedPrice/cStatistic.FList(i).FtotReducedPrice,2) & "%</td>"
			end if
			rateHTML = rateHTML + "<td align='center'>" & FormatNumber(100.0*cStatistic.FList(i).FnewTotProfit/cStatistic.FList(i).FtotProfit,2) & "%</td>"
			rateHTML = rateHTML + "<td align='center'>" & FormatNumber(100.0*cStatistic.FList(i).FnewTotReducedNo/cStatistic.FList(i).FtotReducedNo,2) & "%</td>"
		else
			rateHTML = rateHTML + "<td align='center'></td><td align='center'></td><td align='center'></td>"
		end if

		totReducedPriceSUM = totReducedPriceSUM + cStatistic.FList(i).FtotReducedPrice
		totProfitSUM = totProfitSUM + cStatistic.FList(i).FtotProfit
		totReducedNoSUM = totReducedNoSUM + cStatistic.FList(i).FtotReducedNo
		newTotReducedPriceSUM = newTotReducedPriceSUM + cStatistic.FList(i).FnewTotReducedPrice
		newTotProfitSUM = newTotProfitSUM + cStatistic.FList(i).FnewTotProfit
		newTotReducedNoSUM = newTotReducedNoSUM + cStatistic.FList(i).FnewTotReducedNo
	Next

	sumTitleHTML = "<tr bgcolor=""#EFEFEF"" style=""font-weight:bold;"">"
	sumTitleHTML = sumTitleHTML & "<td align=""center"" rowspan=""3"" width=""150"">총 계</td>"
	sumTitleHTML = sumTitleHTML & "<td align=""center"" width=""90"" height=""25"">전체</td>"
	sumTitleHTML = sumTitleHTML & "<td align=""center"">" & FormatNumber(totReducedPriceSUM,0) & "</td>"
	sumTitleHTML = sumTitleHTML & "<td align=""center"">" & FormatNumber(totProfitSUM,0) & "</td>"
	sumTitleHTML = sumTitleHTML & "<td align=""center"">" & FormatNumber(totReducedNoSUM,0) & "</td>"
	sumTitleHTML = sumTitleHTML & totHTML & "<td></td>"
	sumTitleHTML = sumTitleHTML & "</tr>"
	sumTitleHTML = sumTitleHTML & "<tr bgcolor=""#EFEFEF"" style=""font-weight:bold;"">"
	sumTitleHTML = sumTitleHTML & "<td align=""center"" height=""25"">신상품</td>"
	sumTitleHTML = sumTitleHTML & "<td align=""center"">" & FormatNumber(newTotReducedPriceSUM,0) & "</td>"
	sumTitleHTML = sumTitleHTML & "<td align=""center"">" & FormatNumber(newTotProfitSUM,0) & "</td>"
	sumTitleHTML = sumTitleHTML & "<td align=""center"">" & FormatNumber(newTotReducedNoSUM,0) & "</td>"
	sumTitleHTML = sumTitleHTML & newHTML& "<td></td>"
	sumTitleHTML = sumTitleHTML & "</tr>"
	sumTitleHTML = sumTitleHTML & "<tr bgcolor=""#E8E8E8"">"
	sumTitleHTML = sumTitleHTML & "<td align=""center"" height=""25"">비율</td>"
	sumTitleHTML = sumTitleHTML & "<td align=""center"">" & chkIIF(totReducedNoSUM<>0,FormatNumber(100.0*newTotReducedPriceSUM/totReducedPriceSUM,2),0) & "%</td>"
	sumTitleHTML = sumTitleHTML & "<td align=""center"">" & chkIIF(totProfitSUM<>0,FormatNumber(100.0*newTotProfitSUM/totProfitSUM,2),0) & "%</td>"
	sumTitleHTML = sumTitleHTML & "<td align=""center"">" & chkIIF(totReducedNoSUM<>0,FormatNumber(100.0*newTotReducedNoSUM/totReducedNoSUM,2),0) & "%</td>"
	sumTitleHTML = sumTitleHTML & rateHTML& "<td></td>"
	sumTitleHTML = sumTitleHTML & "</tr>"
end if

'카테고리 필드 지정
For i = 0 To cStatistic.FResultCount - 1
	'FcateFullName
	for j = 0 to 19
		if (arrCate(j) = "") then
			arrCate(j) = cStatistic.FList(i).FcateFullName
			exit for
		elseif (arrCate(j) = cStatistic.FList(i).FcateFullName) then
			exit for
		end if
	next
next


'// 통계 리스트
cStatistic.fStatistic_Monthly_brand_NewItemMeachul()

if (xl = "Y") then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=new_item_xl.xls"
else
%>

<script>
function searchSubmit()
{
    frm.submit();
}

function popXL()
{
    frmXL.submit();
}
</script>

<!-- 검색 시작 -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				결제월 : <% DrawYMBoxdynamic "syear", vSYear, "smonth", vSMonth, "" %>
				&nbsp;
				구매유형: 
				<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
				&nbsp;
				브랜드: <%	drawSelectBoxDesignerWithName "makerid", makerid %>
			</td>
		</tr>
	    </table>
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="검색" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<p />

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#EEEEEE">
	<tr>
		<td align="left">
			* 결제일자 기준 자료입니다.(일단위로 데이타를 갱신합니다.)
		</td>
		<td align="right">
			<input type="button" class="button" value="엑셀받기" onClick="popXL()">
		</td>
	</tr>
</table>

<p />

<div style="overflow:scroll; width:100%; height:100%; padding:0px;">
<% end if %>
<table align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="table-layout:fixed;min-width:<%=(cateCount+3)*270%>px; ">
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" rowspan="2" colspan="2">구분</td>
	<td align="center" colspan="3">전체</td>
	<% for j = 0 to 19 %>
	<% if (arrCate(j) <> "") then %>
	<td align="center" colspan="3"><%= arrCate(j) %></td>
	<% end if %>
	<% next %>
    <td align="center" width="50" rowspan="2">비고</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td align="center" width="100">판매액</td>
	<td align="center" width="90">수익</td>
	<td align="center" width="80">상품수</td>
	<% for j = 0 to 19 %>
	<% if (arrCate(j) <> "") then %>
    <td align="center" width="100">판매액</td>
	<td align="center" width="90">수익</td>
	<td align="center" width="80">상품수</td>
	<% end if %>
	<% next %>
</tr>
<%=sumTitleHTML%>
<%
For i = 0 To cStatistic.FResultCount - 1
	if (curMakerid <> cStatistic.FList(i).Fmakerid) then
		curMakerid = cStatistic.FList(i).Fmakerid

		totHTML = ""
		newHTML = ""
		rateHTML = ""
		totReducedPriceSUM = 0
		totProfitSum = 0
		totReducedNoSUM = 0
		newTotReducedPriceSUM = 0
		newTotProfitSUM = 0
		newTotReducedNoSUM = 0

		for j = 0 to 19
			found = False
			if (arrCate(j) <> "") then
				For k = 0 To cStatistic.FResultCount - 1
					if (cStatistic.FList(k).Fmakerid = curMakerid) and (cStatistic.FList(k).FcateFullName = arrCate(j)) then
						totHTML = totHTML + "<td align='center'>" & FormatNumber(cStatistic.FList(k).FtotReducedPrice,0) & "</td><td align='center'>" & FormatNumber(cStatistic.FList(k).FtotProfit,0) & "</td><td align='center'>" & FormatNumber(cStatistic.FList(k).FtotReducedNo,0) & "</td>"
						newHTML = newHTML + "<td align='center'>" & FormatNumber(cStatistic.FList(k).FnewTotReducedPrice,0) & "</td><td align='center'>" & FormatNumber(cStatistic.FList(k).FnewTotProfit,0) & "</td><td align='center'>" & FormatNumber(cStatistic.FList(k).FnewTotReducedNo,0) & "</td>"
						if cStatistic.FList(k).FtotReducedPrice <> 0 and cStatistic.FList(k).FtotReducedNo <> 0 then
							if (100.0*cStatistic.FList(k).FnewTotReducedPrice/cStatistic.FList(k).FtotReducedPrice >= 5.0) then
								rateHTML = rateHTML + "<td align='center'><b><font color='red'>" & FormatNumber(100.0*cStatistic.FList(k).FnewTotReducedPrice/cStatistic.FList(k).FtotReducedPrice,2) & "%</font></b></td>"
							else
								rateHTML = rateHTML + "<td align='center'>" & FormatNumber(100.0*cStatistic.FList(k).FnewTotReducedPrice/cStatistic.FList(k).FtotReducedPrice,2) & "%</td>"
							end if
							rateHTML = rateHTML + "<td align='center'>" & FormatNumber(100.0*cStatistic.FList(k).FnewTotProfit/cStatistic.FList(k).FtotProfit,2) & "%</td>"
							rateHTML = rateHTML + "<td align='center'>" & FormatNumber(100.0*cStatistic.FList(k).FnewTotReducedNo/cStatistic.FList(k).FtotReducedNo,2) & "%</td>"
						else
							rateHTML = rateHTML + "<td align='center'></td><td align='center'></td><td align='center'></td>"
						end if

						totReducedPriceSUM = totReducedPriceSUM + cStatistic.FList(k).FtotReducedPrice
						totProfitSUM = totProfitSUM + cStatistic.FList(k).FtotProfit
						totReducedNoSUM = totReducedNoSUM + cStatistic.FList(k).FtotReducedNo
						newTotReducedPriceSUM = newTotReducedPriceSUM + cStatistic.FList(k).FnewTotReducedPrice
						newTotProfitSUM = newTotProfitSUM + cStatistic.FList(k).FnewTotProfit
						newTotReducedNoSUM = newTotReducedNoSUM + cStatistic.FList(k).FnewTotReducedNo

						found = True
					end if
				next
				if found = False then
					totHTML = totHTML + "<td align='center'></td><td align='center'></td><td align='center'></td>"
					newHTML = newHTML + "<td align='center'></td><td align='center'></td><td align='center'></td>"
					rateHTML = rateHTML + "<td align='center'></td><td align='center'></td><td align='center'></td>"
				end if
			else
				exit for
			end if
		next
		%>
<tr bgcolor="#FFFFFF">
	<td align="center" rowspan="3" width="150">
		<%= cStatistic.FList(i).Fmakerid %>
	</td>
	<td align="center" width="90" height="25">
		전체
	</td>
	<td align="center"><%= FormatNumber(totReducedPriceSUM,0) %></td>
	<td align="center"><%= FormatNumber(totProfitSUM,0) %></td>
	<td align="center"><%= FormatNumber(totReducedNoSUM,0) %></td>
    <%= totHTML %>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" height="25">
		신상품
	</td>
	<td align="center"><%= FormatNumber(newTotReducedPriceSUM,0) %></td>
	<td align="center"><%= FormatNumber(newTotProfitSUM,0) %></td>
	<td align="center"><%= FormatNumber(newTotReducedNoSUM,0) %></td>
	<%= newHTML %>
	<td></td>
</tr>
<tr bgcolor="#EEEEEE">
	<td align="center" height="25">
		비율
	</td>
	<td align="center">
		<%
		if totReducedPriceSUM <> 0 and totReducedNoSUM <> 0 then
			response.write FormatNumber(100.0*newTotReducedPriceSUM/totReducedPriceSUM,2) & "%"
		end if
		%>
	</td>
	<td align="center">
		<%
		if totReducedPriceSUM <> 0 and totProfitSUM <> 0 then
			response.write FormatNumber(100.0*newTotProfitSUM/totProfitSUM,2) & "%"
		end if
		%>
	</td>
	<td align="center">
		<%
		if totReducedPriceSUM <> 0 and totReducedNoSUM <> 0 then
			response.write FormatNumber(100.0*newTotReducedNoSUM/totReducedNoSUM,2) & "%"
		end if
		%>
	</td>
	<%= rateHTML %>
	<td></td>
</tr>
		<%
	end if
%>
<%
		if (i mod 500)=0 then Response.Flush
	next
%>
<% if (xl <> "Y") then %>
<tr>
	<td colspan="5" align="center" style="background-color:#F2F2F2;"><%sbDisplayPaging "page", page, cStatistic.FTotalCount, pageSize, 5,menupos %></td>
	<td colspan="55" style="background-color:#F2F2F2;"></td>
</tr>
<% end if %>
</table>
<br /><br />
<% if (xl <> "Y") then %>
</div>

<form name="frmXL" method="get" style="margin:0px;">
	<input type="hidden" name="xl" value="Y">
	<input type="hidden" name="page" value="<%= page %>">
	<input type="hidden" name="syear" value="<%= vSYear %>">
	<input type="hidden" name="smonth" value="<%= vSMonth %>">
	<input type="hidden" name="makerid" value="<%= makerid %>">
	<input type="hidden" name="vPurchasetype" value="<%= vPurchasetype %>">
</form>
<% end if %>
<%
	Set cStatistic = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
