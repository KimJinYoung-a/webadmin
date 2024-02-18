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
<%

dim v3MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay
dim xl

v3MonthDate	= DateAdd("m",-3,now())

vSYear		= NullFillWith(request("syear"),Year(v3MonthDate))
vSMonth		= NullFillWith(request("smonth"),Month(v3MonthDate))
vSDay		= NullFillWith(request("sday"),"01")
vEYear		= NullFillWith(request("eyear"),Year(now))
vEMonth		= NullFillWith(request("emonth"),Month(now))
vEDay		= NullFillWith(request("eday"),Day(now))
xl			= NullFillWith(request("xl"),"")

dim cStatistic
Set cStatistic = New cStaticTotalClass_list
cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
cStatistic.fStatistic_NewItemMeachul()

dim i, j, k, arrCate(20)

For i = 0 To cStatistic.FTotalCount - 1
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

dim totHTML, newHTML, rateHTML
dim curYYYY, curWeekno, found
dim totReducedPriceSUM, totReducedNoSUM, newTotReducedPriceSUM, newTotReducedNoSUM

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
				결제일자 : <% DrawDateBoxdynamic vSYear, "syear", vEYear, "eyear", vSMonth, "smonth", vEMonth, "emonth", vSDay, "sday", vEDay, "eday" %>
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
			* 결제일자 기준 자료입니다.(일단위로 최근 2주 데이타를 갱신합니다.)
		</td>
		<td align="right">
			<input type="button" class="button" value="엑셀받기" onClick="popXL()">
		</td>
	</tr>
</table>

<p />

<div style="overflow:scroll; width:100%; height:85%; padding:0px;">
<% end if %>
<table width="4000" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="table-layout:fixed;">
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" rowspan="2" colspan="2">구분</td>
	<td align="center" colspan="2">전체</td>
	<% for j = 0 to 19 %>
	<% if (arrCate(j) <> "") then %>
	<td align="center" colspan="2"><%= arrCate(j) %></td>
	<% end if %>
	<% next %>
    <td align="center" width="50" rowspan="2">비고</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td align="center" width="100">판매액</td>
	<td align="center" width="80">상품수</td>
	<% for j = 0 to 19 %>
	<% if (arrCate(j) <> "") then %>
    <td align="center" width="100">판매액</td>
	<td align="center" width="80">상품수</td>
	<% end if %>
	<% next %>
</tr>
<%
For i = 0 To cStatistic.FTotalCount - 1
	if (curYYYY <> cStatistic.FList(i).Fyyyy or curWeekno <> cStatistic.FList(i).Fweekno) then
		curYYYY = cStatistic.FList(i).Fyyyy
		curWeekno = cStatistic.FList(i).Fweekno

		totHTML = ""
		newHTML = ""
		rateHTML = ""
		totReducedPriceSUM = 0
		totReducedNoSUM = 0
		newTotReducedPriceSUM = 0
		newTotReducedNoSUM = 0

		for j = 0 to 19
			found = False
			if (arrCate(j) <> "") then
				For k = 0 To cStatistic.FTotalCount - 1
					if (cStatistic.FList(k).Fyyyy = curYYYY) and (cStatistic.FList(k).Fweekno = curWeekno) and (cStatistic.FList(k).FcateFullName = arrCate(j)) then
						totHTML = totHTML + "<td align='center'>" & FormatNumber(cStatistic.FList(k).FtotReducedPrice,0) & "</td><td align='center'>" & FormatNumber(cStatistic.FList(k).FtotReducedNo,0) & "</td>"
						newHTML = newHTML + "<td align='center'>" & FormatNumber(cStatistic.FList(k).FnewTotReducedPrice,0) & "</td><td align='center'>" & FormatNumber(cStatistic.FList(k).FnewTotReducedNo,0) & "</td>"
						if cStatistic.FList(k).FtotReducedPrice <> 0 and cStatistic.FList(k).FtotReducedNo <> 0 then
							if (100.0*cStatistic.FList(k).FnewTotReducedPrice/cStatistic.FList(k).FtotReducedPrice >= 5.0) then
								rateHTML = rateHTML + "<td align='center'><b><font color='red'>" & FormatNumber(100.0*cStatistic.FList(k).FnewTotReducedPrice/cStatistic.FList(k).FtotReducedPrice,2) & "%</font></b></td>"
							else
								rateHTML = rateHTML + "<td align='center'>" & FormatNumber(100.0*cStatistic.FList(k).FnewTotReducedPrice/cStatistic.FList(k).FtotReducedPrice,2) & "%</td>"
							end if
							rateHTML = rateHTML + "<td align='center'>" & FormatNumber(100.0*cStatistic.FList(k).FnewTotReducedNo/cStatistic.FList(k).FtotReducedNo,2) & "%</td>"
						else
							rateHTML = rateHTML + "<td align='center'></td><td align='center'></td>"
						end if

						totReducedPriceSUM = totReducedPriceSUM + cStatistic.FList(k).FtotReducedPrice
						totReducedNoSUM = totReducedNoSUM + cStatistic.FList(k).FtotReducedNo
						newTotReducedPriceSUM = newTotReducedPriceSUM + cStatistic.FList(k).FnewTotReducedPrice
						newTotReducedNoSUM = newTotReducedNoSUM + cStatistic.FList(k).FnewTotReducedNo

						found = True
					end if
				next
				if found = False then
					totHTML = totHTML + "<td align='center'></td><td align='center'></td>"
					newHTML = newHTML + "<td align='center'></td><td align='center'></td>"
					rateHTML = rateHTML + "<td align='center'></td><td align='center'></td>"
				end if
			else
				exit for
			end if
		next
		%>
<tr bgcolor="#FFFFFF">
	<td align="center" rowspan="3" width="150">
		<%= cStatistic.FList(i).Fyyyy %>년 <%= cStatistic.FList(i).Fweekno %>주<br />
		(<%= Right(cStatistic.FList(i).FStartDate,5) %> ~ <%= Right(cStatistic.FList(i).FEndDate,5) %>)
	</td>
	<td align="center" width="90" height="25">
		전체
	</td>
	<td align="center"><%= FormatNumber(totReducedPriceSUM,0) %></td>
	<td align="center"><%= FormatNumber(totReducedNoSUM,0) %></td>
    <%= totHTML %>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" height="25">
		신상품
	</td>
	<td align="center"><%= FormatNumber(newTotReducedPriceSUM,0) %></td>
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
<% next %>
</table>
<% if (xl <> "Y") then %>
</div>

<form name="frmXL" method="get" style="margin:0px;">
	<input type="hidden" name="xl" value="Y">
	<input type="hidden" name="syear" value="<%= vSYear %>">
	<input type="hidden" name="eyear" value="<%= vEYear %>">
	<input type="hidden" name="smonth" value="<%= vSMonth %>">
	<input type="hidden" name="emonth" value="<%= vEMonth %>">
	<input type="hidden" name="sday" value="<%= vSDay %>">
	<input type="hidden" name="eday" value="<%= vEDay %>">
</form>

<% end if %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
