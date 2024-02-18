<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/analysiscls.asp"-->

<%
response.write "수정중 : <a href='/admin/maechul/maechul_sum.asp?menupos=1013'><font color=blue>매출통계v2&gt;&gt;매출통계 참조</font></a>"
dbget.close()	:	response.End

dim yyyy1,mm1
dim yyyy2,mm2
yyyy1 = request("yyyy1")
mm1 = request("mm1")
'yyyy2 = request("yyyy2")
'mm2 = request("mm2")


dim dt
if yyyy1="" then
	'dt = dateserial(year(Now),month(now)-3,1)
	'yyyy1 = Left(CStr(dt),4)
	'mm1 = Mid(CStr(dt),6,2)

	dt = dateserial(year(Now),month(now),1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)

	dt = dateserial(year(Now),month(now)+1,1)
	yyyy2 = Left(CStr(dt),4)
	mm2 = Mid(CStr(dt),6,2)
else
	dt = CStr(dateserial(yyyy1,mm1+1,1))
	yyyy2 = Left(CStr(dt),4)
	mm2 = Mid(CStr(dt),6,2)
end if


dim nextyyyymm
'nextyyyymm = CStr(dateserial(yyyy2,mm2+1,1))
'response.write yyyy1 + "-" + mm1 + "-01"
'response.write yyyy2 + "-" + mm2 + "-01"
dim oanal
set oanal = new CAnalysis
oanal.FRectYYYYMMDD = yyyy1 + "-" + mm1 + "-01"
oanal.FRectYYYYMMDD2 = yyyy2 + "-" + mm2 + "-01"

oanal.FBeasongPay = 2700
oanal.GetMeachulWithCopons
oanal.GetMinusMeachulSum
oanal.GetTenBeasongcount
oanal.GetMeaipSum

'oanal.getOnlineMonthlyGain

dim i

'response.write "FRectYYYYMMDD : " + oanal.FRectYYYYMMDD + "<br>"
'response.write "FRectYYYYMMDD2 : " + oanal.FRectYYYYMMDD2 + "<br>"
%>
<table width="800" border="0" cellpadding="2" cellspacing="1" class="a">
<tr>
	<td> (배송비단가 2,700), 카드수수료, 제휴몰수수료 등 제외</td>
</tr>
</table>
<table width="800" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="rectorder" value="">
	<tr>
		<td class="a" >
		검색기간:<% DrawYMBox yyyy1,mm1 %>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="800" border="0" cellpadding="2" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF" align=center>
	<td width=150>검색월</td>
	<td width=200><%= oanal.FOneItem.Fyyyymm %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>결제건수</td>
	<td><%= FormatNumber(oanal.FOneItem.FMCnt,0) %> 건</td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>주문금액</td>
	<td><%= FormatNumber(oanal.FOneItem.FTotalSum,0) %> 원</td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>결제금액</td>
	<td><%= FormatNumber(oanal.FOneItem.FSubTotalPrice,0) %> 원</td>
	<td>
		<% if oanal.FOneItem.FTotalSum<>0 then %>
		<%= clng(oanal.FOneItem.FSubTotalPrice/oanal.FOneItem.FTotalSum*100*100)/100 %> %
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>객단가</td>
	<td>
		<% if oanal.FOneItem.FMCnt<>0 then %>
		<%= FormatNumber(oanal.FOneItem.FSubTotalPrice/oanal.FOneItem.FMCnt,0) %> 원
		<% end if %>
	</td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td colspan="3"></td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>반품건수</td>
	<td><%= FormatNumber(oanal.FOneItem.Fminuscnt,0) %> 건</td>
	<td>
		<% if oanal.FOneItem.FMCnt<>0 then %>
		<%= clng(oanal.FOneItem.Fminuscnt/oanal.FOneItem.FMCnt*100*100)/100 %> % (결제건수 대비)
		<% end if %>
	</td>
</tr>
<!--
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>반품주문금액</td>
	<td><%= FormatNumber(oanal.FOneItem.FminusTotalSum,0) %> 원</td>
	<td>

	</td>
</tr>
-->
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>반품결제금액</td>
	<td><%= FormatNumber(oanal.FOneItem.FminusSubTotalPrice,0) %> 원</td>
	<td>
		<% if oanal.FOneItem.FSubTotalPrice<>0 then %>
		<%= clng(Abs(oanal.FOneItem.FminusSubTotalPrice)/Abs(oanal.FOneItem.FSubTotalPrice)*100*100)/100 %> % (결제액 대비)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td colspan="3"></td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>쿠폰사용</td>
	<td>
		<%= FormatNumber(oanal.FOneItem.Ftencardspend,0) %> 원
	</td>
	<td>
		<% if oanal.FOneItem.FTotalSum<>0 then %>
		<%= clng(oanal.FOneItem.Ftencardspend/oanal.FOneItem.FTotalSum*100*100)/100 %> %  (주문액대비)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>마일리지</td>
	<td>
		<%= FormatNumber(oanal.FOneItem.Fmiletotalprice,0) %> 원
	</td>
	<td>
		<% if oanal.FOneItem.FTotalSum<>0 then %>
		<%= clng(oanal.FOneItem.Fmiletotalprice/oanal.FOneItem.FTotalSum*100*100)/100 %> %  (주문액대비)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>SKT멤버십</td>
	<td>
		<%= FormatNumber(oanal.FOneItem.Fspendmembership,0) %> 원
	</td>
	<td>
		<% if oanal.FOneItem.FTotalSum<>0 then %>
		<%= clng(oanal.FOneItem.Fspendmembership/oanal.FOneItem.FTotalSum*100*100)/100 %> %  (주문액대비)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>올엣할인</td>
	<td>
		<%= FormatNumber(oanal.FOneItem.Fallatdiscountprice,0) %> 원
	</td>
	<td>
		<% if oanal.FOneItem.FTotalSum<>0 then %>
		<%= clng(oanal.FOneItem.Fallatdiscountprice/oanal.FOneItem.FTotalSum*100*100)/100 %> %  (주문액대비)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>할인소계</td>
	<td>
		<%= FormatNumber(oanal.FOneItem.getTotalDiscountsum,0) %> 원
	</td>
	<td>
		<% if oanal.FOneItem.FTotalSum<>0 then %>
		<%= clng(oanal.FOneItem.getTotalDiscountsum/oanal.FOneItem.FTotalSum*100*100)/100 %> %  (주문액대비)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td colspan="3"></td>
</tr>
</table>


<br>
<table width="800" border="0" cellpadding="2" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>배송건수<br>(텐배 포함 건수)</td>
	<td width=200><%= Formatnumber(oanal.FOneItem.FBeasongCnt,0) %></td>
	<td>
		<% if oanal.FOneItem.FMCnt<>0 then %>
		<%= clng(oanal.FOneItem.FBeasongCnt/oanal.FOneItem.FMCnt*100*100)/100 %> %
		<% end if %>
		(결제건수 대비)
	</td>
</tr>
</table>

<br>
<table width="800" border="0" cellpadding="2" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>실결제액</td>
	<td width=200><%= Formatnumber(oanal.FOneItem.getRealSubTotalPrice,0) %></td>
	<td>(결제금액 - 반품결제금액)</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>상품매입가</td>
	<td width=200><%= Formatnumber(oanal.FOneItem.FMeaipTotal,0) %></td>
	<td>
		<% if oanal.FOneItem.getRealSubTotalPrice<>0 then %>
		<%= clng(oanal.FOneItem.FMeaipTotal/(oanal.FOneItem.getRealSubTotalPrice)*100*100)/100 %> %  (실결제액 대비)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>배송비</td>
	<td width=200><%= formatnumber(oanal.FOneItem.GetBeasongTotal,0) %></td>
	<td>(텐바이텐배송건수 * 배송단가(<%= formatnumber(oanal.FBeasongPay,0) %>)</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>카드수수료</td>
	<td width=200></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>제휴수수료</td>
	<td width=200></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td width=150>수익</td>
	<td width=200><%= formatnumber(oanal.FOneItem.GetSuic,0) %></td>
	<td>
		<% if oanal.FOneItem.getRealSubTotalPrice<>0 then %>
		<%= clng(oanal.FOneItem.GetSuic/(oanal.FOneItem.getRealSubTotalPrice)*100*100)/100 %> %  (실결제액 대비)
		<% end if %>
		(실결제액대비)
	</td>
</tr>
<tr bgcolor="#FFFFFF" align=center>
	<td colspan="3"></td>
</tr>
</table>

<br><br>
<!--
<br><br>

<table width="800" border="0" cellpadding="2" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF" align=center>
	<td width=60>총결제건수</td>
	<td width=70>총주문금액</td>
	<td width=70>총결제금액<br>(A)</td>
	<td width=70>객단가</td>
	<td width=70>총매입가<br>(B)</td>
	<td width=70>배송건수</td>
	<td width=70>배송금액<br>(C)</td>
	<td width=70>마일리지사용</td>
	<td width=70>쿠폰사용</td>
	<td width=70>수익<br>(A-B-C)</td>
	<td width=70>수익율</td>
</tr>
<% if false then %>
<% for i=0 to oanal.FResultCount -1 %>
<tr bgcolor="#FFFFFF">
	<td align=center><%= FormatNumber(oanal.FItemList(i).FMCnt,0) %></td>
	<td align=right><%= FormatNumber(oanal.FItemList(i).FTotalSum,0) %></td>
	<td align=right><%= FormatNumber(oanal.FItemList(i).FSubTotalPrice,0) %></td>
	<td align=right>
	<% if oanal.FItemList(i).FMCnt<>0 then %>
	<%= FormatNumber(oanal.FItemList(i).FSubTotalPrice/oanal.FItemList(i).FMCnt,0) %>
	<% end if %>
	</td>
	<td align=right><%= FormatNumber(oanal.FItemList(i).FMeaipTotal,0) %></td>
	<td align=center><%= FormatNumber(oanal.FItemList(i).FBeasongCnt,0) %></td>
	<td align=right><%= FormatNumber(oanal.FItemList(i).GetBeasongTotal,0) %></td>
	<td align=right><%= FormatNumber(oanal.FItemList(i).Fmiletotalprice,0) %></td>
	<td align=right><%= FormatNumber(oanal.FItemList(i).Ftencardspend,0) %></td>
	<td align=right><%= FormatNumber(oanal.FItemList(i).GetSuic,0) %></td>
	<td align=center>
	<% if oanal.FItemList(i).FSubTotalPrice<>0 then %>
		<%= FormatNumber(oanal.FItemList(i).GetSuic/oanal.FItemList(i).FSubTotalPrice*100,0) %> %
	<% end if %>
	</td>
</tr>
<% next %>
<% end if %>
</table>
-->
<%
set oanal = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->