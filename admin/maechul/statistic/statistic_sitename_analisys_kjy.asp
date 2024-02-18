<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  온라인 매출집계-판매처별
' History : 2012.10.09 강준구 생성
'			2013.01.08 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim i, cStatistic, vSiteName, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vIsBanPum, v6Ago
vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",-13,now())))
vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",-13,now())))
vEYear		= NullFillWith(request("eyear"),Year(now))
vEMonth		= NullFillWith(request("emonth"),Month(now))

Dim strSql, arrRows
strSql = "exec [db_statistics_order].[dbo].[usp_TEN_meachul_kjy] '"& vSYear & "-" & TwoNumber(vSMonth) & "-01" &"', '"& vEYear & "-" & TwoNumber(vEMonth) & "-01" &"'"
rsSTSget.CursorLocation = adUseClient
rsSTSget.CursorType = adOpenStatic
rsSTSget.LockType = adLockOptimistic
rsSTSget.Open strSql, dbSTSget
If Not(rsSTSget.EOF or rsSTSget.BOF) Then
	arrRows = rsSTSget.getRows
End If
rsSTSget.close

rw strSql
%>
<script language="javascript">
function searchSubmit(){
	frm.submit();
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
				<%
					'### 년
					Response.Write "<select name=""syear"" class=""select"">"
					For i=Year(now) To 2001 Step -1
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSYear),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### 월
					Response.Write "<select name=""smonth"" class=""select"">"
					For i=1 To 12
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSMonth),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'#############################

					'### 년
					Response.Write "<select name=""eyear"" class=""select"">"
					For i=Year(now) To 2001 Step -1 ''Year(v6MonthDate)
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEYear),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"

					'### 월
					Response.Write "<select name=""emonth"" class=""select"">"
					For i=1 To 12
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEMonth),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"
				%>
				&nbsp;&nbsp;
			</td>
		</tr>
	    </table>
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="검색" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<br>
* cjmall : 2021-03-11(상품분류개선) / 2021-11-03 (분류,카테고리매칭업무 제휴전달)</br>
* 11번가 : 2021-09-27부터 작업시작</br>
* lfmall : 2021-10-19(승인완료) / 2021-11-02(등록스케줄러 파이프라인 개선)</br>
* 인터파크 : 2021-12-15 상품 등록 수정 완료</br>
* 스토어팜 : 2022-03-31 등록스케줄러 파이프라인 개선</br>
<br>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td align="center">제휴몰</td>
	<td align="center">연월</td>
    <td align="center">주문수</td>
    <td align="center">구매총액</td>
    <td align="center">보너스쿠폰사용액</td>
    <td align="center">매출액</td>
</tr>
<%
If isArray(arrRows) Then
	For i=0 To Ubound(arrRows, 2)
%>
<tr <%= Chkiif(arrRows(6, i)="1","bgcolor=SKYBLUE","bgcolor=#FFFFFF") %>>
	<td align="center"><%= arrRows(0, i) %></td>
	<td align="center"><%= arrRows(1, i) %></td>
	<td align="center"><%= FormatNumber(arrRows(2, i), 0) %></td>
	<td align="center"><%= FormatNumber(arrRows(3, i), 0) %></td>
	<td align="center"><%= FormatNumber(arrRows(4, i), 0) %></td>
	<td align="center"><%= FormatNumber(arrRows(5, i), 0) %></td>
</tr>
<%
	Next
End If
%>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
