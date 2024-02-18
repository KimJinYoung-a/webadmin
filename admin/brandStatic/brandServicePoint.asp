<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 월간브랜드서비스지수
' History : 서동석 생성
'			2023.11.16 한용민 수정(전시카테고리 검색 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/brand/brandClass.asp"-->
<%
dim yyyy, mm, makerID, i, dispCate, arrList, CBrandService
	dispCate = requestCheckvar(request("disp"),16)
	yyyy	= req("yyyy1", Left(Date,4))
	mm		= req("mm1", Mid(Date,6,2))
	makerID = req("makerID", "")

set CBrandService = new CBrandServiceList
	CBrandService.frectyyyy = yyyy
	CBrandService.frectmm = mm
	CBrandService.frectmakerID = makerID
	CBrandService.frectdispCate = dispCate
	CBrandService.fBrandServiceList()

if CBrandService.FtotalCount > 0 then
	arrList=CBrandService.fArrList
end if

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
	       	* 년월: &nbsp;<% DrawYMBox yyyy,mm %>
			* 브랜드ID: <input type="text" class="text" name="makerID" value="<%=makerID%>">
			* 전시카테고리: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
</table>
</form>
<!-- 검색 끝 -->
 
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		* 평균출고일 환산점수 : 100-(평균출고소요일*10) ---> 50점이 최하점수
		<Br>* 지연출고 환산점수 : 100-(지연출고건수/총출고건수*5) ---> 50점이 최하점수
		<Br>* CS클레임 환산점수 : 100-((품절취소+반품+맞교환등)/총출고건수*5) ---> 50점이 최하점수
		<Br>* 상품문의 환산점수 : 100-평균시간  ---> 50점이 최하점수

		<Br><Br>* 서비스지수 : 4개의 환산점수 평균(상품문의가 없을경우, 3개의 환산점수 평균)
		<Br>* 월 총출고건수가 10개 미만일 경우, 서비스지수 산정이 무의미할듯. 일단 모두 산정하고 나중에 검토
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="25">
        검색결과 : <b><%= CBrandService.FtotalCount %></b>
    </td>
</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td rowspan="2" width="120">
		<%If makerID <> "" Then %>
			년월
		<%Else %>
			브랜드ID
		<%End If%>
		</td>
		<td rowspan="2">총출고건수<br>(업체배송)</td>
        <td colspan="2">평균출고소요일</td>
        <td colspan="2">지연출고(D+4이상)</td>
        <td colspan="4">클레임관련</td>
        <td colspan="3">상품문의</td>
		<td rowspan="2" width="60"><b>서비스지수</b></td>
		<td colspan="3">상품후기</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td>평균출고<br>소요일</td>
        <td><b>환산<br>점수</b></td>
        <td>D+4이상<br>출고건수</td>
        <td><b>환산<br>점수</b></td>
        <td>취소<br>(품절)</td>
        <td>반품<br>(불량/오배송등)</td>
        <td>맞교환<br>(불량/오배송등)</td>
        <td><b>환산<br>점수</b></td>
        <td>상품문의건수</td>
        <td>답변소요시간</td>
        <td><b>환산<br>점수</b></td>
		<td>작성수</td>
		<td>1점후기</td>
		<td>평균<br>점수</td>
	</tr>
<%
Dim servicePoint, servicePointText, pnt1, pnt2, pnt3, pnt4
Dim rowCnt
Dim sRs(14)

If IsArray(arrList) Then 
	rowCnt = UBound(arrList,2) + 1
%>

	<%For i=0 To UBound(arrList,2)%>
    <tr align="center" bgcolor="#FFFFFF">
	<%
		' Row 합산
		sRs(1) = sRs(1) + CDbl(arrList(1,i))			'후기수
		sRs(2) = sRs(2) + CDbl(arrList(2,i))			'후기점수
		sRs(3) = sRs(3) + CDbl(arrList(3,i))			'상품문의수
		sRs(4) = sRs(4) + CDbl(arrList(4,i))			'답변소요시간
		sRs(5) = sRs(5) + CDbl(arrList(5,i))			'출고수
		sRs(6) = sRs(6) + CDbl(arrList(6,i))			'출고소요일
		sRs(7) = sRs(7) + CDbl(arrList(7,i))			'품절수
		sRs(8) = sRs(8) + CDbl(arrList(8,i))			'반품수
		sRs(9) = sRs(9) + CDbl(arrList(9,i))			'교환수
		sRs(10) = sRs(10) + CDbl(arrList(10,i))		'출고지연수
		sRs(11) = sRs(11) + CDbl(arrList(11,i))		'1점 후기수
		sRs(12) = sRs(12) + CDbl(arrList(12,i))		'2점 후기수
		sRs(13) = sRs(13) + CDbl(arrList(13,i))		'3점 후기수
		sRs(14) = sRs(14) + CDbl(arrList(14,i))		'4점 후기수

		' 평균 항목 재계산
		If CDbl(arrList(1,i)) > 0 Then
			arrList(2,i) = FormatNumber(CDbl(arrList(2,i)) / CDbl(arrList(1,i)) ,2)
		End If 
		If CDbl(arrList(3,i)) > 0 Then
			arrList(4,i) = FormatNumber(CDbl(arrList(4,i)) / CDbl(arrList(3,i)) ,1)
		End If 
		If CDbl(arrList(5,i)) > 0 Then
			arrList(6,i) = FormatNumber(CDbl(arrList(6,i)) / CDbl(arrList(5,i)) ,2)
		End If 

		' 서비스 지수 산출 공식
		pnt1 = 0
		pnt2 = 0
		pnt3 = 0
		pnt4 = 0

		' 총출고건수가 있을때
		If arrList(5,i) > 0 Then 
			''평균출고일 환산점수 : 100-(평균출고소요일*10) ---> 50점이 최하점수
			pnt1 = 100 - CInt(10 * CDbl(arrList(6,i)))
			If pnt1 < 50 Then pnt1 = 50

			''지연출고 환산점수 : 100-(지연출고건수/총출고건수%*5) ---> 50점이 최하점수
			pnt2 = 100 - CInt(500 * CDbl(arrList(10,i)) / CDbl(arrList(5,i)) )
			If pnt2 < 50 Then pnt2 = 50

			''CS클레임 환산점수 : 100-((품절취소+반품+맞교환등)/총출고건수%*5) ---> 50점이 최하점수
			pnt3 = 100 - CInt(500 * CDbl(arrList(7,i)+arrList(8,i)+arrList(9,i)) / CDbl(arrList(5,i)))
			If pnt3 < 50 Then pnt3 = 50
		End If 

		' 상품문의건수가 있을때
		If arrList(3,i) > 0 Then 
			''상품문의 환산점수 : 100-평균시간  ---> 50점이 최하점수
			pnt4 = 100 - CLng(arrList(4,i))  ''CInt => CLng ''2016/04/28
			If pnt4 < 50 Then pnt4 = 50
		End If 

		' 출고건수가 있을때
		If arrList(5,i) > 0 Then
			' 상품문의건수가 있을때
			If arrList(3,i) > 0 Then 
				servicePoint = (pnt1 + pnt2 + pnt3 + pnt4) / 4
			Else
				servicePoint = (pnt1 + pnt2 + pnt3) / 3
			End If 

			servicePointText = FormatNumber(servicePoint,2) & "점"
			sRs(0) = sRs(0) + servicePoint

		Else
			servicePointText = "-"
			rowCnt = rowCnt - 1
		End If 

		If pnt1 = 0 Then pnt1 = "-"
		If pnt2 = 0 Then pnt2 = "-"
		If pnt3 = 0 Then pnt3 = "-"
		If pnt4 = 0 Then pnt4 = "-"
	%>
		<td><%=arrList(0,i)%></td>
		<td><%=arrList(5,i)%></td>
		<td><%=arrList(6,i)%>일</td>
		<td><%=pnt1%></td>
		<td><%=arrList(10,i)%></td>
		<td><%=pnt2%></td>

		<td><%=arrList(7,i)%></td>
		<td><%=arrList(8,i)%></td>
		<td><%=arrList(9,i)%></td>
		<td><%=pnt3%></td>

		<td><%=arrList(3,i)%></td>
		<td><%=arrList(4,i)%>시간</td>
		<td><%=pnt4%></td>
    	<td><%=servicePointText%></td>
		<td><%=arrList(1,i)%></td>
		<td><%=arrList(11,i)%></td>
		<td><%=arrList(2,i) %>점</td>
	</tr>
	<%Next%>
    <tr align="center" bgcolor="#FFFFFF">
	<%
		If CDbl(sRs(1)) > 0 Then
			sRs(2) = FormatNumber(CDbl(sRs(2)) / CDbl(sRs(1)) ,2)
		End If 
		If CDbl(sRs(3)) > 0 Then
			sRs(4) = FormatNumber(CDbl(sRs(4)) / CDbl(sRs(3)) ,1)
		End If 
		If CDbl(sRs(5)) > 0 Then
			sRs(6) = FormatNumber(CDbl(sRs(6)) / CDbl(sRs(5)) ,2)
		End If
	%>
    	<td><b>합계 or 평균</b></td>
		<td><b><%=FormatNumber(sRs(5),0)%></b></td>
		<td><b><%=sRs(6)%></b>일</td>
		<td>&nbsp;</td>
		<td><b><%=FormatNumber(sRs(10),0)%></b></td>
		<td>&nbsp;</td>

		<td><b><%=FormatNumber(sRs(7),0)%></b></td>
		<td><b><%=FormatNumber(sRs(8),0)%></b></td>
		<td><b><%=FormatNumber(sRs(9),0)%></b></td>
		<td>&nbsp;</td>

		<td><b><%=FormatNumber(sRs(3),0)%></b></td>
		<td><b><%=sRs(4)%></b>시간</td>
		<td>&nbsp;</td>
		<td><b><%=FormatNumber( sRs(0) / rowCnt ,2) %></b>점</td>
		<td><b><%=FormatNumber(sRs(1),0)%></b></td>
		<td><b><%=FormatNumber(sRs(11),0)%></b></td>
		<td><b><%=sRs(2)%></b>점</td>
    </tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<%
End If 
%>
</table>

<%
set CBrandService = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
