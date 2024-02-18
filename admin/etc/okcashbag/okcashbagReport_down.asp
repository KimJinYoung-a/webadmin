<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : OkCashbag관리
' History : 서동석 생성
'			2023.03.22 한용민 수정(권한 수기 아이디 박혀 있는부분 공통 권한 변수로 자동화. 소스 표준코드로 수정.)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/othermall/okcashbagCls.asp"-->
<%
'if (NOT C_ADMIN_AUTH) then
'    response.write "관리자만 접속 가능합니다. 관리자 문의 요망"
'    dbget.Close() :response.end
'end if

' 관리자 이거나 개발운영팀 이거나 제휴파트 일경우
If not(C_ADMIN_AUTH or C_SYSTEM_Part or C_partnership_part) Then
    response.write "관리자 및 해당 담당자만  접속 가능합니다. 관리자 문의 요망"
    dbget.Close() :response.end
end if

dim ArrIDX
ArrIDX = request("arod")

dim sSdate,sEdate, userid, orderserial, SearchDateType, vRdSite
sSdate 		= requestCheckVar(Request("iSD"),10)
sEdate 		= requestCheckVar(Request("iED"),10)
userid 		= requestCheckVar(Request("uId"),32)
orderserial	= requestCheckVar(Request("oSn"),12)
SearchDateType = requestCheckVar(request("dType"),2)
IF SearchDateType="" THEN SearchDateType="od"

vRdSite		= requestCheckVar(Request("rdsite"),10)
If vRdSite = "" Then
	vRdSite = "okcashbag"
End If

dim OrderType
OrderType = requestCheckVar(Request("otp"),2)
IF OrderType="" Then OrderType="no"

dim CurrPage
CurrPage = requestCheckVar(request("pg"),3)
IF CurrPage="" THEN CurrPage =1

dim sPageSize
	sPageSize = 10000	' 기존 코딩이 이렇게 되어 있어서 어쩔수 없어서 우선 제한 1만개로 박아놓음.. 이 이상 늘어날경우 페이징 구조를 getrows 로 받아와야함.

dim oCash,intLp
Set oCash = New CashbagCls
oCash.FCurrPage=CurrPage
oCash.FPageSize=sPageSize
oCash.FArrIDX = ArrIDX
oCash.FStartDate 	= sSdate
oCash.FEndDate 		= sEdate
oCash.Forderserial 	= orderserial
oCash.FOrderType 	= OrderType
oCash.FSearchType	= SearchDateType
oCash.FRdSite		= vRdSite

IF OrderType="N" Then 		'//정상건 업데이트
	oCash.updateNormalOrder()
ELSEIF OrderType ="C" Then	'//취소건  업데이트
	oCash.updateCancelOrder()
ELSEIF OrderType="UN" or OrderType ="UC" Then '// 출력 된 내역 (정상,취소)
	oCash.getUpdatedOrder()
END IF

downPersonalInformation_rowcnt=oCash.FTotalCount
%>
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
<%
dim SaveFilename
SaveFilename = "okcashbag.xls"

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_" & SaveFilename & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '버퍼사용여부
%>

<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<style type="text/css">
.mso {mso-number-format:"\@";}
		</style>

	</head>
	<body>
		<table width="100%" align="center" border="1" cellpadding="3" cellspacing="1" class="mso">
			<tr bgcolor="<%= adminColor("sky") %>">
	<!--<td align="center" width="20"><input type="checkbox" name="chkAll" onClick="jsChkAll(this.checked);"></td>-->
	<td align="center" width="100" >주문번호</td>
	<td align="center" width="80">장바구니번호</td>
	<td align="center">총결제금액</td>
	<td align="center">주문일자</td>
	<td align="center">배송일자</td>
	<td align="center">주문자</td>
	<td align="center">캐쉬백번호</td>
	<td align="center">적립포인트</td>
			</tr>
<% IF oCash.FResultcount<=0 Then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="10" align="center"> 일치하는 데이타가 없습니다.</td>
	</tr>
<% ELSE %>

	<% FOR intLp=0 To oCash.FResultcount-1 %>
	<tr bgcolor="#FFFFFF">
		<td align="center" style="mso-number-format:'\@'"><%= oCash.FItemList(IntLp).FOrderSerial %></td>
		<td align="center" style="mso-number-format:'\@'"><%= oCash.FItemList(IntLp).FShoppingBagNo %></td>
		<td align="center" style="mso-number-format:'\@'"><%= FormatNumber(oCash.FItemList(IntLp).FPointCash,0) %></td>
		<td align="center" style="mso-number-format:'\@'"><%= replace(DateValue (oCash.FItemList(IntLp).FRegdate),"-","") %></td>
		<td align="center" style="mso-number-format:'\@'"><%= replace(DateValue (oCash.FItemList(IntLp).FBeadaldate),"-","") %></td>
		<td align="center" style="mso-number-format:'\@'"><%= oCash.FItemList(IntLp).FBuyName %></td>
		<td align="center" style="mso-number-format:'\@'"><%= oCash.FItemList(IntLp).FCashBagCardNo %></td>
		<td align="center" style="mso-number-format:'\@'"><%= FormatNumber(oCash.FItemList(IntLp).FPoint,0) %></td>
	</tr>

	<%
        if intLp mod 500 = 0 then
            Response.Flush		' 버퍼리플래쉬
        end if
	NEXT
	%>
<% End IF %>
		</table>
	</body>
</html>
<script>opener.document.location.reload();</script>

<!-- 표 하단바 끝-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
