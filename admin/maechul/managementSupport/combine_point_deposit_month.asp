<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  온라인 오프라인 마일리지 & 예치금 통합관리
' History : 2013.11.12 한용민 생성
'           2018.03.12 허진원 - 마일리지 구분 추가(구매/프로모션)
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/mileage/combine_point_deposit_cls.asp" -->
<%
Dim i, yyyy1,mm1,yyyy2,mm2, fromDate ,toDate ,ocombine, srcGbn, targetGbn
	yyyy1   = requestcheckvar(request("yyyy1"),10)
	mm1     = requestcheckvar(request("mm1"),10)
	yyyy2   = requestcheckvar(request("yyyy2"),10)
	mm2     = requestcheckvar(request("mm2"),10)
	srcGbn     = requestcheckvar(request("srcGbn"),1)
	targetGbn     = requestcheckvar(request("targetGbn"),4)

if (yyyy1="") then yyyy1 = Cstr(Year( dateadd("m",-3,date()) ))
if (mm1="") then mm1 = Cstr(Month( dateadd("m",-3,date()) ))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (srcGbn="") then srcGbn="M"

fromDate = left(DateSerial(yyyy1, mm1,"01"),7)
toDate = left(DateSerial(yyyy2, mm2+1,"01"),7)

Set ocombine = New ccombine_point_deposit
	ocombine.FRectStartdate = fromDate
	ocombine.FRectEndDate = toDate
	ocombine.FRectsrcGbn = srcGbn
	ocombine.FRecttargetGbn = targetGbn
	ocombine.FPageSize = 500
	ocombine.FCurrPage	= 1
	ocombine.fcombine_point_deposit_month()

'행 표시 구분량
dim rowSpanNo: rowSpanNo=1
dim colSpanNo: colSpanNo=1
if srcGbn="M" then
	rowSpanNo=2
	colSpanNo=2
end if

%>

<script language="javascript">
function searchSubmit(){
	frm.submit();
}

function pop_detail_list(yyyy1, mm1, dd1, yyyy2, mm2, dd2, GbnCd){
	var pop_detail_list = window.open('/admin/maechul/managementsupport/combine_point_deposit_list.asp?yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&srcGbn=<%=srcGbn%>&targetGbn=<%=targetGbn%>&GbnCd='+GbnCd+'&menupos=<%=menupos%>','pop_detail_list','width=1024,height=768,scrollbars=yes,resizable=yes');
	pop_detail_list.focus();
}

function refreshSmr(yyyymm){
    if (confirm(yyyymm+' 재작성 하시겠습니까?')){
        document.frmAct.mode.value="refreshpointDepositSummary";
        document.frmAct.yyyymm.value=yyyymm;
        document.frmAct.submit();
    }
}
</script>

<!-- 검색 시작 -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">검색</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				* 날짜 : <% DrawYMBoxdynamic "yyyy1",yyyy1,"mm1",mm1,"" %> ~ <% DrawYMBoxdynamic "yyyy2",yyyy2,"mm2",mm2,"" %>
				<p>
				* 구분 : <% drawoffshop_commoncode "srcGbn", srcGbn, "srcGbn", "MAIN", "", "  " %>
				&nbsp;&nbsp;
				* 채널 : <% drawoffshop_commoncode "targetGbn", targetGbn, "targetGbn", "MAIN", "", "  " %>
			</td>
		</tr>
	    </table>
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="검색" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<Br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<p />

* 당월 승인액 = 주문시사용 + 반품/취소적립 + 기타<br />
* 예치금/마일리지는 로그테이블에서 데이터 가져옴.<br />
* 예치금<br />
&nbsp; - 무통장환불 : 미수금 무통장환불 선수금(예치금환급) 과 금액이 일치해야 합니다.<br />
&nbsp; - 주문시사용 : 결제로그 결제방식 예치금과 금액이 일치해야 합니다.<br />

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= ocombine.FresultCount %></b> ※ 총 500건까지 검색 됩니다.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td rowspan="<%=rowSpanNo%>">날짜</td>
    <td rowspan="<%=rowSpanNo%>">누적합계(월말)</td>
    <td rowspan="<%=rowSpanNo%>">월별합계</td>
    <td rowspan="<%=rowSpanNo%>">주문적립</td>
    <% if (srcGbn="M") then %>
    <td colspan="<%=colSpanNo%>">이벤트적립</td>
    <% elseif (srcGbn="G") then %>
    <td>이벤트</td>
    <% else %>
    <td>반품/취소적립</td>
    <% end if %>
    <% if (srcGbn="G") then %>
    <td rowspan="<%=rowSpanNo%>">기프(팅)콘전환</td>
    <td rowspan="<%=rowSpanNo%>">기프트카드등록(+)</td>
    <% else %>
    <td rowspan="<%=rowSpanNo%>">상품사용기</td>
    <td rowspan="<%=rowSpanNo%>">CS적립</td>
    <% end if %>
    <% if (srcGbn="M") then %>
    <td rowspan="<%=rowSpanNo%>">오프>온 전환</td>
    <% elseif (srcGbn="G") then %>
    <td rowspan="<%=rowSpanNo%>">기프트카드등록(-)</td>
    <% else %>
    <td rowspan="<%=rowSpanNo%>">기프(팅)콘전환</td>
    <% end if %>
    <!-- 사용-->
    <td rowspan="<%=rowSpanNo%>">주문시사용</td>
    <% if (srcGbn="M") then %>
    <td rowspan="<%=rowSpanNo%>">기타사용</td>
    <% elseif (srcGbn="G") then %>
    <td rowspan="<%=rowSpanNo%>">무통장 환불</td>
    <% else %>
    <td rowspan="<%=rowSpanNo%>">무통장 환불</td>
    <% end if %>
    <td rowspan="<%=rowSpanNo%>">회원탈퇴</td>
    <td rowspan="<%=rowSpanNo%>">소멸</td>
    <td rowspan="<%=rowSpanNo%>">기타</td>
	<td rowspan="<%=rowSpanNo%>"><b>당월승인액</b></td>
    <% if (C_ADMIN_AUTH) then %><td rowspan="<%=rowSpanNo%>">ACT</td><% end if %>
</tr>
<% if (srcGbn="M") then %>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>구매</td>
	<td>이벤트</td>
</tr>
<% end if %>
<% if C_MngPowerUser or C_ADMIN_AUTH then %>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>YYYYMM</td>
    <td>accpointsum</td>
	<td>pointsum</td>
	<td>ORD</td>
	<td colspan="<%=colSpanNo%>">GNE</td>
	<td>GNI</td>
	<td>GNC</td>
	<td>SFT</td>
	<td>SPO</td>
	<td>SPE</td>
	<td>RTD</td>
    <td>XPR</td>
	<td>ETC</td>
	<td></td>
	<td>
	    <img src="/images/icon_reload.gif" onClick="refreshSmr('<%= Left(dateAdd("m",-1,now()),7) %>');" style="cursor:pointer;" alt="재작성" title="<%= Left(dateAdd("m",-1,now()),7) %> 내역 재작성">
	    <% if (day(now())>5) then '// 2017-09-07, 20일에서 5일로, skyer9 %>
	    ,
	    <img src="/images/icon_reload.gif" onClick="refreshSmr('<%= Left(dateAdd("m",0,now()),7) %>');" style="cursor:pointer;" alt="재작성" title="<%= Left(dateAdd("m",0,now()),7) %> 내역 재작성">
	    <% end if %>
	</td>
</tr>
<% end if %>
<%
dim totETC, totGNE, totSPE, totGNC, totGNI, totSPO, totSFT, totRTD, totORD, totpointsum, totXPR, totGOE, totGPE
	totETC=0
	totGNE=0
	totGOE=0		'// 이벤트 마일리지 구매적립(jucyo:1100)
	totGPE=0		'// 이벤트 마일리지 프로모션적립(jucyo:1000)
	totSPE=0
	totGNC=0
	totGNI=0
	totSPO=0
	totSFT=0
	totRTD=0
	totORD=0
	totXPR=0
	totpointsum=0

Dim oPoint : oPoint=0
Dim PPoint : PPoint=0

if srcGbn="M" and toDate="2013-12" and targetGbn="ONAC" then
    oPoint =1092762032
end if

if srcGbn="M" and toDate="2013-12" and targetGbn="OF" then
    oPoint =99979777
end if

if srcGbn="M" and toDate="2013-12" and targetGbn="" then
    oPoint =99979777+1092762032
end if

if srcGbn="M" and toDate="2013-12" and targetGbn="AC" then
    oPoint =8887135+4590279
end if


if srcGbn="M" and toDate="2013-12" and targetGbn="ON" then
    oPoint =1092762032-(8887135+4590279)
end if


if ocombine.FresultCount > 0 then

For i = 0 To ocombine.FresultCount -1
totpointsum = totpointsum + ocombine.fitemlist(i).fpointsum
totETC = totETC + ocombine.fitemlist(i).fETC
totGNE = totGNE + ocombine.fitemlist(i).fGNE
totGOE = totGOE + ocombine.fitemlist(i).fGOE
totGPE = totGPE + ocombine.fitemlist(i).fGPE
totSPE = totSPE + ocombine.fitemlist(i).fSPE
totGNC = totGNC + ocombine.fitemlist(i).fGNC
totGNI = totGNI + ocombine.fitemlist(i).fGNI
totSPO = totSPO + ocombine.fitemlist(i).fSPO
totSFT = totSFT + ocombine.fitemlist(i).fSFT
totRTD = totRTD + ocombine.fitemlist(i).fRTD
totORD = totORD + ocombine.fitemlist(i).fORD
totXPR = totXPR + ocombine.fitemlist(i).fXPR

oPoint = oPoint-pPoint
%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFFFFF';>
	<td height="25">
		<%= ocombine.fitemlist(i).fYYYYMM %>
	</td>
	<td align="right" bgcolor="#9DCFFF">
         <%= FormatNumber(ocombine.fitemlist(i).faccpointsum,0) %>
	</td>
	<td align="right" bgcolor="#E6B9B8">
        <%= FormatNumber(ocombine.fitemlist(i).fpointsum,0) %>
    </td>
	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','ORD');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fORD,0) %>
		</a>
	</td>
	<% if (srcGbn="M") then %>
	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','GOE');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fGOE,0) %>
		</a>
	</td>
	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','GPE');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fGPE,0) %>
		</a>
	</td>
	<% else %>
	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','GNE');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fGNE,0) %>
		</a>
	</td>
	<% end if %>
	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','GNI');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fGNI,0) %>
		</a>
	</td>
	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','GNC');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fGNC,0) %>
		</a>
	</td>
	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','SFT');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fSFT,0) %>
		</a>
	</td>
	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','SPO');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fSPO,0) %>
		</a>
	</td>
    <td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','SPE');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fSPE,0) %>
		</a>
	</td>
	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','RTD');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fRTD,0) %>
		</a>
	</td>
	<td align="right">
	    <a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','XPR');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fXPR,0) %>
		</a>
	</td>

	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','ETC');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fETC,0) %>
		</a>
	</td>

	<td align="right">
		<b><%
		Select Case srcGbn
			Case "D"
				response.write FormatNumber((ocombine.fitemlist(i).fGNE + ocombine.fitemlist(i).fSPO + ocombine.fitemlist(i).fETC),0)
			Case Else
				response.write FormatNumber((ocombine.fitemlist(i).fSPO + ocombine.fitemlist(i).fETC),0)
		End Select
		%></b>
	</td>

	<% if (C_ADMIN_AUTH) then %><td>
	    <% if (DateAdd("m",+3,CDate(ocombine.fitemlist(i).fYYYYMM+"-01"))>now()) and (ocombine.fitemlist(i).fYYYYMM>="2014-01") then %>
	    <img src="/images/icon_reload.gif" onClick="refreshSmr('<%= ocombine.fitemlist(i).fYYYYMM %>');" style="cursor:pointer;">
	    <% end if %>
	</td><% end if %>
</tr>
<%
PPoint = ocombine.fitemlist(i).fpointsum
%>
<% next %>

<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td height="25">합계</td>
    <td align="right"></td>
    <td align="right"><%= FormatNumber(totpointsum,0) %></td>
	<td align="right">
		<%= FormatNumber(totORD,0) %>
	</td>
	<% if (srcGbn="M") then %>
	<td align="right">
		<%= FormatNumber(totGOE,0) %>
	</td>
	<td align="right">
		<%= FormatNumber(totGPE,0) %>
	</td>
	<% else %>
	<td align="right">
		<%= FormatNumber(totGNE,0) %>
	</td>
	<% end if %>
	<td align="right">
		<%= FormatNumber(totGNI,0) %>
	</td>
	<td align="right">
		<%= FormatNumber(totGNC,0) %>
	</td>
	<td align="right">
		<%= FormatNumber(totSFT,0) %>
	</td>
	<td align="right">
		<%= FormatNumber(totSPO,0) %>
	</td>
    <td align="right">
		<%= FormatNumber(totSPE,0) %>
	</td>
	<td align="right">
		<%= FormatNumber(totRTD,0) %>
	</td>
    <td align="right">
		<%= FormatNumber(totXPR,0) %>
	</td>

	<td align="right">
		<%= FormatNumber(totETC,0) %>
	</td>
	<td align="right">
	</td>
	<% if (C_ADMIN_AUTH) then %><td></td><% end if %>
</tr>

<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="21">등록된 내용이 없습니다.</td>
</tr>
<% end if %>
</table>
<form name="frmAct" method="post" action="pointsum_process.asp">
<input type="hidden" name="mode">
<input type="hidden" name="yyyymm">
</form>
<%
Set ocombine = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
