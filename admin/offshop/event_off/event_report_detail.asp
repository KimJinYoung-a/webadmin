<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  오프라인이벤트 통계
' History : 2010.03.25 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/eventreport_Cls.asp"-->
<%
dim SType , evt_code,i ,page, grpWidth ,shopid ,t_TotalCost, t_FTotalNo,yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim fromDate , toDate ,ftotselljumuncnt ,datefg, inc3pl
	datefg = requestCheckVar(request("datefg"),10)
	SType = requestCheckVar(request("SType"),10)
	evt_code = requestCheckVar(request("evt_code"),10)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if datefg = "" then datefg = "event"
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
			
fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

'/매장
if (C_IS_SHOP) then
	
	'//직영점일때
	if C_IS_OWN_SHOP then
		
		'/어드민권한 점장 미만
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if		
	else
		shopid = C_STREETSHOPID
	end if
else
	'/업체
	if (C_IS_Maker_Upche) then

	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if	

dim oReport
set oReport = new Cevtreport_list
	oReport.FRectevt_code = evt_code
	oReport.FRectshopid = shopid
	oReport.FRectStartDay = fromDate
	oReport.FRectEndDay = toDate
	oReport.frectdatefg = datefg
	oReport.FRectInc3pl = inc3pl
	
t_TotalCost = 0
t_FTotalNo  = 0
ftotselljumuncnt = 0
%>

<script language="javascript">
	
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
	
	function regsubmit(){
		frm.submit();
	}

	//상품매출
	function item_detail(shopid,evt_code,yyyy1,mm1,dd1,yyyy2,mm2,dd2){
		location.href='?evt_code='+evt_code+'&shopid='+shopid+'&SType=T&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&menupos=<%=menupos%>&datefg=jumun';
	}

</script>

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">  
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* 기간 :
				<% draweventmaechul_datefg "datefg" ,datefg ," onchange='regsubmit()'"%>				
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				&nbsp;&nbsp;
				<%
				'직영/가맹점
				if (C_IS_SHOP) then
				%>	
				
					<% if not C_IS_OWN_SHOP and shopid <> "" then %>
						* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
					<% end if %>
				<% else %>
					* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
				<% end if %>
				<p>
				* 이벤트번호 : <input type="text" name="evt_code" size="10" value="<%= evt_code %>">
	            &nbsp;&nbsp;
	            <b>* 매출처구분</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>				
			</td>
		</tr>
		</table> 
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="regsubmit();">
	</td>
</tr>
</table>
<!-- 표 상단바 끝-->
<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
    </td>
    <td align="right">
		분류: 
		<input type="radio" name="SType" value="D" <% If SType = "D" Then response.write "checked" %> onclick="regsubmit();">날짜별
		<input type="radio" name="SType" value="T" <% If SType = "T" Then response.write "checked" %> onclick="regsubmit();">상품별    	
    </td>        
</tr>
</form>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<%
'// 날짜별 이벤트 통계 
if SType = "D" then

	'//통계에서 가져옴
	oReport.geteventdate_sum()
%>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			검색결과 : <b><%= oReport.FResultCount %></b> ※ 총1000건 까지만 검색 됩니다.
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>구매일</td>
		<td>매장</td>
		<td>이벤트<Br>코드</td>
		<td>이벤트명</td>
		<td>매출액</td>
		<td>매출<br>건수</td>
		<td>판매<br>수량</td>
		<td>그래프</td>
		<td>비고</td>
	</tr>
	<% 
	if oReport.FResultCount > 0 then 
	for i=0 to oReport.FResultCount-1
	
	t_TotalCost = t_TotalCost + oReport.FItemList(i).fsellsum
	t_FTotalNo  = t_FTotalNo + oReport.FItemList(i).fsum_cnt
	ftotselljumuncnt = ftotselljumuncnt + oReport.FItemList(i).ftotselljumuncnt
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td><%= oReport.FItemList(i).fshopregdate %></td>
		<td><%= oReport.FItemList(i).fshopid %></td>
		<td><%= oReport.FItemList(i).fevt_code %></td>
		<td><%= oReport.FItemList(i).fevt_name %></td>
		<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(oReport.FItemList(i).fsellsum,0) %></td>
		<td align="right"><%= FormatNumber(oReport.FItemList(i).ftotselljumuncnt,0) %></td>
		<td align="right"><%= oReport.FItemList(i).fsum_cnt %></td>
		<td width="200" align="left">
			<%				
				if oReport.maxc>0 then
					grpWidth = Clng(oReport.FItemList(i).fsellsum/oReport.maxc*200)
				else
					grpWidth = 0
				end if
			%>
			<img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%=grpWidth%>">
		</td>
		<td width=100>
			<input type="button" class="button" value="상품상세" onclick="item_detail('<%= oReport.FItemList(i).fshopid %>','<%= oReport.FItemList(i).fevt_code %>','<%= left(oReport.FItemList(i).fshopregdate,4) %>','<%= mid(oReport.FItemList(i).fshopregdate,6,2) %>','<%= right(oReport.FItemList(i).fshopregdate,2) %>','<%= left(oReport.FItemList(i).fshopregdate,4) %>','<%= mid(oReport.FItemList(i).fshopregdate,6,2) %>','<%= right(oReport.FItemList(i).fshopregdate,2) %>');">		
		</td>		
	</tr>
	<% 
	next
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td colspan=4>총합</td>		
		<td align="right"><%= FormatNumber(t_TotalCost,0) %></td>
		<td align="right"><%= FormatNumber(ftotselljumuncnt,0) %></td>
		<td align="right"><%= FormatNumber(t_FTotalNo,0) %></td>
		<td colspan=3></td>
	</tr>
		
	<% ELSE %>
	<tr  align="center" bgcolor="#FFFFFF">
		<td colspan="20">등록된 내용이 없습니다.</td>
	</tr>
	<% 
	end if 
	%>
<% 
'// 상품별 이벤트 통계 
elseif SType = "T" then 

	oReport.geteventitem_sum()
%>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			검색결과 : <b><%= oReport.FResultCount %></b> ※ 총1000건 까지만 검색 됩니다.
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>매장</td>
		<td>이벤트<Br>코드</td>	
		<td>이벤트명</td>
		<td>상품코드<br>브랜드</td>
		<td>상품명<font color='blue'>(옵션명)<font></td>		
		<td>매출액</td>
		<td>판매<br>수량</td>
		<td>그래프</td>
	</tr>
	<% if oReport.FResultCount > 0 then %>
	<% for i=0 to oReport.FResultCount-1 %>
	<%
	t_TotalCost = t_TotalCost + oReport.FItemList(i).fsellsum
	t_FTotalNo  = t_FTotalNo + oReport.FItemList(i).fsum_cnt
	%>
	<tr bgcolor="#FFFFFF" align="center">
		<td>
			<%= oReport.FItemList(i).fshopid %>
		</td>		
		<td>
			<%= oReport.FItemList(i).fevt_code %>
		</td>
		<td>
			<%= oReport.FItemList(i).fevt_name %>
		</td>				
		<td>
			<%= oReport.FItemList(i).fitemgubun %>-<%= oReport.FItemList(i).FItemid %>-<%= oReport.FItemList(i).fitemoption %><br><%= oReport.FItemList(i).fmakerid %>
		</td>
		<td>			
			<%= oReport.FItemList(i).fitemname %>
			<% 
			if oReport.FItemList(i).fitemoption <> "0000" then
				response.write "<font color='blue'>("&oReport.FItemList(i).fitemoptionname&")<font>"
			end if
			%>
		</td>
		<td align="right" bgcolor="#E6B9B8">
			<%= FormatNumber(oReport.FItemList(i).fsellsum,0) %>
		</td>	
		<td align="right">
			<%= oReport.FItemList(i).fsum_cnt %>
		</td>
		<td align="left" width=200>
			<%			
			if oReport.maxc>0 then
				grpWidth = Clng(oReport.FItemList(i).fsellsum/oReport.maxc*200)
			else
				grpWidth = 0
			end if
			%>
			<img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%=grpWidth%>">
		</td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan=5 align="center">총합</td>
		<td align="right">
			<%= FormatNumber(t_TotalCost,0) %>
		</td>
		<td align="right">
			<%= FormatNumber(t_FTotalNo,0) %>
		</td>		
		<td></td>
	</tr>
	<% ELSE %>
	<tr  align="center" bgcolor="#FFFFFF">
		<td colspan="20">등록된 내용이 없습니다.</td>
	</tr>	
	<% end if %>
<%
end if
%>
</table>	

<%
set oReport = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->