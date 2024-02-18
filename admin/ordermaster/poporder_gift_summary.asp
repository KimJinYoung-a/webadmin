<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################################
'	2007년 11월 29일 한용민 개발
'	2008년 8월 21일 한용민 수정
'###########################################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/ordergiftcls.asp"-->

<%
dim evt_code, isupchebeasong, date_display , gift_code, chkOldOrder
dim viewType,page,dateview,dateview1 
dim i , balju_code
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim date1,date2,Edate
	balju_code = request("balju_code")
	dateview = request("dateview")
	dateview1 = request("dateview1")
	if dateview1 = "" then dateview1 = "yes" 	
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")	
	yyyy2 = request("yyyy2")
	dd2 = request("dd2")

	if yyyy1 = "" then yyyy1 = year(now)
	if mm1 = "" then mm1 = month(now)
	if dd1 = "" then dd1 = day(now)
	if yyyy2 = "" then yyyy2 = year(now)
	if mm2 = "" then mm2 = month(now)
	if dd2 = "" then dd2 = day(now)	 				 		
	if dateview<>"" then
		mm2 = Num2Str(request("mm2"),2,"0","R")
	end if
	if page="" then page=1
	gift_code	= request("gift_code")
	evt_code        = request("evt_code")
	isupchebeasong  = request("isupchebeasong")
	viewType        = request("viewType")
	date_display = request("date_display")
	chkOldOrder = request("chkOldOrder")

if (viewType="") then viewType="summary"

if viewType = "summary" then		'합계를 선택시 
	dim oOrderGiftcount
	set oOrderGiftcount = new COrderGift
	oOrderGiftcount.FPageSize = 500
	oOrderGiftcount.FCurrPage = page
	oOrderGiftcount.FRectisupchebeasong = isupchebeasong
	oOrderGiftcount.frectdateview = dateview
	oOrderGiftcount.frectdateview1 = dateview1	
	oOrderGiftcount.FRecteventid = evt_code
	oOrderGiftcount.FRectBaljuid = balju_code
	oOrderGiftcount.FRectgift_code = gift_code
	oOrderGiftcount.frectdate_display = date_display
	oOrderGiftcount.frectchkOldOrder = chkOldOrder
		if yyyy1 <> "" then
			oOrderGiftcount.FRectStartdate = yyyy1 & "-" & mm1 & "-" & dd1
		end if
		if yyyy2 <> "" then
			oOrderGiftcount.FRectEndDate = dateadd("d",1,yyyy2 & "-" & mm2 & "-" & dd2)	
		end if
	oOrderGiftcount.GeteventOrderGiftcount	
elseif viewType = "list" then		'내역리스트 선택시
	dim oOrderGift
	set oOrderGift = new COrderGift
	oOrderGift.FPageSize = 500
	oOrderGift.FCurrPage = page
	oOrderGift.FRectisupchebeasong = isupchebeasong
	oOrderGift.frectdateview = dateview
	oOrderGift.frectdateview1 = dateview1		
	oOrderGift.FRecteventid = evt_code
	oOrderGift.FRectBaljuid = balju_code
	oOrderGift.FRectgift_code = gift_code		
		if yyyy1 <> "" then
			oOrderGift.FRectStartdate = yyyy1 & "-" & mm1 & "-" & dd1
		end if
		if yyyy2 <> "" then
			oOrderGift.FRectEndDate = dateadd("d",1,yyyy2 & "-" & mm2 & "-" & dd2)
		end if
	oOrderGift.GeteventOrderGiftList
end if	
%>

<script language="javascript">
	
function submits()
{
	if (frm.evt_code.value == "" && frm.balju_code.value == "" && frm.gift_code.value == "" ){
		alert('이벤트코드,출고지시코드,사은품코드중 한개를 입력하세요');
		frm.evt_code.focus();
	}else{
		frm.submit();
	}
}

function EnDisabledDateBox(comp){
	document.frm.yyyy1.disabled = !comp.checked;
	document.frm.yyyy2.disabled = !comp.checked;
	document.frm.mm1.disabled = !comp.checked;
	document.frm.mm2.disabled = !comp.checked;
	document.frm.dd1.disabled = !comp.checked;
	document.frm.dd2.disabled = !comp.checked;
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			이벤트코드 : <input type="text" class="text" name="evt_code" value="<%= evt_code %>" size="8" maxlength="8">
        	출고지시코드 : <input type="text" class="text" name="balju_code" value="<%= balju_code %>" size="8" maxlength="8">
        	사은품코드 : <input type="text" class="text" name="gift_code" value="<%= gift_code %>" size="8" maxlength="8">
        	<input type="radio" name="viewType" value="summary" <%= chkIIF (viewType="summary","checked","") %> >합계
        	<input type="radio" name="viewType" value="list" <%= chkIIF (viewType="list","checked","") %> >리스트
        	<input type="radio" name="isupchebeasong" value=""  <%= chkIIF (isupchebeasong="","checked","") %> >전체
        	<input type="radio" name="isupchebeasong" value="N" <%= chkIIF (isupchebeasong="N","checked","") %> >텐배
        	<input type="radio" name="isupchebeasong" value="Y" <%= chkIIF (isupchebeasong="Y","checked","") %> >업배
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button" value="검색" onClick="submits();">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td>
			<input type="radio" name="dateview1" name="dateview1" value="yes"  <% if dateview1="yes" then  response.write "checked" %>>주문일기준(무통장 결제전 포함)
			&nbsp;
			<input type="radio" name="dateview1" name="dateview1" value="yes2"  <% if dateview1="yes2" then  response.write "checked" %>>주문일기준(결제완료)
			&nbsp;
			<input type="radio" name="dateview1" name="dateview1" value="no"  <% if dateview1="no" then  response.write "checked" %>>출고지시일기준
			&nbsp;	
        	<input type=checkbox name="dateview" value="no" <% if dateview="no" then  response.write "checked" %> onclick="EnDisabledDateBox(this)">기간검색
        	<% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		</td>
	</tr>
	
</table>

<p>


<% if viewType = "summary" then %>	<!--구분내역이 합계일 경우에만...-->
	<table width="100%" border="0" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr bgcolor="#FFFFFF">
		<td>&nbsp;날짜표시 X : 
			<input type="checkbox" name="date_display" value="on" <% if date_display = "on" then response.write "checked" %> onclick="frm.submit();">  
			&nbsp;/&nbsp;6개월 이전 : 
			<input type="checkbox" name="chkOldOrder" value="on" <% if chkOldOrder = "on" then response.write "checked" %> onclick="frm.submit();">  
		</td>	
	</tr></form>	
	</table>
<% end if %>
<!-- 표 상단바 끝-->

<%
dim counttotal, ppcnt,ppdate
counttotal = 0
%>
<!-- 합계 시작 -->
<% if viewType = "summary" then %>
	<% if oOrderGiftcount.fresultcount >0 then %>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		    	<% if date_display <> "on" then %>
			        <% if dateview1 = "yes" Or dateview1 = "yes2" then %>
			       		<td width="105">주문일</td>
			   		<% elseif dateview1 = "no" then %>
			   			<td width="50">출고지시일</td>
			   		<% end if %>
		   		<% end if %>
		        <td width="50">Gift ID</td>		   		
				<td width="50">사은품수</td>        
				<td>사은품</td>
				<% if evt_code <> "" then %>
					<td>이벤트코드</td>
				<% end if %>
				<% if balju_code <> "" then %>				
					<td>출고지시코드</td>				
		   		<% end if %>
		        <td>이벤트코드</td>
		        <td>이벤트명</td>		        			
				<td width="50">배송구분</td>
				<td>조건</td>
		    </tr>
		    <% for i = 0 to oOrderGiftcount.fresultcount -1 %>
		    <%		'루프 돌면서 변수에 들어 있는 값이 다음 코드와 틀릴때 까지 다 더한다. 
		    if ppdate <> oOrderGiftcount.FItemList(i).Fgift_code then
				ppdate = oOrderGiftcount.FItemList(i).Fgift_code
				ppcnt = 0
			end if
			ppcnt = ppcnt + oOrderGiftcount.FItemList(i).fgift_code_count		'코드별 합계
		    %>
		    <tr height="20" bgcolor="#FFFFFF" align="center">
		    	<% if date_display <> "on" then %>		    		    	
		       		<td><%= oOrderGiftcount.FItemList(i).Fbaljudate %></td>	  
		        <% end if %> 	        
		        <td><%= oOrderGiftcount.FItemList(i).Fgift_code %></td>		        
		        <td><%= oOrderGiftcount.FItemList(i).fgift_code_count %>
		        	<% counttotal = counttotal + oOrderGiftcount.FItemList(i).fgift_code_count %></td>			
		        <td><%= oOrderGiftcount.FItemList(i).getGiftName %></td>
				<% if evt_code <> "" then %>		        
		        	<td><%= oOrderGiftcount.FItemList(i).Fevt_code %></td>
		        <% end if %>	
				<% if balju_code <> "" then %>	
					<td><%= oOrderGiftcount.FItemList(i).Fbaljuid %></td>		        
		        <% end if %>
		        <td><%= oOrderGiftcount.FItemList(i).Fevt_code %></td>			        
		        <td><%= oOrderGiftcount.FItemList(i).Fevt_name %></td>	        
		        <td>
		        <% if oOrderGiftcount.FItemList(i).Fisupchebeasong="Y" then %>  
			    업체
			    <% else %>
			    텐배
			    <% end if %>  
			    </td>
		        <td><%= oOrderGiftcount.FItemList(i).GetEventConditionStr %></td>
		    </tr>
		    <% if date_display <> "on" then %>
			    <% if i+1<oOrderGiftcount.fresultcount then %> 
					<% if ppdate<>oOrderGiftcount.FItemList(i+1).Fgift_code then %>
					    <tr height="20" bgcolor="#EEEEEE">
						    <td>사은품(<%=ppdate%>)소계</td>
						    <td colspan="8"><%=ppcnt%></td>
					    </tr>						
					<% end if %>
				<% end if %>
			<% end if %>	
			<% next %>
		    <% if date_display <> "on" then %>			
			    <tr height="20" bgcolor="#EEEEEE">
				    <td>사은품(<%=ppdate%>)소계</td>
				    <td colspan="8"><%=ppcnt%></td>
			    </tr>			
			<% end if %>	
			<tr height="20" bgcolor="#FFFFFF">
				<td> 총사은품합계</td>
				<td colspan="8"><%= counttotal %></td>
			</tr>	
		</table>	
	<% else %>
		<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align=center bgcolor="#FFFFFF">검색 결과가 없습니다.</td>
		</tr>
		</table>
	<% end if %>		
<!-- 합계 끝 -->

<!-- 내역 시작 -->
<% elseif (viewType = "list") then %>
		<% if oOrderGift.FResultCount > 0 then %>
		<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
		    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		        <% if dateview1 = "yes" Or dateview1 = "yes2" then %>
		       		<td width="150">주문일</td>
		   		<% elseif dateview1 = "no" then %>
		   			<td width="120">출고지시일</td>
		   		<% end if %>
				<td width="50">출고지시ID</td>
				<td width="100">주문번호</td>
				<td width="50">이벤트ID</td>
				<td>이벤트명</td>
				<td width="50">Gift ID</td>
				<td>사은품</td>
				<td width="50">배송구분</td>
				<td width="100">기간</td>
				<td>조건</td>
			</tr>
			<% for i=0 to oOrderGift.FResultCount -1 %>
			<tr align="center" bgcolor="#FFFFFF">
			    <td><%= oOrderGift.FItemList(i).Fbaljudate %></td>
			    <td><%= oOrderGift.FItemList(i).FBaljuID %></td>
			    <td><%= oOrderGift.FItemList(i).Forderserial %></td>
			    <td><%= oOrderGift.FItemList(i).Fevt_code %></td>
			    <td><%= oOrderGift.FItemList(i).Fevt_name %></td>
			    <td><%= oOrderGift.FItemList(i).Fgift_code %></td>
			    <td><%= oOrderGift.FItemList(i).getGiftName %></td>
			    <td>
			    <% if oOrderGift.FItemList(i).Fisupchebeasong="Y" then %>  
			    업체
			    <% else %>
			    텐배
			    <% end if %>  
			    </td>
			    
			    <td>
			        <%= oOrderGift.FItemList(i).Fevt_startdate %>
			        ~ <br>
			        <%= oOrderGift.FItemList(i).Fevt_enddate %>
			    </td>
			    <td>
			        <%= oOrderGift.FItemList(i).GetEventConditionStr %>
			    </td>
			</tr>
			<% next %>
			<tr bgcolor="#FFFFFF">
			    <td colspan="10" align="center">
			    
			    </td>
			</tr>
		</table>			
	<% else %>
		<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td align=center bgcolor="#FFFFFF">검색 결과가 없습니다.</td>
			</tr>
		</table>
	<% end if %>		
<% end if %>
<!-- 내역 끝 -->


<script language='javascript'>
EnDisabledDateBox(document.frm.dateview);
</script>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->