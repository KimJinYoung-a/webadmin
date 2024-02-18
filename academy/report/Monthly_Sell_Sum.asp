<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/academy_reportcls.asp"-->
<%
dim ck_joinmall,ck_ipjummall,ck_pointmall
dim research
dim opt_rect
dim yyyy1,mm1, accountdiv

ck_joinmall = RequestCheckvar(request("ck_joinmall"),2)
ck_ipjummall = RequestCheckvar(request("ck_ipjummall"),2)
ck_pointmall = RequestCheckvar(request("ck_pointmall"),2)
research = RequestCheckvar(request("research"),2)
yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1	  = RequestCheckvar(request("mm1"),2)
accountdiv =RequestCheckvar(request("accountdiv"),3)
opt_rect = "all"

if research<>"on" then
	if ck_joinmall="" then ck_joinmall="on"
	if ck_ipjummall="" then ck_ipjummall="on"
end if

dim stdate
if yyyy1="" then
	stdate = CStr(Now)
	stdate = DateSerial(Left(stdate,4), CLng(Mid(stdate,6,2))-1,1)
	yyyy1 = Left(stdate,4)
	mm1 = Mid(stdate,6,2)
end if
dim oreport
set oreport = new CJumunMaster
oreport.Faccountdiv=accountdiv
oreport.FRectFromDate = yyyy1 + "-" + mm1 + "-01"

oreport.SearchMallSellrePort4


dim i,p1,p2,p3,p4
dim maybe_monthcount
dim maybe_monthsum
dim dayno, currno

dim nowdate,nowyyyymm

if opt_rect="all" then
	nowdate = CStr(date)
	nowyyyymm = left(nowdate,7)
	currno = CInt(right(nowdate,2))

	nowdate = dateserial(Left(nowdate,4),Mid(nowdate,6,2)+1,0)
	dayno = CInt(right(nowdate,2))
end if

%>


<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
			검색기간 :
			<% DrawYMBox yyyy1,mm1 %> ~ 현재 (주문일)&nbsp;&nbsp;
			<input type="radio" name="accountdiv" value="" <% if accountdiv="" then response.write "checked" %>>전체
			<input type="radio" name="accountdiv" value="100" <% if accountdiv="100" then response.write "checked" %>>신용카드
			<input type="radio" name="accountdiv" value="7" <% if accountdiv="7" then response.write "checked" %>>무통장
			<input type="radio" name="accountdiv" value="20" <% if accountdiv="20" then response.write "checked" %>>실시간이체
			<input type="radio" name="accountdiv" value="80" <% if accountdiv="80" then response.write "checked" %>>ALL@
			<input type="radio" name="accountdiv" value="900" <% if accountdiv="900" then response.write "checked" %>>수기입력
			<br>

		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00">
        <tr align="center">
          <td width="120" class="a"><font color="#FFFFFF">기간</font></td>
          <td class="a"><font color="#FFFFFF"></font></td>
          <td class="a" width="100"><font color="#FFFFFF">금액(원)</font></td>
          <td class="a" width="50"><font color="#FFFFFF">건수</font></td>
        </tr>

		<% for i=0 to oreport.FResultCount-1 %>
		<%
			if Left(oreport.FMasterItemList(i).Fsitename,7)=nowyyyymm then

				maybe_monthsum	 = CLng(oreport.FMasterItemList(i).Fselltotal * dayno / currno)
				maybe_monthcount = CLng(oreport.FMasterItemList(i).Fsellcnt * dayno / currno)


				if maybe_monthcount>oreport.maxc then
					oreport.maxc = maybe_monthcount
				end if

				if maybe_monthsum>oreport.maxt then
					oreport.maxt = maybe_monthsum
				end if

				p3 = Clng(maybe_monthsum/oreport.maxt*100)
				p4 = Clng(maybe_monthcount/oreport.maxc*100)
			end if

			if oreport.maxt<>0 then
				p1 = Clng(oreport.FMasterItemList(i).Fselltotal/oreport.maxt*100)
			end if

			if oreport.maxc<>0 then
				p2 = Clng(oreport.FMasterItemList(i).Fsellcnt/oreport.maxc*100)
			end if
		%>
				<tr bgcolor="#FFFFFF" height="35" class="a">
					<td width="120">
						<%= oreport.FMasterItemList(i).Fsitename %>
          </td>
          <td>
          	<% if Left(oreport.FMasterItemList(i).Fsitename,7)=nowyyyymm then %>
          	<div align="left"> <img src="/images/dot4.gif" height="3" width="<%= p3 %>%"></div><br>
          	<div align="left"> <img src="/images/dot4.gif" height="3" width="<%= p4 %>%"></div><br>
          	<% end if %>
						<div align="left"> <img src="/images/dot1.gif" height="3" width="<%= p1 %>%"></div><br>
          	<div align="left"> <img src="/images/dot2.gif" height="3" width="<%= p2 %>%"></div>
         	</td>
					<td class="a" width="100" align="right">
		  			<% if Left(oreport.FMasterItemList(i).Fsitename,7)=nowyyyymm then %>
		  				<font color="#AAAAAA"><%= FormatNumber(maybe_monthsum,0) %></font><br>
		  			<% end if %>
		   			<%= FormatNumber(oreport.FMasterItemList(i).Fselltotal,0) %><br>
          </td>
					<td class="a" width="50" align="right">
		  			<% if Left(oreport.FMasterItemList(i).Fsitename,7)=nowyyyymm then %>
				  	<font color="#AAAAAA"><%= FormatNumber(maybe_monthcount,0) %></font><br>
				  	<% end if %>
				   	<%= FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) %>
					</td>

        </tr>
        <% next %>
</table>
<%
set oreport = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->