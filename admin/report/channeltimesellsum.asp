<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->

사용중지
<%

dbget.close()	:	response.End


dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim fromDate,toDate
dim ck_joinmall,ck_ipjummall,ck_pointmall,research,seltime

seltime = request("seltime")
if seltime = "" then seltime="12:00:00"

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

research = request("research")
ck_joinmall = request("ck_joinmall")
ck_ipjummall = request("ck_ipjummall")
ck_pointmall = request("ck_pointmall")

if research<>"on" then
	if ck_joinmall="" then ck_joinmall="on"
	if ck_ipjummall="" then ck_ipjummall="on"
	'if ck_pointmall="" then ck_pointmall="on"
end if

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = "1"

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

dim oreport
set oreport = new CJumunMaster
oreport.FRectFromDate = fromDate
oreport.FRectToDate = toDate
oreport.FRectJoinMallNotInclude = ck_joinmall
oreport.FRectExtMallNotInclude = ck_ipjummall
oreport.FRectPointNotInclude = ck_pointmall
oreport.SearchMallSellTimerePortChannel

dim timereport
set timereport = new CJumunMaster
timereport.FRectFromDate = fromDate
timereport.FRectToDate = toDate
timereport.FRectJoinMallNotInclude = ck_joinmall
timereport.FRectExtMallNotInclude = ck_ipjummall
timereport.FRectPointNotInclude = ck_pointmall
timereport.FRectToDateTime = seltime
timereport.SearchMallSellTimerePortChannel1

dim i,ix,p1,p2
dim p3,p4
%>
<font size="2">흐미 이것이 너무 기간을 많게하면 서버가 힘들어해요~~ --;<br>
한달 이내만 검색해주셈...</font>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		검색시간대 :
		<select name="seltime">
			<option value="">선택</option>
			<option value="09:00:00" <% if seltime="09:00:00" then response.write "selected" %>>오전9</option>
			<option value="10:00:00" <% if seltime="10:00:00" then response.write "selected" %>>오전10</option>
			<option value="11:00:00" <% if seltime="11:00:00" then response.write "selected" %>>오전11</option>
			<option value="12:00:00" <% if seltime="12:00:00" then response.write "selected" %>>오후12</option>
			<option value="13:00:00" <% if seltime="13:00:00" then response.write "selected" %>>오후1</option>
			<option value="14:00:00" <% if seltime="14:00:00" then response.write "selected" %>>오후2</option>
			<option value="15:00:00" <% if seltime="15:00:00" then response.write "selected" %>>오후3</option>
			<option value="16:00:00" <% if seltime="16:00:00" then response.write "selected" %>>오후4</option>
			<option value="17:00:00" <% if seltime="17:00:00" then response.write "selected" %>>오후5</option>
			<option value="18:00:00" <% if seltime="18:00:00" then response.write "selected" %>>오후6</option>
			<option value="19:00:00" <% if seltime="19:00:00" then response.write "selected" %>>오후7</option>
			<option value="20:00:00" <% if seltime="20:00:00" then response.write "selected" %>>오후8</option>
			<option value="21:00:00" <% if seltime="21:00:00" then response.write "selected" %>>오후9</option>
			<option value="22:00:00" <% if seltime="22:00:00" then response.write "selected" %>>오후10</option>
			<option value="23:00:00" <% if seltime="23:00:00" then response.write "selected" %>>오후11</option>
			<option value="24:00:00" <% if seltime="24:00:00" then response.write "selected" %>>오후12</option>
		</select>
		<input type="checkbox" name="ck_joinmall" <% if ck_joinmall="on" then response.write "checked" %> >제휴몰 포함
		<input type="checkbox" name="ck_ipjummall" <% if ck_ipjummall="on" then response.write "checked" %> >입점몰 포함
		<input type="checkbox" name="ck_pointmall" <% if ck_pointmall="on" then response.write "checked" %> >포인트몰 포함
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00">
        <tr align="center">
          <td width="120" class="a"><font color="#FFFFFF">기간</font></td>
          <td class="a" width="600"><font color="#FFFFFF"></font></td>
          <td class="a" width="120"><font color="#FFFFFF">내용</font></td>
        </tr>

		<% for i=0 to oreport.FResultCount-1 %>
		<%
			if oreport.maxt<>0 then
				p1 = Clng(oreport.FMasterItemList(i).Fsellcnt/oreport.maxc*100)
				p2 = Clng(timereport.FMasterItemList(i).Fsellcnt/oreport.FMasterItemList(i).Fsellcnt*100)
			end if
		%>
<%
dim arr,add_nal
   arr = split(oreport.FMasterItemList(i).Fsitename, "-")
   add_nal = cstr(DateSerial(arr(0), arr(1), arr(2) + 1))

dim CHtimereport
set CHtimereport = new CJumunMaster
CHtimereport.FRectFromDate = oreport.FMasterItemList(i).Fsitename
CHtimereport.FRectToDate = add_nal
CHtimereport.FRectJoinMallNotInclude = ck_joinmall
CHtimereport.FRectExtMallNotInclude = ck_ipjummall
CHtimereport.FRectPointNotInclude = ck_pointmall
CHtimereport.FRectToDateTime = seltime
CHtimereport.SearchMallSellTimerePortChannel2
%>
        <tr bgcolor="#FFFFFF" height="10"  class="a">
		  <td width="120" height="10">
          	<%= oreport.FMasterItemList(i).Fsitename %>(<%= oreport.FMasterItemList(i).GetDpartName %>)
          </td>
          <td  height="10" width="600">
			<br><div align="left"> <img src="/images/dot2.gif" height="4" width="<%= p1 %>%"></div><br>
			<div align="left"> <img src="/images/dot10.gif" height="4" width="<%= p2 %>%"></div><br>
				<% for ix=0 to CHtimereport.FResultCount-1 %>
					<div align="left"> <img src="/images/dot<% = CStr(CLng(CHtimereport.FMasterItemList(ix).Fitemgubun)) %>.gif" height="4" width="<%= Clng(CHtimereport.FMasterItemList(ix).Fsellcnt/oreport.FMasterItemList(i).Fsellcnt*100) %>%"></div><br>
				<% next %>
          </td>
		  <td class="a" width="160" align="right">
		   <%= FormatNumber(oreport.FMasterItemList(i).Fselltotal,0) %>원 (<%= FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) %>건)&nbsp;<font color="#808080">총매출</font> <br>
		   <%= FormatNumber(timereport.FMasterItemList(i).Fselltotal,0) %>원 (<%= FormatNumber(timereport.FMasterItemList(i).Fsellcnt,0) %>건)&nbsp;<font color="#808080">시간대</font><br>
				<% for ix=0 to CHtimereport.FResultCount-1 %>
		   <%= FormatNumber(CHtimereport.FMasterItemList(ix).Fselltotal,0) %>원 (<%= FormatNumber(CHtimereport.FMasterItemList(ix).Fsellcnt,0) %>건)&nbsp;<font color="#808080"><% = CHtimereport.FMasterItemList(ix).GetChannelName_Kor %></font><br>
				<% next %>
		  </td>
        </tr>
        <% next %>
</table>
<%
set oreport = Nothing
set timereport = Nothing
set CHtimereport = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->