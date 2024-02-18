<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim research, segumtype
dim thismonth

research = request("research")
segumtype = request("segumtype")


thismonth = Left(CStr(DateSerial(year(now()),month(now())-1,1)),7)
%>

<script language='javascript'>

function getCSV(searchtype){
    location.href = '/admin/upchejungsan/pop_jungsan_upload_csv.asp?searchtype=' + searchtype;
}

function getExcel(searchtype){
    location.href = '/admin/upchejungsan/pop_jungsan_upload_excel.asp?searchtype=' + searchtype;
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	<img src="/images/icon_star.gif" align="absbottom"> <strong>정산내역 업로드리스트</strong>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<%

dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FRectNotIncludeWonChon = "on"
ojungsan.FRectYYYYMM = thismonth
ojungsan.FRectbankingupflag = "Y"

ojungsan.JungsanFixedList

dim ipsum,i
ipsum =0
%>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
    	<td colspan="5" >금월(<%= thismonth %>) 세금계산서 (<%= ojungsan.FresultCount %>건)</td>
    	<td align=right>
    	  <a href="javascript:getExcel('thismonth')"><img src="/images/iexcel.gif" border=0></a>
    	  <a href="javascript:getCSV('thismonth')"><img src="/images/icon_arrow_link.gif" border=0></a>
    	</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="60">은행</td>
        <td>계좌</td>
        <td width="80">정산금액</td>
        <td width="120">업체명</td>
        <td>사업자등록번호</td>
        <td width="120">(주)텐바이텐</td>
    </tr>
<% for i=0 to ojungsan.FresultCount-1 %>
<%
ipsum = ipsum + ojungsan.FItemList(i).GetTotalSuplycash
%>

	<% if ojungsan.FItemList(i).GetTotalSuplycash<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
        <% if ojungsan.FItemList(i).Fipkum_bank = "홍콩샹하이" then %>
		<td>HSBC</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "단위농협" then %>
		<td>농협</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "제일" then %>
		<td>SC제일</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "시티" then %>
		<td>한국씨티</td>
		<% else %>
		<td><%= ojungsan.FItemList(i).Fipkum_bank %></td>
		<% end if %>

        <td><%= ojungsan.FItemList(i).Fipkum_acctno %></td>
        <td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalSuplycash,0) %></td>
        <td><%= Left(ojungsan.FItemList(i).Fcompany_name,9) %></td>
        <td><%=ojungsan.FItemList(i).Fcompany_no%></td>
        <td>(주)텐바이텐</td>
	</tr>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="2"></td>
		<td align="right"><%= FormatNumber(ipsum,0) %></td>
		<td colspan="3"></td>
	</tr>
</table>

<%
ojungsan.FRectYYYYMM = ""
ojungsan.FRectNotIncludeWonChon = "on"
ojungsan.FRectNotYYYYMM = thismonth
ojungsan.FRectbankingupflag = "Y"

ojungsan.JungsanFixedList

ipsum =0
%>

<br>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
    	<td colspan="5" >전월 세금계산서 (<%= ojungsan.FresultCount %>건)</td>
    	<td align=right>
    	  <a href="javascript:getExcel('prevmonth')"><img src="/images/iexcel.gif" border=0></a>
    	  <a href="javascript:getCSV('prevmonth')"><img src="/images/icon_arrow_link.gif" border=0></a>
    	</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="60">은행</td>
      <td>계좌</td>
      <td width="80">정산금액</td>
      <td width="120">업체명</td>
      <td>사업자등록번호</td>
      <td width="120">(주)텐바이텐</td>
     </tr>
<% for i=0 to ojungsan.FresultCount-1 %>
<%
ipsum = ipsum + ojungsan.FItemList(i).GetTotalSuplycash
%>

	<% if ojungsan.FItemList(i).GetTotalSuplycash<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
    	<% if ojungsan.FItemList(i).Fipkum_bank = "홍콩샹하이" then %>
		<td>HSBC</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "단위농협" then %>
		<td>농협</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "제일" then %>
		<td>SC제일</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "시티" then %>
		<td>한국씨티</td>
		<% else %>
		<td><%= ojungsan.FItemList(i).Fipkum_bank %></td>
		<% end if %>

        <td><%= ojungsan.FItemList(i).Fipkum_acctno %></td>
        <td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalSuplycash,0) %></td>
        <td><%= Left(ojungsan.FItemList(i).Fcompany_name,9) %></td>
        <td><%=ojungsan.FItemList(i).Fcompany_no%></td>
        <td>(주)텐바이텐</td>
	</tr>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="2"></td>
		<td align="right"><%= FormatNumber(ipsum,0) %></td>
		<td colspan="3"></td>
	</tr>
</table>

<%
ojungsan.FRectYYYYMM = ""
ojungsan.FRectNotYYYYMM = ""
ojungsan.FRectNotIncludeWonChon = ""
ojungsan.FRectOnlyIncludeWonChon = "on"
ojungsan.FRectbankingupflag = "Y"

ojungsan.JungsanFixedList

ipsum =0
%>
<br>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
    	<td colspan="5" >원천징수 대상자 (<%= ojungsan.FresultCount %>건)</td>
    	<td align=right>
    	  <a href="javascript:getExcel('withholding')"><img src="/images/iexcel.gif" border=0></a>
    	  <a href="javascript:getCSV('withholding')"><img src="/images/icon_arrow_link.gif" border=0></a>
    	</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="60">은행</td>
      <td>계좌</td>
      <td width="80">정산금액</td>
      <td width="120">업체명</td>
      <td>사업자등록번호</td>
      <td width="120">(주)텐바이텐</td>
    </tr>
<% for i=0 to ojungsan.FresultCount-1 %>
<%
ipsum = ipsum + ojungsan.FItemList(i).GetTotalWithHoldingJungSanSum
%>

	<% if ojungsan.FItemList(i).GetTotalSuplycash<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
        <% if ojungsan.FItemList(i).Fipkum_bank = "홍콩샹하이" then %>
		<td>HSBC</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "단위농협" then %>
		<td>농협</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "제일" then %>
		<td>SC제일</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "시티" then %>
		<td>한국씨티</td>
		<% else %>
		<td><%= ojungsan.FItemList(i).Fipkum_bank %></td>
		<% end if %>

        <td><%= ojungsan.FItemList(i).Fipkum_acctno %></td>
        <td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalWithHoldingJungSanSum,0) %></td>
        <td><%= Left(ojungsan.FItemList(i).Fcompany_name,9) %></td>
        <td><%=ojungsan.FItemList(i).Fcompany_no%></td>
        <td>(주)텐바이텐</td>
	</tr>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="2"></td>
		<td align="right"><%= FormatNumber(ipsum,0) %></td>
		<td colspan="3"></td>
	</tr>
</table>

<%
set ojungsan = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->