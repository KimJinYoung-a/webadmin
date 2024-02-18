<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/company/lib/companybodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/designer_reportcls.asp"-->
<%
const Maxlines = 10
dim totalpage, totalnum, q, i


dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim gotopage,ojumun
dim fromDate,toDate,jnx,tmpStr,siteId, ttCnt, ttPrice
dim showtype, IsAdmin, settle2
dim oldlist

showtype = request("showtype")
settle2 = request("settle2")
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
oldlist = request("oldlist")

'response.write session("ssBctID") & "<br>"
'response.write settle2 & "<br>"
'response.write showtype & "<br>"
'response.write settle2

''기본값적용..
ttCnt = 0
ttPrice = 0

if (settle2="") then settle2= "d"

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))


	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)

set ojumun = new CJumunMaster

ojumun.FRectDesignerID = "haepumdal"
ojumun.FRectFromDate = fromDate
ojumun.FRectToDate = toDate
ojumun.FRectSettle2 = settle2
ojumun.FRectOldJumun = oldlist
ojumun.SearchSellrePort_HPDCase

%>
<script language='javascript'>
function Check(){
   if(document.frm.elements[2].checked == true){
       document.frm.ckipkumdiv4.value="on";
   }
   document.frm.submit();
}
</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<form name="frm" method="get" action="sellreportHPD.asp">
    <input type="hidden" name="showtype" value="<%= showtype %>">
    <input type="hidden" name="menupos" value="<%= request("menupos") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
            <!--
        	<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역
        	-->
			&nbsp;&nbsp;
			검색기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
          	&nbsp;&nbsp;
            <input type="radio" name="settle2" value="m" <% if (settle2="m") then response.write("checked") %> >월별
            <input type="radio" name="settle2" value="d" <% if (settle2="d") then response.write("checked") %> >일별
        </td>
        <td align="right">
        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80">구분</td>
		<td></td>
		<td width="60">주문건수</td>
		<td width="80">판매가합계</td>
	</tr>
		
	<% for i=0 to ojumun.FResultCount - 1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= ojumun.FMasterItemList(i).Fseldate %></td>
		<td>
			<% if Not (IsNull(ojumun.FMasterItemList(i).Fselltotal)) then %>
			<div align="left" title="금액"><img src="/images/dot1.gif" height="3" width="<%= CLng((ojumun.FMasterItemList(i).Fselltotal/ojumun.maxt)*600) %>"></div><br>
			<div align="left" title="건수"><img src="/images/dot2.gif" height="3" width="<%= CLng((ojumun.FMasterItemList(i).Fsellcnt/ojumun.maxc)*600) %>"></div>
			<% end if %>
		</td>
		<td>
			<% if Not (IsNull(ojumun.FMasterItemList(i).Fselltotal)) then %>
			<%= ojumun.FMasterItemList(i).Fsellcnt %>
			<% end if %>
		</td>
		<td align="right">
			<% if Not (IsNull(ojumun.FMasterItemList(i).Fselltotal)) then %>
			<%= FormatNumber(FormatCurrency(ojumun.FMasterItemList(i).Fselltotal),0) %>
			<% end if %>
		</td>
	</tr>
	<%
			'총 합계 계산
			if Not (IsNull(ojumun.FMasterItemList(i).Fselltotal)) then
				ttCnt = ttCnt + ojumun.FMasterItemList(i).Fsellcnt
				ttPrice = ttPrice + ojumun.FMasterItemList(i).Fselltotal
			end if
		next
		if ttPrice>0 then
			Response.Write "<tr align=center bgcolor=#F8F8F8>" &_
							"<td colspan=2><b>총 계</b></td>" &_
							"<td><b>" & ttCnt & "</b></td>" &_
							"<td align=right><b>" & FormatNumber(ttPrice,0) & "</b></td>" &_
							"</tr>"
		end if
	%>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<%
set ojumun = Nothing
%>
<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

</body>
</html>
