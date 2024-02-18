<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/company/lib/companybodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/partner_reportcls.asp"-->
<%
const Maxlines = 10
dim totalpage, totalnum, q, i


dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim gotopage,ojumun
dim fromDate,toDate,jnx,tmpStr,siteId
dim showtype, IsAdmin, settle2
dim tinginclude
dim searchId,ckipkumdiv4

showtype = request("showtype")
settle2 = request("settle2")
ckipkumdiv4 = request("ckipkumdiv4")
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

'response.write session("ssBctID") & "<br>"
'response.write settle2 & "<br>"
'response.write showtype & "<br>"
'response.write settle2

''서동팔 수정..
''기본값적용..
If gotopage <> "" then
   session("gotopage") = CInt(gotopage)
else
   Session("gotopage") = 1
   gotopage = session("gotopage")
end if

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

if session("ssBctDiv")="999" then
	ojumun.FRectRdSite = session("ssBctID")
else
	ojumun.FRectDesignerID = session("ssBctID")
end if


ojumun.FRectFromDate = fromDate
ojumun.FRectToDate = toDate
ojumun.FRectSettle2 = settle2
ojumun.FRectIpkumDiv4 = "on"
ojumun.SearchSellrePort

%>
<script language='javascript'>
function Check(){
   if(document.frm.elements[1].checked == true){
       document.frm.ckipkumdiv4.value="on";
   }
   document.frm.submit();
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="sellreport.asp">
      <input type="hidden" name="showtype" value="<%= showtype %>">
      <input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>

          	옵션:
            <input type="radio" name="settle2" value="m" <% if (settle2="m") then response.write("checked") %> >월별
            <input type="radio" name="settle2" value="d" <% if (settle2="d") then response.write("checked") %> >일별
		<td class="a" align="right">
			<a href="javascript:Check();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00">
        <tr align="center">
          <td width="120" class="a"><font color="#FFFFFF">구분</font></td>
          <td class="a" width="600"><font color="#FFFFFF"></font></td>
          <td class="a" width="120"><font color="#FFFFFF">내용</font></td>
        </tr>

        <% for i=0 to ojumun.FResultCount - 1 %>
        <tr bgcolor="#FFFFFF" height="10">
          <td width="120" height="10" class="a">
          	<%= ojumun.FMasterItemList(i).Fseldate %>
          </td>
          <td  height="10" width="600">
          <% if Not (IsNull(ojumun.FMasterItemList(i).Fselltotal)) then %>
          	<div align="left"> <img src="/images/dot1.gif" height="4" width="<%= CLng((ojumun.FMasterItemList(i).Fselltotal/ojumun.maxt)*600) %>"></div><br>
          	<div align="left"> <img src="/images/dot2.gif" height="4" width="<%= CLng((ojumun.FMasterItemList(i).Fsellcnt/ojumun.maxc)*600) %>"></div>
          <% end if %>
          </td>
		  <td class="a" width="120">
		  <% if Not (IsNull(ojumun.FMasterItemList(i).Fselltotal)) then %>
		  	<div align="right"> <%= FormatNumber(FormatCurrency(ojumun.FMasterItemList(i).Fselltotal),0) & "원"%> </div>
		  	<div align="right"> <%= ojumun.FMasterItemList(i).Fsellcnt & "건"%> </div>
		  <% end if %>
		  </td>
        </tr>
        <% next %>
</table>
<%
set ojumun = Nothing
%>
<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

</body>
</html>
