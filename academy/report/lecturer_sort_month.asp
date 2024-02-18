<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/academy_reportcls.asp"-->


<%
dim oreport
dim stdate
dim yyyy1,mm1
Dim sort

yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1	  = RequestCheckvar(request("mm1"),2)
sort	  = RequestCheckvar(request("sort"),6)

if yyyy1="" then
	stdate = CStr(Now)
	yyyy1 = Left(stdate,4)
	mm1 = Mid(stdate,6,2)
end if

set oreport = new CJumunMaster
oreport.FRectFromDate = yyyy1 + "-" + mm1
oreport.FRectSort = sort
oreport.GetLecturerMonthMeaChul

Dim i,p1,p2
Dim premonth_sellsum,premonth_sellcnt

dim selltotal, sellcnt
%>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a">
			검색기간 : <% DrawYMBox yyyy1,mm1 %>&nbsp;&nbsp;&nbsp;
			<select name="sort">
				<option value="tsum" <% If sort="tsum" Then response.write "selected" %>>고액순</option>
				<option value="tcnt" <% If sort="tcnt" Then response.write "selected" %>>건수순</option>
				<option value="name" <% If sort="name" Then response.write "selected" %>>강사명순</option>
			</select>
		</td>
		<td class="a" align="left">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0" align="sbsmiddle"></a>
		</td>
		<td class="a" align="right">
			<input type="button" onclick="downloadexcel();" value="엑셀다운로드" class="button">
		</td>
	</tr>
	</form>
</table>
<table border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00" class="a">
        <tr align="center">
          <td width="200" class="a"><font color="#FFFFFF">강사</font></td>
          <td class="a" width="600"><font color="#FFFFFF"></font></td>
          <td class="a" width="100"><font color="#FFFFFF">액수</font></td>
          <td class="a" width="50"><font color="#FFFFFF">건수</font></td>
          <td class="a" width="80"><font color="#FFFFFF">객단가</font></td>
        </tr>

		<% for i=0 to oreport.FResultCount-1 %>
		<%
			if oreport.maxt<>0 then
				p1 = Clng(oreport.FMasterItemList(i).Fselltotal/oreport.maxt*100)
			end if

			if oreport.maxc<>0 then
				p2 = Clng(oreport.FMasterItemList(i).Fsellcnt/oreport.maxc*100)
			end if

			selltotal = selltotal + oreport.FMasterItemList(i).Fselltotal
			sellcnt = sellcnt + oreport.FMasterItemList(i).Fsellcnt

		%>
        <tr bgcolor="#FFFFFF" height="10"  >
				<td width="200" height="10">
					<font color="#808080"><%= oreport.FMasterItemList(i).Fsitename %></font> (<font color="#0080C0"><%= oreport.FMasterItemList(i).Flecturer %></font>)
				</td>
				<td  height="10" width="600">
					 <div align="left"> <img src="/images/dot1.gif" height="4" width="<%= p1 %>%"></div><br>
					 <div align="left"> <img src="/images/dot2.gif" height="4" width="<%= p2 %>%"></div>
				</td>
				<td class="a" width="100" align="right">
					 <%= FormatNumber(oreport.FMasterItemList(i).Fselltotal,0) %>원
				</td>
				<td class="a" width="80" align="right">
					 <%= FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) %>건
				</td>
				<td class="a" width="80" align="right">
				<% if oreport.FMasterItemList(i).Fsellcnt<>0 then %>
					 <%= FormatNumber(Clng(oreport.FMasterItemList(i).Fselltotal/oreport.FMasterItemList(i).Fsellcnt),0) %>원
				<% end if %>
				</td>
        </tr>
        <% next %>
        <tr bgcolor="#FFFFFF">
        	<td>Total</td>
        	<td></td>
        	<td align="right"> <%= FormatNumber(selltotal,0) %>원 </td>
        	<td align="right"> <%= FormatNumber(sellcnt,0) %>건 </td>
        	<td align="right">
        	<% if sellcnt<>0 then %>
        		<%= FormatNumber(selltotal/sellcnt,0) %>원
        	<% end if %>
        	</td>
        </tr>
</table>
<%
set oreport = Nothing
%>
<script type="text/javascript">
<!--
function downloadexcel(){
    document.frm.target = "view"; 
    document.frm.action = "/academy/report/statistic_lecturer_month_excel.asp";  
	document.frm.submit();
    document.frm.target = ""; 
    document.frm.action = "";  
}
//-->
</script>
<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->