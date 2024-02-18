<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->

<%
dim page,shopid,i,ix
dim yyyy1,mm1,dd1
dim fromDate,toDate
dim oldlist

oldlist = request("oldlist")

page = request("page")
if page="" then page=1

yyyy1 = request("yyyy1")
mm1 = request("mm1")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = "1"


fromDate = left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
toDate = Left(CStr(DateSerial(yyyy1,mm1+1,dd1)),10)

dim oreport
set oreport = new CJumunMaster
oreport.FRectFromDate = fromDate
oreport.FRectToDate = toDate
oreport.FRectOldJumun = oldlist

oreport.GetWeeklySellReport


dim selltotal
dim selltotal_jj, sellcnt_jj, DpartCount_jj
dim selltotal_jm, sellcnt_jm, DpartCount_jm

dim avgsell,avgselltotal

selltotal =0

for i=0 to oreport.FResultCount -1
	selltotal 	= selltotal + oreport.FMasterItemList(i).Fselltotal
	if oreport.FMasterItemList(i).FDpartCount<>0 then
		avgsell		= CLng(oreport.FMasterItemList(i).Fselltotal/oreport.FMasterItemList(i).FDpartCount)
		avgselltotal = avgselltotal + avgsell
	end if

	if oreport.FMasterItemList(i).Fdpart="1" or oreport.FMasterItemList(i).Fdpart="7" then
 		selltotal_jm	= selltotal_jm + oreport.FMasterItemList(i).Fselltotal
 		sellcnt_jm		= sellcnt_jm + oreport.FMasterItemList(i).Fsellcnt
 		DpartCount_jm	= DpartCount_jm + oreport.FMasterItemList(i).FDpartCount
 	else
 		selltotal_jj	= selltotal_jj + oreport.FMasterItemList(i).Fselltotal
 		sellcnt_jj		= sellcnt_jj + oreport.FMasterItemList(i).Fsellcnt
 		DpartCount_jj	= DpartCount_jj + oreport.FMasterItemList(i).FDpartCount
 	end if

next




%>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
			매출월 : <% DrawYMBox yyyy1,mm1 %>&nbsp;&nbsp;		<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역조회
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>

<table width="100%" border="0" cellpadding="4" cellspacing="1" class="a" bgcolor="#3d3d3d">
<tr bgcolor="#D1D1D1" align="center">
	<td>요일</td>
	<td>총매출</td>
	<td>총구매건수</td>
	<td>일수</td>
	<td>평균매출</td>
	<td>평균구매건수</td>
	<td>평균객단가</td>
	<td>평균매출점유율</td>
</tr>
<% for i=0 to oreport.FResultCount -1 %>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= oreport.FMasterItemList(i).GetDpartName %></td>
	<td align="right" ><%= FormatNumber(oreport.FMasterItemList(i).Fselltotal,0) %></td>
	<td align="center"><%= FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) %></td>
	<td align="center"><%= oreport.FMasterItemList(i).FDpartCount %></td>
	<td align="right">
		<% if oreport.FMasterItemList(i).FDpartCount<>0 then %>
			<% avgsell = CLng(oreport.FMasterItemList(i).Fselltotal/oreport.FMasterItemList(i).FDpartCount) %>
			<%= FormatNumber(avgsell,0) %>
		<% end if %>
	</td>
	<td align="center">
		<% if oreport.FMasterItemList(i).FDpartCount<>0 then %>
			<%= FormatNumber(CLng(oreport.FMasterItemList(i).Fsellcnt/oreport.FMasterItemList(i).FDpartCount),0) %>
		<% end if %>
	</td>
	<td align="right">
		<% if oreport.FMasterItemList(i).Fsellcnt<>0 then %>
			<%= FormatNumber(CLng(oreport.FMasterItemList(i).Fselltotal/oreport.FMasterItemList(i).Fsellcnt),0) %>
		<% end if %>
	</td>
	<td align="center">
		<% if avgselltotal<>0 then %>
			<%= CLng(avgsell/avgselltotal*100*100)/100 %> %
		<% end if  %>
	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" height="20">
	<td colspan="8"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">주중</td>
	<td align="right" ><%= FormatNumber(selltotal_jj,0) %></td>
	<td align="center"><%= FormatNumber(sellcnt_jj,0) %></td>
	<td align="center"><%= DpartCount_jj %></td>
	<td align="right">
		<% if DpartCount_jj<>0 then %>
			<%= FormatNumber(CLng(selltotal_jj/DpartCount_jj),0) %>
		<% end if %>
	</td>
	<td align="center">
		<% if DpartCount_jj<>0 then %>
			<%= FormatNumber(CLng(sellcnt_jj/DpartCount_jj),0) %>
		<% end if %>
	</td>
	<td align="right">
		<% if sellcnt_jj<>0 then %>
			<%= FormatNumber(CLng(selltotal_jj/sellcnt_jj),0) %>
		<% end if %>
	</td>
	<td align="center">
		<% if selltotal<>0 then %>
			<%= CLng(selltotal_jj/selltotal*100*100)/100 %> %
		<% end if  %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td align="center">주말</td>
	<td align="right" ><%= FormatNumber(selltotal_jm,0) %></td>
	<td align="center"><%= FormatNumber(sellcnt_jm,0) %></td>
	<td align="center"><%= DpartCount_jm %></td>
	<td align="right">
		<% if DpartCount_jm<>0 then %>
			<%= FormatNumber(CLng(selltotal_jm/DpartCount_jm),0) %>
		<% end if %>
	</td>
	<td align="center">
		<% if DpartCount_jm<>0 then %>
			<%= FormatNumber(CLng(sellcnt_jm/DpartCount_jm),0) %>
		<% end if %>
	</td>
	<td align="right">
		<% if sellcnt_jm<>0 then %>
			<%= FormatNumber(CLng(selltotal_jm/sellcnt_jm),0) %>
		<% end if %>
	</td>
	<td align="center">
		<% if selltotal<>0 then %>
			<%= CLng(selltotal_jm/selltotal*100*100)/100 %> %
		<% end if  %>
	</td>
</tr>
</table>
<br><br>


<%
set oreport= Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->