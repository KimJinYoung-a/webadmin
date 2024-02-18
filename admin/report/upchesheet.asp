<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/upchereportcls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim fromDate,toDate
dim designer

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
designer = request("designer")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = "1"

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

dim i,p1,p2

dim oreport
set oreport = new CUpcheReport
oreport.FRectFromDate = CStr(fromDate)
oreport.FRectToDate = CStr(toDate)
oreport.FRectDesigner = designer

if designer<>"" then
	oreport.GetUpcheSheet1
	oreport.GetUpcheSheet2
	oreport.GetEventItem
	oreport.GetUpcheBestItem
	oreport.GetUpcheSheet3
	'oreport.GetUpcheAllMeaChul
end if

dim premaechul, prerank
premaechul =0
prerank =0


%>
<script language='javascript'>
function showReBuy(){

}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
		��ü : <% drawSelectBoxDesigner "designer",designer %>
		&nbsp;&nbsp;
		�˻��Ⱓ :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="1200" border="0" cellspacing="1" cellpadding="3" bgcolor=#3d3d3d class="a">
<tr bgcolor="#FFFFFF">
	<td colspan="16">**������ 1�� �̳��� �˻��Ͻñ� �����մϴ�..</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="80" align="center">�Ⱓ</td>
	<td width="30" align="center">�ֹ�<br>�Ǽ�</td>
	<td width="60" align="center">�����</td>
	<td width="60" align="center">���ܰ�</td>
	<td width="60" align="center">�̺�Ʈ����</td>
	<td width="60" align="center">����Ʈ��ǰ1</td>
	<td width="60" align="center">����Ʈ��ǰ2</td>
	<td width="60" align="center">����Ʈ��ǰ3</td>
	<td width="60" align="center">���⼺����</td>
	<td width="60" align="center">���ż���(��/��)(%)</td>
	<td width="60" align="center">Ÿ�ٿ���(��/��)(��)</td>
	<td width="60" align="center">��ü�����</td>
	<td width="60" align="center">������</td>
	<td width="60" align="center">����</td>
	<td width="60" align="center">��������</td>
	<!-- <td width="60" align="center">�籸����</td> -->
</tr>
<% for i=0 to oreport.FResultCount-1 %>

<tr bgcolor="#FFFFFF">
	<td align="center"><%= oreport.FItemList(i).FDateGubun %></td>
	<td align="right"><%= FormatNumber(oreport.FItemList(i).FSubCount,0) %></td>
	<td align="right"><%= FormatNumber(oreport.FItemList(i).FSubSellTotal,0) %></td>
	<td align="right"><%= FormatNumber(CLng(oreport.FItemList(i).FSubSellTotal/oreport.FItemList(i).FSubCount),0) %></td>
	<td align="center">
	<% if oreport.FItemList(i).FEventNo>0 then %>
	o
	<% else %>
	-
	<% end if %>
	</td>
	<td align="center"><%= oreport.FItemList(i).FBestItem1 %></td>
	<td align="center"><%= oreport.FItemList(i).FBestItem2 %></td>
	<td align="center"><%= oreport.FItemList(i).FBestItem3 %></td>
	<td align="center">
	<% if premaechul=0 then %>
	-
	<% else %>
	<%= FormatNumber(CLng((oreport.FItemList(i).FSubSellTotal-premaechul)/premaechul*100),0) %> %
	<% end if %>
	</td>
	<td align="center">
	<% if oreport.FItemList(i).FMWTotal=0 then %>
	-
	<% else %>
	<%= CLng(oreport.FItemList(i).FManCount/oreport.FItemList(i).FMWTotal*100) %>/<%= CLng(oreport.FItemList(i).FWoManCount/oreport.FItemList(i).FMWTotal*100) %>
	<% end if %>
	</td>
	<td align="center">
	<%= oreport.FItemList(i).getManTargetNai %>/<%= oreport.FItemList(i).getWoManTargetNai %>
	</td>
	<td align="center"><%= FormatNumber(oreport.FItemList(i).FTotalSum,0) %></td>
	<td align="center">
	<% if oreport.FItemList(i).FTotalSum=0 then %>
	-
	<% else %>
	<%= FormatNumber(CLng(oreport.FItemList(i).FSubSellTotal/oreport.FItemList(i).FTotalSum*100),0) %> %
	<% end if %>
	</td>
	<td align="center"><%= oreport.FItemList(i).FRank %>/<%= oreport.FItemList(i).FTotalUpCheCount %></td>
	<td align="center">
	<% if prerank=0 then %>
	-
	<% else %>
	<%= prerank - oreport.FItemList(i).FRank %>
	<% end if %>
	</td>
	<!--
	<td align="center"><a href="#" onclick="showReBuy('<%= oreport.FItemList(i).FDateGubun %>','<%= designer %>')">-&gt;</a></td>
	-->
</tr>
<%
premaechul = oreport.FItemList(i).FSubSellTotal
prerank = oreport.FItemList(i).FRank
%>
<% next %>
</table>
<%
set oreport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->