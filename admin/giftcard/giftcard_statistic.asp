<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/giftCls.asp"-->

<%
	Dim GiftStatisticNew, iCurrentpage, page, vGubun, vSDate, vEDate, arrList, intLoop, iTotCnt, i, vTotal1, vTotal2, vTotal3, vTotal4, vParam
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	page 			= NullFillWith(requestCheckVar(request("page"),5),1)
	vGubun			= NullFillWith(requestCheckVar(request("gubun"),10),"")
	vSDate			= NullFillWith(requestCheckVar(request("sdate"),10),DateAdd("d",-15,date))
	vEDate			= NullFillWith(requestCheckVar(request("edate"),10),date)
	vTotal1 = 0
	vTotal2 = 0
	vTotal3 = 0
	vTotal4 = 0
	
	vParam = "&menupos="&Request("menupos")&"&gubun="&vGubun&"&sdate="&vSDate&"&edate="&vEDate&""
	
	Set GiftStatisticNew = new ClsGift
	GiftStatisticNew.FGubun = vGubun
	GiftStatisticNew.FSDate = vSDate
	GiftStatisticNew.FEDate = vEDate
	GiftStatisticNew.FGiftStatisticNew
	
	iTotCnt = GiftStatisticNew.ftotalcount
%>

<script language="javascript">
function chkfrm()
{

	return true;
}
</script>


<!-- 리스트 시작 -->
<form name="frm" method="get" action="giftcard_statistic.asp" onSubmit="return chkfrm()">
<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="40" bgcolor="FFFFFF">
	<td colspan="10">
		구매일 : 
		<input type="text" name="sdate" size="10" maxlength=10 readonly value="<%= vSDate %>">
		<a href="javascript:calendarOpen(frm.sdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		&nbsp;~&nbsp;
		<input type="text" name="edate" size="10" maxlength=10 readonly value="<%= vEDate %>">
		<a href="javascript:calendarOpen(frm.edate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		&nbsp;
		구분 : 
		<select name="gubun">
			<option value="">-전체-</option>
			<option value="10x10" <%=CHKIIF(vGubun="10x10","selected","")%>>10x10</option>
			<option value="550" <%=CHKIIF(vGubun="550","selected","")%>>기프팅</option>
			<option value="560" <%=CHKIIF(vGubun="560","selected","")%>>기프티콘</option>
		</select>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="submit" class="button" value="검 색">
		&nbsp;
	</td>
</tr>
</table>
</form>

<table cellpadding="0" cellspacing="0" border="0" class="a">
<tr height="30">
	<td align="right" width="535"><input type="button" value="결과값엑셀다운" class="button" onClick="location.href='giftcard_statistic_xls.asp?1=1<%=vParam%>';"></td>
</tr>
</table>

<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#E6E6E6">
	<td align="center" rowspan="2">금액권</td>
	<td align="center" colspan="2">구매</td>
	<td align="center" colspan="2">등록</td>
</tr>
<tr bgcolor="#E6E6E6">
	<td align="center">구매 건수</td>
	<td align="center">구매액</td>
	<td align="center">등록 건수</td>
	<td align="center">등록액</td>
</tr>
<%
	If GiftStatisticNew.FResultCount <> 0 Then
		For i = 0 To GiftStatisticNew.FResultCount -1
%>
		<tr bgcolor="FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
			<td width="100" align="center"><%= GetCardName(GiftStatisticNew.FItemList(i).fsubtotalprice) %></td>
			<td width="100" align="center"><%=GiftStatisticNew.FItemList(i).foidx%></td>
			<td width="100" align="center"><%=FormatNumber(GiftStatisticNew.FItemList(i).fsubtotalprice,0) %></td>
			<td width="100" align="center"><%=GiftStatisticNew.FItemList(i).fridx%></td>
			<td width="100" align="center"><%=FormatNumber(GiftStatisticNew.FItemList(i).fcardprice,0) %></td>
		</tr>
<%
			vTotal1 = CLng(vTotal1) + CLng(GiftStatisticNew.FItemList(i).foidx)
			vTotal2 = CLng(vTotal2) + CLng(GiftStatisticNew.FItemList(i).fsubtotalprice)
			vTotal3 = CLng(vTotal3) + CLng(GiftStatisticNew.FItemList(i).fridx)
			vTotal4 = CLng(vTotal4) + CLng(GiftStatisticNew.FItemList(i).fcardprice)
		Next
		
		Response.Write "<tr bgcolor=""FFFFFF"" height=""30""><td align=""center"" bgcolor=""#E6E6E6"">합계</td>"
		Response.Write "	<td align=""center"">" & FormatNumber(vTotal1,0) & "</td>"
		Response.Write "	<td align=""center"">" & FormatNumber(vTotal2,0) & "</td>"
		Response.Write "	<td align=""center"">" & FormatNumber(vTotal3,0) & "</td>"
		Response.Write "	<td align=""center"">" & FormatNumber(vTotal4,0) & "</td>"
		Response.Write "</tr>"
	Else
%>
		<tr bgcolor="#FFFFFF" height="30">
			<td colspan="20" align="center" class="page_link">[데이터가 없습니다.]</td>
		</tr>
<%
	End If
%>
</table>

<% Set GiftStatisticNew = Nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->