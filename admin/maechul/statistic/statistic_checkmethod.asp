<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp" -->
<%
'	public Facct200			'��ġ��
'	public Facct900			'����Ʈī��
'	public Facct100			'�ſ�ī��
'	public Facct20			'�ǽð���ü
'	public Facct7			'������
'	public Facct400			'�޴���
'	public Facct560			'����Ƽ��
'	public Facct550			'������
'	public Facct110			'OK+�ſ�
'	public Facct80			'�þ�
'	public Facct50			'������
	
	
	Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vIsBanPum, vbizsec
	dim sellchnl
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= request("sitename")
	vDateGijun	= NullFillWith(request("date_gijun"),"regdate")
	vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",-13,now())))
	vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",-13,now())))
	vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",-13,now())))
	vEYear		= NullFillWith(request("eyear"),Year(now))
	vEMonth		= NullFillWith(request("emonth"),Month(now))
	vEDay		= NullFillWith(request("eday"),Day(now))
	vIsBanPum	= NullFillWith(request("isBanpum"),"all")
	vbizsec     = NullFillWith(request("bizsec"),"")
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	
	Dim vTot_Miletotalprice, vTot_Acct200, vTot_Acct900, vTot_Acct100, vTot_Acct20, vTot_Acct7, vTot_Acct400, vTot_Acct560, vTot_Acct550, vTot_Acct110, vTot_Acct80, vTot_Acct50, vTot_TotalSum
	
	Set cStatistic = New cStaticTotalClass_list
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectIsBanPum = vIsBanPum
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	cStatistic.FRectBizSectionCd = vbizsec
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.fStatistic_checkmethod()
%>

<script language="javascript">
function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}
	
function searchSubmit()
{
	if(frm.syear.value == <%=Year(v6MonthDate)%> && frm.smonth.value < <%=Month(v6MonthDate)%>)
	{
		alert("6������������ �ǽð��˻��� �����մϴ�.");
	}
	else
	{
		frm.submit();
	}
}
</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				* �Ⱓ :&nbsp;
				<select name="date_gijun" class="select">
					<option value="regdate" <%=CHKIIF(vDateGijun="regdate","selected","")%>>�ֹ���</option>
					<option value="ipkumdate" <%=CHKIIF(vDateGijun="ipkumdate","selected","")%>>������</option>
				</select>
				<%
					'### ��
					Response.Write "<select name=""syear"" class=""select"">"
					For i=Year(now) To Year(v6MonthDate) Step -1
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSYear),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"
					
					'### ��
					Response.Write "<select name=""smonth"" class=""select"">"
					For i=1 To 12
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSMonth),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"
					
					'### ��
					Response.Write "<select name=""sday"" class=""select"">"
					For i=1 To 31
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vSDay),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;~&nbsp;"
					
					'#############################
					
					'### ��
					Response.Write "<select name=""eyear"" class=""select"">"
					For i=Year(now) To Year(v6MonthDate) Step -1
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEYear),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"
					
					'### ��
					Response.Write "<select name=""emonth"" class=""select"">"
					For i=1 To 12
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEMonth),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>&nbsp;"
					
					'### ��
					Response.Write "<select name=""eday"" class=""select"">"
					For i=1 To 31
						Response.Write "<option value=""" & i & """ " & CHKIIF(CStr(i)=CStr(vEDay),"selected","") & ">" & i & "</option>"
					Next
					Response.Write "</select>"
					
					
					'### ����Ʈ����
					Response.Write "<br>* ����Ʈ���� : "
					Call Drawsitename("sitename", vSiteName)
					
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* �⺻ ����μ� : "
					Call DrawBizSectionGain("O,T","bizsec", vbizsec,"")
				%>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				* �ֹ����� :
				<select name="isBanpum" class="select">
					<option value="all" <%=CHKIIF(vIsBanPum="all","selected","")%>>��ǰ����</option>
					<option value="<>" <%=CHKIIF(vIsBanPum="<>","selected","")%>>��ǰ����</option>
					<option value="=" <%=CHKIIF(vIsBanPum="=","selected","")%>>��ǰ�Ǹ�</option>
				</select>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                	* ä�α��� 
                	<% drawSellChannelComboBox "sellchnl",sellchnl %> 
			</td>
		</tr>
	    </table>
	</td>	
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="�˻�" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<br>
�� �ǽð� �����ʹ� �ֱ� 6�������� �����͸� �˻� �����մϴ�.
<br>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" rowspan="2" colspan="2">�Ⱓ</td>
    <td align="center" colspan="3"></td>
    <td align="center" colspan="9">�ǰ�����</td>
    <td align="center" width="150" rowspan="2">�����հ�</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td align="center">���ϸ���</td>
    <td align="center">��ġ��</td>
    <td align="center">����Ʈī��</td>
    <td align="center">�ſ�ī��</td>
    <td align="center">�ǽð�����</td>
    <td align="center">������</td>
    <td align="center">�޴���</td>
    <td align="center">����Ƽ��</td>
    <td align="center">������</td>
    <td align="center">OKĳ�ù�</td>
    <td align="center">All@ī��</td>
    <td align="center">������</td>
</tr>
<%
	For i = 0 To cStatistic.FTotalCount -1
%>
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<% if right(FormatDateTime(cStatistic.flist(i).FRegdate,1),3) = "�����" then %>
				<font color="blue"><%= cStatistic.flist(i).FRegdate %></font>
			<% elseif right(FormatDateTime(cStatistic.flist(i).FRegdate,1),3) = "�Ͽ���" then %>
				<font color="red"><%= cStatistic.flist(i).FRegdate %></font>
			<% else %>
				<%= cStatistic.flist(i).FRegdate %>
			<% end if %>
		</td>
		<td align="center"><%= DateToWeekName(DatePart("w",cStatistic.FList(i).FRegdate)) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FMiletotalprice) %></td>
		<td align="right" style="padding-right:5px;" <%=CHKIIF(cStatistic.FList(i).FDifferent<>0,"bgcolor=""silver""","")%>><%= NullOrCurrFormat(cStatistic.FList(i).Facct200) %></td>
		<td align="right" style="padding-right:5px;" <%=CHKIIF(cStatistic.FList(i).FDifferent<>0,"bgcolor=""silver""","")%>><%= NullOrCurrFormat(cStatistic.FList(i).Facct900) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Facct100) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Facct20) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Facct7) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Facct400) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Facct560) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Facct550) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Facct110) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Facct80) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Facct50) %></td>
		<td align="right" style="padding-right:5px;" bgcolor="#E6B9B8"><b><%= NullOrCurrFormat(cStatistic.FList(i).FTotalSum) %></b></td>
	</tr>
<%
		vTot_Miletotalprice	= vTot_Miletotalprice + CLng(cStatistic.FList(i).FMiletotalprice)
		vTot_Acct200		= vTot_Acct200 + CLng(cStatistic.FList(i).Facct200)
		vTot_Acct900		= vTot_Acct900 + CLng(cStatistic.FList(i).Facct900)
		vTot_Acct100		= vTot_Acct100 + CLng(cStatistic.FList(i).Facct100)
		vTot_Acct20			= vTot_Acct20 + CLng(cStatistic.FList(i).Facct20)
		vTot_Acct7			= vTot_Acct7 + CLng(cStatistic.FList(i).Facct7)
		vTot_Acct400		= vTot_Acct400 + CLng(cStatistic.FList(i).Facct400)
		vTot_Acct560		= vTot_Acct560 + CLng(cStatistic.FList(i).Facct560)
		vTot_Acct550		= vTot_Acct550 + CLng(cStatistic.FList(i).Facct550)
		vTot_Acct110		= vTot_Acct110 + CLng(cStatistic.FList(i).Facct110)
		vTot_Acct80			= vTot_Acct80 + CLng(cStatistic.FList(i).Facct80)
		vTot_Acct50			= vTot_Acct50 + CLng(cStatistic.FList(i).Facct50)
		vTot_TotalSum		= vTot_TotalSum + CLng(cStatistic.FList(i).FTotalSum)
	
	Next
%>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" colspan="2">�հ�</td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Miletotalprice)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct200)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct900)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct100)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct20)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct7)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct400)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct560)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct550)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct110)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct80)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_Acct50)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_TotalSum)%></b></td>
</tr>
</table>

<% Set cStatistic = Nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->