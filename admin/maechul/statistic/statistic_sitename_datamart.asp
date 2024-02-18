<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �¶��� ��������-�Ǹ�ó��
' History : 2012.10.09 ���ر� ����
'			2013.01.08 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_datamart.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->

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


	Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vIsBanPum, v6Ago
	dim sellchnl, inc3pl

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
	v6Ago		= NullFillWith(request("is6ago"),"")
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	inc3pl = request("inc3pl")

	Dim vTot_CountOrder, vTot_TotalSum, vTot_TenCardSpend, vTot_AllAtDiscountprice, vTot_Maechul, vTot_Miletotalprice, vTot_Subtotalprice

	Set cStatistic = New cStaticTotalClass_list
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectIsBanPum = vIsBanPum
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	cStatistic.FRect6MonthAgo = v6Ago
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.FRectInc3pl = inc3pl  ''2014/01/15 �߰�
	cStatistic.fStatistic_sitename()
%>

<script language="javascript">
function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}

function searchSubmit()
{
	if((frm.syear.value == <%=Year(v6MonthDate)%> && frm.smonth.value < <%=Month(v6MonthDate)%>) && (frm.is6ago.checked == false))
	{
		alert("6�������� �����ʹ� 6�������������͸� üũ�ϼž� �����մϴ�.");
	}
	else
	{
		if ((CheckDateValid(frm.syear.value, frm.smonth.value, frm.sday.value) == true) && (CheckDateValid(frm.eyear.value, frm.emonth.value, frm.eday.value) == true)) {
			frm.submit();
		}
	}
}

function detailStatistic(y1,m1,d1,y2,m2,d2,sitename,date_gijun,sellchnl,is6ago)
{
	var detailpop = window.open("/admin/maechul/statistic/statistic_daily_datamart.asp?syear="+y1+"&smonth="+m1+"&sday="+d1+"&eyear="+y2+"&emonth="+m2+"&eday="+d2+"&sitename="+sitename+"&date_gijun="+date_gijun+"&sellchnl="+sellchnl+"&is6ago="+is6ago,"detailpop","width=1000,height=780,scrollbars=yes,resizable=yes");
	detailpop.focus();
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
					For i=Year(now) To 2001 Step -1
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
					For i=Year(now) To 2001 Step -1 ''Year(v6MonthDate)
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


					'### 6��������������check
					Response.Write "<input type=""checkbox"" name=""is6ago"" value=""o"" "
					If v6Ago = "o" Then
						Response.Write "checked"
					End If
					Response.Write ">6��������������"

					'### ����Ʈ����
					Response.Write "<br>* ����Ʈ���� : "
					Call Drawsitename("sitename", vSiteName)
				%>
				&nbsp;&nbsp;
                	* ä�α���
                	<% drawSellChannelComboBox "sellchnl",sellchnl %>
				&nbsp;&nbsp;&nbsp;
				* �ֹ����� :
				<select name="isBanpum" class="select">
					<option value="all" <%=CHKIIF(vIsBanPum="all","selected","")%>>��ǰ����</option>
					<option value="<>" <%=CHKIIF(vIsBanPum="<>","selected","")%>>��ǰ����</option>
					<option value="=" <%=CHKIIF(vIsBanPum="=","selected","")%>>��ǰ�Ǹ�</option>
				</select>
				&nbsp;&nbsp;&nbsp;
				<b>* ����ó����</b>
        	    <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
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
* �˻� �Ⱓ�� ������� ����� �������ϴ�. �׷��� �˻� ��ư�� Ŭ���� �� �ƹ� ������ ����δٰ� ���� �˻���ư�� Ŭ������ ������.
<br>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td align="center">ä��</td>
	<td align="center">�Ǹ�óID(����Ʈ)</td>
    <td align="center">�Ǽ�</td>
    <% if (NOT C_InspectorUser) then %>
    <td align="center">�Һ��ڰ�</td>
    <td align="center">���αݾ�</td>
    <td align="center">�ǸŰ�<br>(���ΰ�)</td>
    <td align="center">��ǰ����<br>����</td>
    <td align="center">�����Ѿ�</td>
    <td align="center">���ʽ�����<br>����</td>
    <td align="center">��Ÿ����</td>
    <% end if %>
    <td align="center">�����</td>
    <td align="center">���</td>
    <!--<td align="center">���ϸ���</td>
    <td align="center">�����Ѿ�</td>-->
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"></td>
	<td align="center">�հ�</td>
	<td align="center"><span id="t1"></span></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="center"></td>
	<td align="center"></td>
	<td align="center"></td>
	<td align="center"></td>
	<td align="right" style="padding-right:5px;"><span id="t2"></span></td>
	<td align="right" style="padding-right:5px;"><span id="t3"></span></td>
	<td align="right" style="padding-right:5px;"><span id="t4"></span></td>
	<% end if %>
	<td align="right" style="padding-right:5px;"><b><span id="t5"></span></b></td>
	<td align="center"></td>
</tr>
<%
	For i = 0 To cStatistic.FTotalCount -1
%>
	<tr bgcolor="#FFFFFF">
		<td align="center"><%= getSellChannelName(cStatistic.flist(i).Fbeadaldiv) %></td>
		<td align="center"><%= cStatistic.flist(i).FSiteName %></td>
		<td align="center"><%= cStatistic.flist(i).FCountOrder %></td>
		<% if (NOT C_InspectorUser) then %>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="right" style="padding-right:5px;" bgcolor="#9DCFFF"><%= NullOrCurrFormat(CDBl(cStatistic.FList(i).FTotalSum)) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(CDBl(cStatistic.FList(i).FTenCardSpend)) %></td>
		<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(CDBl(cStatistic.FList(i).FAllAtDiscountprice)) %></td>
		 <% end if %>
		<td align="right" style="padding-right:5px;" bgcolor="#E6B9B8"><b><%= NullOrCurrFormat(CDBl(cStatistic.FList(i).FMaechul)) %></b></td>
		<td align="center" >
			[<a href="javascript:detailStatistic('<%=vSYear%>','<%=vSMonth%>','<%=vSDay%>','<%=vEYear%>','<%=vEMonth%>','<%=vEDay%>','<%= cStatistic.flist(i).FSiteName %>','<%= vDateGijun %>','<%= sellchnl %>','<%= v6Ago %>')">�Ϻ�</a>]
		</td>
		<!--<td align="right" style="padding-right:5px;"><%'= NullOrCurrFormat(CDBl(cStatistic.FList(i).FMiletotalprice)) %></td>
		<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><%'= NullOrCurrFormat(CDBl(cStatistic.FList(i).FSubtotalprice)) %></td>-->
	</tr>
<%
		vTot_CountOrder			= vTot_CountOrder + CDBl(NullOrCurrFormat(cStatistic.FList(i).FCountOrder))
		vTot_TotalSum			= vTot_TotalSum + CDBl(NullOrCurrFormat(cStatistic.FList(i).FTotalSum))
		vTot_TenCardSpend		= vTot_TenCardSpend + CDBl(NullOrCurrFormat(cStatistic.FList(i).FTenCardSpend))
		vTot_AllAtDiscountprice	= vTot_AllAtDiscountprice + CDBl(NullOrCurrFormat(cStatistic.FList(i).FAllAtDiscountprice))
		vTot_Maechul			= vTot_Maechul + CDBl(NullOrCurrFormat(cStatistic.FList(i).FMaechul))
		'vTot_Miletotalprice		= vTot_Miletotalprice + CDBl(NullOrCurrFormat(cStatistic.FList(i).FMiletotalprice))
		'vTot_Subtotalprice		= vTot_Subtotalprice + CDBl(NullOrCurrFormat(cStatistic.FList(i).FSubtotalprice))

	Next
%>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"></td>
	<td align="center">�հ�</td>
	<td align="center"><%=NullOrCurrFormat(vTot_CountOrder)%></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="center"></td>
	<td align="center"></td>
	<td align="center"></td>
	<td align="center"></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_TotalSum)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_TenCardSpend)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_AllAtDiscountprice)%></td>
	<% end if %>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_Maechul)%></b></td>
	<td align="center"></td>
	<!--<td align="right" style="padding-right:5px;"><%'=NullOrCurrFormat(vTot_Miletotalprice)%></td>
	<td align="right" style="padding-right:5px;"><%'=NullOrCurrFormat(vTot_Subtotalprice)%></td>-->
</tr>
</table>

<% If cStatistic.FTotalCount > 0 Then %>
<script>
document.getElementById("t1").innerHTML = "<%=NullOrCurrFormat(vTot_CountOrder)%>";
<% if (NOT C_InspectorUser) then %>
document.getElementById("t2").innerHTML = "<%=NullOrCurrFormat(vTot_TotalSum)%>";
document.getElementById("t3").innerHTML = "<%=NullOrCurrFormat(vTot_TenCardSpend)%>";
document.getElementById("t4").innerHTML = "<%=NullOrCurrFormat(vTot_AllAtDiscountprice)%>";
<% end if %>
document.getElementById("t5").innerHTML = "<%=NullOrCurrFormat(vTot_Maechul)%>";
</script>
<% End If %>

<% Set cStatistic = Nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
