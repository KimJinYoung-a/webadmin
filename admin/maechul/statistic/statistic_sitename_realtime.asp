<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �Ǹ�ó������
' History : ������ ����
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp" -->
<%
	Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vIsBanPum, vPurchasetype, vbizsec, vmakerid
	Dim cvDateGijun
	Dim mwdiv, noSubCost
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
	vPurchasetype = request("purchasetype")
	vbizsec     = NullFillWith(request("bizsec"),"")
	vmakerid    = NullFillWith(request("makerid"),"")
    mwdiv       = NullFillWith(request("mwdiv"),"")
	noSubCost	= NullFillWith(request("noSubCost"),"")
    
	Dim vTot_ItemNO, vTot_OrgitemCost, vTot_ItemcostCouponNotApplied, vTot_ItemCost, vTot_BuyCash, vTot_MaechulProfit, vTot_MaechulProfitPer
	Dim vTot_BonusCouponPrice, vTot_ReducedPrice, vTot_MaechulProfit2, vTot_MaechulProfitPer2

	Set cStatistic = New cStaticTotalClass_list
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectIsBanPum = vIsBanPum
	cStatistic.FRectPurchasetype = vPurchasetype
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	cStatistic.FRectBizSectionCd = vbizsec
	cStatistic.FRectMakerid = vmakerid
	cStatistic.FRectMwDiv = mwdiv
	cStatistic.FRectIncSubCost = noSubCost
	cStatistic.fStatistic_sitename_realtime()

	if vDateGijun="ipkumdate" then
	    cvDateGijun="ipkumil"
	elseif vDateGijun="beasongdate" then
	    cvDateGijun="chulgoil"
	elseif vDateGijun="jungsanfixdate" then
	    cvDateGijun="jungsanil"
	else
	    cvDateGijun="jumunil"
    end if

%>

<script language="javascript">
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
					<option value="beasongdate" <%=CHKIIF(vDateGijun="beasongdate","selected","")%>>�����</option>
					<option value="jungsanfixdate" <%=CHKIIF(vDateGijun="jungsanfixdate","selected","")%>>����Ȯ����</option>
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
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* ����Ʈ���� : "
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
				&nbsp;&nbsp;
				* ���Ա��� : 
				<% Call DrawBrandMWUCombo("mwdiv",mwdiv) %>
				&nbsp;&nbsp;
				<label><input type="checkbox" name="noSubCost" value="Y" <%=chkIIF(noSubCost="Y","checked","")%> /> ���/����� ����</label>
				<br>
				* �������� : 
				<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				* �귣�� : <% drawSelectBoxDesigner "makerid",vmakerid %></span>
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
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>�Ǹ�ó</td>
    <td>�Ǹż���</td>
    <td>�Һ��ڰ�[��ǰ]</td>
    <td>�ǸŰ�[��ǰ]<br>(��������)</td>
    <td><b>�����Ѿ�[��ǰ]<br>(��ǰ��������)</b></td>
    <td><b>���ʽ�����<br>����[��ǰ]</b></td>
    <td>��޾�</td>
    <td>�����Ѿ�[��ǰ]<br>(��ǰ��������)</td>
    <td><b>�������</b></td>
    <td>������</td>
    <td>�������2<br>(��޾ױ���)</td>
    <td>������</td>
    <td>���</td>
</tr>
<%
	For i = 0 To cStatistic.FTotalCount -1
%>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= cStatistic.flist(i).FSitename %></td>
	<td align="center"><%= CDbl(cStatistic.FList(i).FItemNO) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FOrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).FItemCost) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FItemCost-cStatistic.FList(i).FReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FBuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).FMaechulProfit) %></b></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).FMaechulProfitPer %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice-cStatistic.FList(i).FBuyCash) %></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).FMaechulProfitPer2 %>%</td>
	<td align="center" >
		[<a href="/admin/upchejungsan/upcheselllistByItem.asp?datetype=<%= cvDateGijun %>&yyyy1=<%=vSYear%>&mm1=<%=vSMonth%>&dd1=<%=vSDay%>&yyyy2=<%=vEYear%>&mm2=<%=vEMonth%>&dd2=<%=vEDay%>&delivertype=all&designer=<%=vmakerid%>" target="_blank">��ǰ��</a>]
		[<a href="/admin/upchejungsan/upcheselllist.asp?datetype=<%= cvDateGijun %>&yyyy1=<%=vSYear%>&mm1=<%=vSMonth%>&dd1=<%=vSDay%>&yyyy2=<%=vEYear%>&mm2=<%=vEMonth%>&dd2=<%=vEDay%>&delivertype=all&designer=<%=vmakerid%>" target="_blank">�Ǻ�</a>]
	</td>
</tr>
<%
	vTot_ItemNO						= vTot_ItemNO + CLng(NullOrCurrFormat(cStatistic.FList(i).FItemNO))
	vTot_OrgitemCost				= vTot_OrgitemCost + CDbl(NullOrCurrFormat(cStatistic.FList(i).FOrgitemCost))
	vTot_ItemcostCouponNotApplied	= vTot_ItemcostCouponNotApplied + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemcostCouponNotApplied))
	vTot_ItemCost					= vTot_ItemCost + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemCost))
	vTot_BonusCouponPrice			= vTot_BonusCouponPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemCost-cStatistic.FList(i).FReducedPrice))
	vTot_ReducedPrice				= vTot_ReducedPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FReducedPrice))
	vTot_BuyCash					= vTot_BuyCash + CDbl(NullOrCurrFormat(cStatistic.FList(i).FBuyCash))
	vTot_MaechulProfit				= vTot_MaechulProfit + CDbl(NullOrCurrFormat(cStatistic.FList(i).FMaechulProfit))
	vTot_MaechulProfit2				= vTot_MaechulProfit2 + CDbl(NullOrCurrFormat(cStatistic.FList(i).FReducedPrice-cStatistic.FList(i).FBuyCash))

	Next

	vTot_MaechulProfitPer = Round(((vTot_ItemCost - vTot_BuyCash)/CHKIIF(vTot_ItemCost=0,1,vTot_ItemCost))*100,2)
	vTot_MaechulProfitPer2 = Round(((vTot_ReducedPrice - vTot_BuyCash)/CHKIIF(vTot_ReducedPrice=0,1,vTot_ReducedPrice))*100,2)
%>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td align="center">�Ѱ�</td>
	<td align="center"><%=vTot_ItemNO%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_OrgitemCost)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_ItemcostCouponNotApplied)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_ItemCost)%></b></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_BonusCouponPrice)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_ReducedPrice)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_BuyCash)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_MaechulProfit)%></b></td>
	<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer%>%</td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_MaechulProfit2)%></td>
	<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer2%>%</td>
	<td></td>
</tr>
</table>
<% Set cStatistic = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
