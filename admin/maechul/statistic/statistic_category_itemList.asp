<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ī�װ�������
' History : ������ ����
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'####################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp" -->
<%
	Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vCateL, vCateM, vCateS, vIsBanPum, vBrandID, vCateGubun, vPurchasetype, vbizsec
	dim sellchnl, mwdiv, categbn, menupos, dispCate
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= requestCheckVar(request("sitename"),32)
	vDateGijun	= NullFillWith(request("date_gijun"),"regdate")
	vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",0,now())))
	vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",0,now())))
	vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",0,now())))
	vEYear		= NullFillWith(request("eyear"),Year(now))
	vEMonth		= NullFillWith(request("emonth"),Month(now))
	vEDay		= NullFillWith(request("eday"),Day(now))
	vCateL		= NullFillWith(request("cdl"),"")
	vCateM		= NullFillWith(request("cdm"),"")
	vCateS		= NullFillWith(request("cds"),"")
	vIsBanPum	= NullFillWith(request("isBanpum"),"all")
	vPurchasetype = requestCheckVar(request("purchasetype"),4)
	vBrandID	= NullFillWith(request("ebrand"),"")
	vbizsec     = NullFillWith(request("bizsec"),"")
	mwdiv       = NullFillWith(request("mwdiv"),"")
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	categbn     = NullFillWith(request("categbn"),"")
	dispCate	= requestCheckVar(request("disp"),16)

	if dispCate<>"" then
		vCateL = left(dispCate,3)
		vCateM = MID(dispCate,4,3)
		vCateS = MID(dispCate,7,3)
	end if

	vCateGubun = "L"
	If vCateL <> "" Then
		vCateGubun = "M"
	End IF
	If vCateM <> "" Then
		vCateGubun = "S"
	End IF
	if (categbn="") then
        categbn="D"
    end if
	dispCate = vCateL & vCateM & vCateS
	'<!-- //-->
	
	Dim vTot_OrderCnt, vTot_ItemNO, vTot_ItemCost, vTot_BuyCash, vTot_MaechulProfit, vTot_MaechulProfitPer
	Dim vTot_ReducedPrice, vTot_MaechulProfit2, vTot_MaechulProfitPer2
	
	Set cStatistic = New cStaticTotalClass_list
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectCateL = vCateL
	cStatistic.FRectCateM = vCateM
	cStatistic.FRectCateS = vCateS
	cStatistic.FRectCateGubun = vCateGubun
	cStatistic.FRectIsBanPum = vIsBanPum
	cStatistic.FRectPurchasetype = vPurchasetype
	cStatistic.FRectMakerID = vBrandID
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	cStatistic.FRectBizSectionCd = vbizsec
	cStatistic.FRectMwDiv = mwdiv
	cStatistic.FRectSellChannelDiv = sellchnl
 
	if (categbn="M") then
	    cStatistic.FStatisticCategoryItems
	else
    	cStatistic.FStatisticCategoryItems  ''2013/10/17 �߰�
    end if
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function searchSubmit()
{
	if(DateCheck() == false)
	{
		return;
	}

	if(frm.syear.value == <%=Year(v6MonthDate)%> && frm.smonth.value < <%=Month(v6MonthDate)%>)
	{
		alert("6������������ �ǽð��˻��� �����մϴ�.");
	}
	else
	{
		frm.submit();
	}
}

function DateCheck()
{
	var date1 = new Date(frm.syear.value,frm.smonth.value,frm.sday.value);
	var date2 = new Date(frm.eyear.value,frm.emonth.value,frm.eday.value);
		
	//�� �񱳰�
	var years  = date2.getFullYear() - date1.getFullYear();
	var months = date2.getMonth() - date1.getMonth();
	var days   = date2.getDate() - date1.getDate();

	var chkmonth = years * 12 + months + (days >= 0 ? 0 : -1);

	//�� �񱳰�
	var day   = 1000 * 3600 * 24;
	var chkday =  parseInt((date2 - date1) / day, 10);
	
	if(chkday > 31)
	{
		alert("��¥ �˻��� 1�� ���ݸ� �˴ϴ�.");
		return false;
	}
	else
	{
		return true;
	}
}
</script>

<div style="font-weight:bold; margin-bottom:10px;padding:6px 10px;border-radius:6px;background-color:#EEF;">[���]ON����-�ǽð�>>ī�װ������� - ��ǰ���</div>
<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="categbn" value="<%=categbn%>" />
<input type="hidden" name="sitename" value="<%=vSiteName%>" />
<input type="hidden" name="bizsec" value="<%=vbizsec%>" />
<input type="hidden" name="isBanpum" value="<%=vIsBanPum%>" />
<input type="hidden" name="sellchnl" value="<%=sellchnl%>" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				<% if left(dispCate,3)<>"999" then %>
				* ����ī�װ�: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
				<% else %>
				* ����ī�װ� : <b>������</b>
				<% end if %>
				<br />
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
				%>
				&nbsp;&nbsp;
				* ���Ա��� : 
				<% Call DrawBrandMWUCombo("mwdiv",mwdiv) %>
				<br />
				* �귣�� : <input type="text" class="text" name="ebrand" value="<%=vBrandID%>" size="20"> <input type="button" class="button" value="IDSearch" onclick="jsSearchBrandID(this.form.name,'ebrand');">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				* �������� : 
				<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
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
<tr bgcolor="<%= adminColor("tabletop") %>"  align="center">
	<td>��ǰ�ڵ�</td>
	<td>�귣��</td>
	<td>��ǰ��</td>
    <td>�Ǹż���</td>
    <% if (NOT C_InspectorUser) then %>
    <td><b>�����Ѿ�[��ǰ]<br>(��ǰ��������)</b></td>
    <% end if %>
    <td>��޾�</td>
    <td>�����Ѿ�[��ǰ]<% if (NOT C_InspectorUser) then %><br>(��ǰ��������)<% end if %></td>
    <td><b>�������</b></td>
    <td>������</td>
    <td>�������2<br>(��޾ױ���)</td>
    <td>������</td>
</tr>
<%
	For i = 0 To cStatistic.FTotalCount -1
%>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= cStatistic.FList(i).FItemId %></td>
	<td align="center"><%= cStatistic.FList(i).FMakerID %></td>
	<td style="padding-left:5px;"><%= cStatistic.FList(i).FitemName %></td>
	<td align="center"><%= cStatistic.FList(i).FItemNO %></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).FItemCost) %></b></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FBuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).FMaechulProfit) %></b></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).FMaechulProfitPer %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice-cStatistic.FList(i).FBuyCash) %></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).FMaechulProfitPer2 %>%</td>
</tr>
<%
	vTot_ItemNO						= vTot_ItemNO + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemNO))
	vTot_ItemCost					= vTot_ItemCost + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemCost))
	vTot_ReducedPrice				= vTot_ReducedPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FReducedPrice))
	vTot_BuyCash					= vTot_BuyCash + CDbl(NullOrCurrFormat(cStatistic.FList(i).FBuyCash))
	vTot_MaechulProfit				= vTot_MaechulProfit + CDbl(NullOrCurrFormat(cStatistic.FList(i).FMaechulProfit))
	vTot_MaechulProfit2				= vTot_MaechulProfit2 + CDbl(NullOrCurrFormat(cStatistic.FList(i).FReducedPrice-cStatistic.FList(i).FBuyCash))
	
	Next
	
	vTot_MaechulProfitPer = Round(((vTot_ItemCost - vTot_BuyCash)/CHKIIF(vTot_ItemCost=0,1,vTot_ItemCost))*100,2)
	vTot_MaechulProfitPer2 = Round(((vTot_ReducedPrice - vTot_BuyCash)/CHKIIF(vTot_ReducedPrice=0,1,vTot_ReducedPrice))*100,2)
%>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td align="center" colspan="3">�Ѱ�</td>
	<td align="center"><%=vTot_ItemNO%></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_ItemCost)%></b></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_ReducedPrice)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_BuyCash)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_MaechulProfit)%></b></td>
	<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer%>%</td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_MaechulProfit2)%></td>
	<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer2%>%</td>
</tr>
</table>
<% Set cStatistic = Nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->