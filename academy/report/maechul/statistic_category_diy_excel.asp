<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �ΰŽ� ��������-ī�װ���
' History : 2016.03.15 corpse2 ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/report/maechul/statisticCls.asp" -->
<%
Dim i, cStatistic, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vCateL, vCateM
dim sellchnl, categbn, vCateS, vCateX, vIsBanPum, vBrandID, vCateGubun, vParam, vSiteName
Dim mwdiv, vCateMRate,vTot_CateMRate, dispCate, maxDepth, linkcate, linkdispcate, vSorting
	v6MonthDate	= DateAdd("m",-6,now())
	vDateGijun	= NullFillWith(RequestCheckvar(request("date_gijun"),16),"regdate")
	vSYear		= NullFillWith(RequestCheckvar(request("syear"),4),Year(DateAdd("d",0,now())))
	vSMonth		= NullFillWith(RequestCheckvar(request("smonth"),2),Month(DateAdd("d",0,now())))
	vSDay		= NullFillWith(RequestCheckvar(request("sday"),2),Day(DateAdd("d",0,now())))
	vEYear		= NullFillWith(RequestCheckvar(request("eyear"),4),Year(now))
	vEMonth		= NullFillWith(RequestCheckvar(request("emonth"),2),Month(now))
	vEDay		= NullFillWith(RequestCheckvar(request("eday"),2),Day(now))
	vCateL		= NullFillWith(RequestCheckvar(request("cdl"),3),"")
	vCateM		= NullFillWith(RequestCheckvar(request("cdm"),3),"")
	vCateS		= NullFillWith(RequestCheckvar(request("cds"),3),"")
	vCateX      = NullFillWith(request("cdx"),"")
	vIsBanPum	= NullFillWith(RequestCheckvar(request("isBanpum"),16),"all")
	vBrandID	= NullFillWith(RequestCheckvar(request("ebrand"),32),"")
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	mwdiv       = NullFillWith(RequestCheckvar(request("mwdiv"),1),"")
	categbn     = NullFillWith(request("categbn"),"")
    dispCate 	= requestCheckvar(request("disp"),16)
    maxDepth    = requestCheckvar(request("selDepth"),1) 
	vSorting	= NullFillWith(RequestCheckvar(request("sorting"),32),"categorynameD")

vSiteName = "diyitem"
if maxDepth = ""   then maxDepth = 0
vCateGubun = "L"
If vCateL <> "" and vCateM <> "" and vCateS<>"" Then
	'vCateGubun = "X"
	vCateGubun = "S"
ELSEIF vCateL <> "" and vCateM <> "" THEN
    vCateGubun = "S"
ELSEif vCateL <> "" Then
	vCateGubun = "M"
End IF
if (categbn="") then
    categbn="D"
end if
if categbn="M" then
    dispCate=""
elseif categbn="D" then
	vCateL="" : vCateM="" : vCateS="" : vCateX=""
end if

vParam = CurrURL() & "?menupos="&Request("menupos")&"&vSiteName="&vSiteName&"&date_gijun="&vDateGijun&"&syear="&vSYear&"&smonth="&vSMonth&"&sday="&vSDay&"&eyear="&vEYear&"&emonth="&vEMonth&"&eday="&vEDay&"&isBanpum="&vIsBanPum&"&ebrand="&vBrandID&"&mwdiv="&mwdiv&"&categbn="&categbn&"&sellchnl="&sellchnl

Dim vTot_OrderCnt, vTot_ItemNO, vTot_couponNotAsigncost, vTot_ItemCost, vTot_BuyCash
Dim vTot_MaechulProfit, vTot_MaechulProfitPer, vTot_BonusCouponPrice, vTot_ReducedPrice, vTot_MaechulProfit2, vTot_MaechulProfitPer2
dim vTot_upcheJungsan, vTot_avgipgoPrice, vTot_overValueStockPrice

Set cStatistic = New cacademyStatic_list
	cStatistic.FRectSiteName = "diyitem"
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectCateL = vCateL
	cStatistic.FRectCateM = vCateM
	cStatistic.FRectCateS = vCateS
	cStatistic.FRectCateGubun = vCateGubun
	cStatistic.FRectIsBanPum = vIsBanPum
	cStatistic.FRectMakerID = vBrandID
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.FRectMwDiv = mwdiv
	cStatistic.FRectCateGbn = categbn
	cStatistic.FRectIncStockAvgPrc = true '' ��ո��԰� ���� ��������.
	cStatistic.FRectSort = vSorting

	if (categbn="M") then
	    cStatistic.fStatistic_diy_category()
	else
	    cStatistic.FRectdispCate = dispCate
        cStatistic.FRectmaxDepth = maxdepth   
    	cStatistic.fStatistic_diy_DispCategory  ''2013/10/17 �߰�
    end if

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>

<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>

<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="25">
		�˻���� : <b><%=cStatistic.FresultCount%></b>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td><%=CateGubun(vCateGubun)%>ī�װ�</td>
    <td>��ǰ����</td>
	<% if (NOT C_InspectorUser) then %>
    <td>�ǸŰ�[��ǰ]<br>(��������)</td>
	<td>�����Ѿ�[��ǰ]<br>(��ǰ��������)</td>
	<td>���ʽ�����<br>����[��ǰ]</td>
	<% end if %>
	<td>��޾�</td>
	<td>�����Ѿ�[��ǰ]<% if (NOT C_InspectorUser) then %><br>(��ǰ��������)<% end if %></td>
	<td>�������</td>
	<td>������1</td>
	<td>�������2<br>(��޾ױ���)</td>
	<td>������2</td>
	<td>ī�װ���<br>���� ����</td>
	<td>��ü<br>�����</td>
	<td>ȸ�����</td>
</tr>
<% if cStatistic.FTotalCount > 0 then %>
<% For i = 0 To cStatistic.FTotalCount -1 %>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= cStatistic.FItemList(i).FCategoryName %></td>
	<td align="center"><%= FormatNumber(CDbl(cStatistic.FItemList(i).FItemNO),0) %></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fcouponNotAsigncost,0) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= FormatNumber(cStatistic.FItemList(i).FItemCost,0) %></b></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FItemCost-cStatistic.FItemList(i).FReducedPrice,0) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FReducedPrice,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FBuyCash,0) %></td>
	<td align="right" style="padding-right:5px;"><b><%= FormatNumber(cStatistic.FItemList(i).FMaechulProfit,0) %></b></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FItemList(i).FMaechulProfitPer %>%</td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FReducedPrice-cStatistic.FItemList(i).FBuyCash,0) %></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FItemList(i).FMaechulProfitPer2 %>%</td>
	<td align="right" style="padding-right:5px;"><%=formatnumber(vCateMRate,2)%>%</td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FupcheJungsan,0) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= FormatNumber(cStatistic.FItemList(i).FReducedPrice - cStatistic.FItemList(i).FupcheJungsan,0) %></b></td>
</tr>
<%
vTot_ItemNO						= vTot_ItemNO + CDbl(FormatNumber(cStatistic.FItemList(i).FItemNO,0))
vTot_couponNotAsigncost	= vTot_couponNotAsigncost + CDbl(FormatNumber(cStatistic.FItemList(i).fcouponNotAsigncost,0))
vTot_ItemCost					= vTot_ItemCost + CDbl(FormatNumber(cStatistic.FItemList(i).FItemCost,0))
vTot_BonusCouponPrice			= vTot_BonusCouponPrice + CDbl(FormatNumber(cStatistic.FItemList(i).FItemCost-cStatistic.FItemList(i).FReducedPrice,0))
vTot_ReducedPrice				= vTot_ReducedPrice + CDbl(FormatNumber(cStatistic.FItemList(i).FReducedPrice,0))
vTot_BuyCash					= vTot_BuyCash + CDbl(FormatNumber(cStatistic.FItemList(i).FBuyCash,0))
vTot_MaechulProfit				= vTot_MaechulProfit + CDbl(FormatNumber(cStatistic.FItemList(i).FMaechulProfit,0))
vTot_MaechulProfit2				= vTot_MaechulProfit2 + CDbl(FormatNumber(cStatistic.FItemList(i).FReducedPrice-cStatistic.FItemList(i).FBuyCash,0))
vTot_CateMRate					= vTot_CateMRate + vCateMRate
vTot_upcheJungsan				= vTot_upcheJungsan + CDbl(FormatNumber(cStatistic.FItemList(i).FupcheJungsan,0))
Next
vTot_MaechulProfitPer = Round(((vTot_ItemCost - vTot_BuyCash)/CHKIIF(vTot_ItemCost=0,1,vTot_ItemCost))*100,2)
vTot_MaechulProfitPer2 = Round(((vTot_ReducedPrice - vTot_BuyCash)/CHKIIF(vTot_ReducedPrice=0,1,vTot_ReducedPrice))*100,2)
%>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td align="center">�Ѱ�</td>
	<td align="center"><%=FormatNumber(vTot_ItemNO,0)%></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_couponNotAsigncost,0)%></td>
	<td align="right" style="padding-right:5px;"><b><%=FormatNumber(vTot_ItemCost,0)%></b></td>
	<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_BonusCouponPrice,0)%></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_ReducedPrice,0)%></td>
	<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_BuyCash,0)%></td>
	<td align="right" style="padding-right:5px;"><b><%=FormatNumber(vTot_MaechulProfit,0)%></b></td>
	<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer%>%</td>
	<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_MaechulProfit2,0)%></td>
	<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer2%>%</td>
	<td align="right" style="padding-right:5px;"><%=formatnumber(vTot_CateMRate,2)%>%</td>
	<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_upcheJungsan,0)%></td>
	<td align="right" style="padding-right:5px;"><b><%=FormatNumber(vTot_ReducedPrice - vTot_upcheJungsan,0)%></b></td>
</tr>
<% end if %>
</table>
<%
Function CateGubun(g)
	If g = "L" Then
		CateGubun = "��"
	ElseIf vCateGubun = "M" Then
		CateGubun = "��"
	ElseIf vCateGubun = "S" Then
		CateGubun = "��"
	End IF
End Function
Set cStatistic = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->