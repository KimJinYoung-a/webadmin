<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �ΰŽ� ��������-��ǰ������
' History : 2016.09.21 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/report/maechul/statisticCls.asp" -->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->

<%
Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vSorting
dim sellchnl, vCateL, vCateM, vCateS, vIsBanPum, mwdiv
dim iCurrPage,iPageSize,iTotalPage,iTotCnt, dispCate,vBrandID ,itemid, sVType
dim  vTotwww_ItemNO,vTotwww_ItemCost,vTotwww_MaechulProfit,vTotwww_BuyCash,vTotma_ItemNO,vTotma_ItemCost,vTotma_MaechulProfit
dim vTotma_BuyCash,vTotout_ItemNO,vTotout_ItemCost,vTotout_MaechulProfit	,vTotout_BuyCash			
dim vTotwww_MaechulProfitPer ,vTotma_MaechulProfitPer ,vTotout_MaechulProfitPer 
Dim vTot_OrderCnt, vTot_ItemNO, vTot_ItemcostCouponNotApplied, vTot_ItemCost, vTot_BuyCash, vTot_MaechulProfit
Dim vTot_MaechulProfitPer, vTot_BonusCouponPrice, vTot_ReducedPrice, vTot_MaechulProfit2, vTot_MaechulProfitPer2
dim vTot_upcheJungsan, lec_cdl, lec_cdm
	iPageSize = 100
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= RequestCheckvar(request("sitename"),16)
	vDateGijun	= NullFillWith(RequestCheckvar(request("date_gijun"),16),"regdate")
	vSYear		= NullFillWith(RequestCheckvar(request("syear"),4),Year(DateAdd("d",0,now())))
	vSMonth		= NullFillWith(RequestCheckvar(request("smonth"),2),Month(DateAdd("d",0,now())))
	vSDay		= NullFillWith(RequestCheckvar(request("sday"),2),Day(DateAdd("d",0,now())))
	vEYear		= NullFillWith(RequestCheckvar(request("eyear"),4),Year(now))
	vEMonth		= NullFillWith(RequestCheckvar(request("emonth"),2),Month(now))
	vEDay		= NullFillWith(RequestCheckvar(request("eday"),2),Day(now))
	vSorting	= NullFillWith(RequestCheckvar(request("sorting"),32),"itemcostD")
	vBrandID	= NullFillWith(RequestCheckvar(request("ebrand"),32),"")
	vCateL		= NullFillWith(RequestCheckvar(request("cdl"),3),"")
	vCateM		= NullFillWith(RequestCheckvar(request("cdm"),3),"")
	vCateS		= NullFillWith(RequestCheckvar(request("cds"),3),"")
	dispCate = requestCheckvar(request("disp"),16)
	itemid      = requestCheckvar(request("itemid"),255)
	vIsBanPum	= NullFillWith(RequestCheckvar(request("isBanpum"),16),"all")
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	mwdiv       = NullFillWith(RequestCheckvar(request("mwdiv"),1),"")
	iCurrPage =requestCheckVar(request("iC"),4)
	sVType      = requestCheckvar(request("rdoVType"),1)
	lec_cdl = RequestCheckvar(request("lec_cdl"),3)
	lec_cdm = RequestCheckvar(request("lec_cdm"),3)
  	if itemid <> "" then
		if checkNotValidHTML(itemid) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
if iCurrPage = "" then iCurrPage = 1
if sVType ="" then sVType = 1

if itemid<>"" then
	dim iA ,arrTemp,arrItemid
	itemid = replace(itemid,",",chr(10))
  	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

Set cStatistic = New cacademyStatic_list
	cStatistic.FRectlec_cdl = lec_cdl
	cStatistic.FRectlec_cdm = lec_cdm
	cStatistic.FRectSort = vSorting
	cStatistic.FRectCateL = vCateL
	cStatistic.FRectCateM = vCateM
	cStatistic.FRectCateS = vCateS
	cStatistic.FRectIsBanPum = vIsBanPum
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.FRectMwDiv = mwdiv
	cStatistic.FRectMakerid = vBrandID
	cStatistic.FRectDispCate = dispCate
	cStatistic.FRectItemid   = itemid 
	cStatistic.FRectVType = sVType
	cStatistic.FPageSize = iPageSize
	cStatistic.FCurrPage = iCurrPage
	cStatistic.FRectIncStockAvgPrc = true '' ��ո��԰� ���� ��������.
	cStatistic.fStatistic_item()

	iTotCnt = cStatistic.FResultCount
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function downloadexcel(){
    document.frm.target = "view"; 
    document.frm.action = "/academy/report/maechul/statistic_item_excel.asp";  
	document.frm.submit();
    document.frm.target = ""; 
    document.frm.action = "";  
}

function searchSubmit(){
    document.frm.target = "_self"; 
    document.frm.action = "/academy/report/maechul/statistic_item.asp";  
	
	if ((CheckDateValid(frm.syear.value, frm.smonth.value, frm.sday.value) == true) && (CheckDateValid(frm.eyear.value, frm.emonth.value, frm.eday.value) == true)) {
		$("#btnSubmit").prop("disabled", true);
		frm.submit(); 
	}
}

function frontitemlink(sitename, itemid){
	var linkurl;

	if (sitename=='diyitem'){
		linkurl = '/diyshop/shop_prd.asp?itemid=' + itemid
	} else if (sitename=='academy'){
		linkurl = '/lecture/lecturedetail.asp?lec_idx=' + itemid
	}else{
		alert('�����ڰ� �����ϴ�.');
	}
	
	var popwin = window.open('<%= wwwFingers %>'+linkurl,'addreg','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function jstrSort(vsorting){
	var tmpSorting = document.getElementById("img"+vsorting)

	if (-1 < tmpSorting.src.indexOf("_alpha")){
		frm.sorting.value= vsorting+"D";
	}else if (-1 < tmpSorting.src.indexOf("_bot")){
		frm.sorting.value= vsorting+"A";
	}else{
		frm.sorting.value= vsorting+"D";
	}
	searchSubmit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>"> 
<input type="hidden" name="iC" value="">
<input type="hidden" name="sorting" value="<%= vsorting %>">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"  >
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>" rowspan=4>�˻� ����</td>
	<td align="left">
		* �Ⱓ:
		<select name="date_gijun" class="select">
			<option value="regdate" <%=CHKIIF(vDateGijun="regdate","selected","")%>>�ֹ���</option>
			<option value="ipkumdate" <%=CHKIIF(vDateGijun="ipkumdate","selected","")%>>������</option>
			<option value="beasongdate" <%=CHKIIF(vDateGijun="beasongdate","selected","")%>>�����</option>
		</select>
		<% DrawDateBoxdynamic vSYear,"syear",vEYear,"eyear",vSMonth,"smonth",vEMonth,"emonth",vSDay,"sday",vEDay,"eday" %>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan=4><input type="button" id="btnSubmit" class="button_s" value="�˻�" onClick="searchSubmit();"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<table class="a" cellpadding="0" border=0 cellspacing="0" width="100%">
		<tr>
			<td colspan="2" width=400>
				* �귣�� : <input type="text" class="text" name="ebrand" value="<%=vBrandID%>" size="20"> <input type="button" class="button" value="IDSearch" onclick="jsSearchBrandID(this.form.name,'ebrand');">
			    &nbsp;
			    * ��ǰ�ڵ� :
			</td>
			<td rowspan="2" align="left"  ><textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea></td>
		</tr> 
		<tr>
			<td width=400>
				* ����ƮŸ��:
			    <input type="radio" name="rdoVType" value="1" <%=CHKIIF(sVType="1","checked","")%>>��ǰ�� 
			    <input type="radio" name="rdoVType" value="2" <%=CHKIIF(sVType="2","checked","")%>>��¥��
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ä�α��� : <% drawSelectBox_SellChannel "sellchnl", sellchnl, "" %>
		&nbsp;
		* �ֹ����� :
		<select name="isBanpum" class="select">
			<option value="all" <%=CHKIIF(vIsBanPum="all","selected","")%>>��ǰ����</option>
			<option value="<>" <%=CHKIIF(vIsBanPum="<>","selected","")%>>��ǰ����</option>
			<option value="=" <%=CHKIIF(vIsBanPum="=","selected","")%>>��ǰ�Ǹ�</option>
		</select>
		&nbsp;
		* ���Ա��� : <% Call DrawBrandMWUPCombo("mwdiv",mwdiv) %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ����Ʈ���� : <% drawradio_academy_sitename "sitename", vSiteName, "", "Y" %>
		<% if vSiteName="" then %>
			�� ī�װ����� �˻� �ϽǷ��� ����Ʈ������ �����ϼž� �մϴ�.
		<% elseif vSiteName="academy" then %>
			* ī�װ��� :  <% DrawSelectBoxLecCategoryLarge "lec_cdl" ,  lec_cdl  , "N" %>

			<% if lec_cdl <> "" Then %>
				* ��ī�װ��� : <% call DrawSelectBoxLecCategoryMid("lec_cdm", lec_cdl, lec_cdm, "N") %>
			<% end if %>
		<% elseif vSiteName="diyitem" then %>
			* ���<!-- #include virtual="/academy/comm/CategorySelectBox.asp"-->
			<script type="text/javascript">
			$(function(){
				chgDispCate('<%=dispCate%>');
			});
			
			function chgDispCate(dc) {
				$.ajax({
					url: "/academy/comm/dispCateSelectBox_response.asp?disp="+dc,
					cache: false,
					async: false,
					success: function(message) {
			       		// ���� �ֱ� 
			       		$("#lyrDispCtBox").empty().html(message);
			       		$("#oDispCate").val(dc);
					}
				});
			}
			</script>
			* ����ī�װ��� : <span id="lyrDispCtBox"></span>
			<input type="hidden" name="disp" id="oDispCate" value="<%=dispCate%>">
		<% end if %>
	</td>
</tr>
</table>
<!-- �˻� �� -->
<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left"></td>
	<td align="right">	
		<!--* ���� : 
		<input type="radio" name="sorting" value="itemno" <%'=CHKIIF(vSorting="itemno","checked","")%>>������
		<input type="radio" name="sorting" value="itemcost" <%'=CHKIIF(vSorting="itemcost","checked","")%>>�����
		<input type="radio" name="sorting" value="profit" <%'=CHKIIF(vSorting="profit","checked","")%>>���ͼ� -->
		<input type="button" onclick="downloadexcel();" value="�����ٿ�ε�" class="button">
	</td>
</tr>
</table>
<!-- �׼� �� -->

</form>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		�˻���� : <b><%=iTotCnt%></b>
		&nbsp;
		������ : <b><%= iCurrPage %> / <%=iTotalPage%></b>
	</td> 
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>" align="center">  

	<%IF sVType = 2  then%>
		<td onClick="jstrSort('ddate'); return false;" style="cursor:hand;">
			��¥
			<img src="/images/list_lineup<%=CHKIIF(vSorting="ddateD","_bot","_top")%><%=CHKIIF(instr(vSorting,"ddate")>0,"_on","")%>.png" id="imgddate">
		</td>
	<%END IF%>

	<td onClick="jstrSort('sitename'); return false;" style="cursor:hand;">
		����Ʈ����
		<img src="/images/list_lineup<%=CHKIIF(vSorting="sitenameD","_bot","_top")%><%=CHKIIF(instr(vSorting,"sitename")>0,"_on","")%>.png" id="imgsitename">
	</td>
	<td>��ǰ�ڵ�</td>
	<td>�귣��</td>
	<td onClick="jstrSort('itemno'); return false;" style="cursor:hand;">
		��ǰ����
		<img src="/images/list_lineup<%=CHKIIF(vSorting="itemnoD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemno")>0,"_on","")%>.png" id="imgitemno">
	</td>

	<% if (NOT C_InspectorUser) then %>
	<td onClick="jstrSort('couponnotasigncost'); return false;" style="cursor:hand;">
		�ǸŰ�[��ǰ]<br>(��������)
		<img src="/images/list_lineup<%=CHKIIF(vSorting="couponnotasigncostD","_bot","_top")%><%=CHKIIF(instr(vSorting,"couponnotasigncost")>0,"_on","")%>.png" id="imgcouponnotasigncost">
	</td>
	<td onClick="jstrSort('itemcost'); return false;" style="cursor:hand;">
		<b>�����Ѿ�[��ǰ]<br>(��ǰ��������)</b>
		<img src="/images/list_lineup<%=CHKIIF(vSorting="itemcostD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemcost")>0,"_on","")%>.png" id="imgitemcost">
	</td>
	<td onClick="jstrSort('itemCostnotexistsbonus'); return false;" style="cursor:hand;">
		<b>���ʽ�����<br>����[��ǰ]</b>
		<img src="/images/list_lineup<%=CHKIIF(vSorting="itemCostnotexistsbonusD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemCostnotexistsbonus")>0,"_on","")%>.png" id="imgitemCostnotexistsbonus">
	</td>
	<% end if %>

	<td onClick="jstrSort('reducedprice'); return false;" style="cursor:hand;">
		��޾�
		<img src="/images/list_lineup<%=CHKIIF(vSorting="reducedpriceD","_bot","_top")%><%=CHKIIF(instr(vSorting,"reducedprice")>0,"_on","")%>.png" id="imgreducedprice">
	</td>
	<td onClick="jstrSort('buycash'); return false;" style="cursor:hand;">
		�����Ѿ�[��ǰ]<% if (NOT C_InspectorUser) then %><br>(��ǰ��������)<% end if %>
		<img src="/images/list_lineup<%=CHKIIF(vSorting="buycashD","_bot","_top")%><%=CHKIIF(instr(vSorting,"buycash")>0,"_on","")%>.png" id="imgbuycash">
	</td>
	<td onClick="jstrSort('maechulprofit1'); return false;" style="cursor:hand;">
		<b>�������</b>
		<img src="/images/list_lineup<%=CHKIIF(vSorting="maechulprofit1D","_bot","_top")%><%=CHKIIF(instr(vSorting,"maechulprofit1")>0,"_on","")%>.png" id="imgmaechulprofit1">
	</td>
	<td onClick="jstrSort('maechulprofitper1'); return false;" style="cursor:hand;">
		������1
		<img src="/images/list_lineup<%=CHKIIF(vSorting="maechulprofitper1D","_bot","_top")%><%=CHKIIF(instr(vSorting,"maechulprofitper1")>0,"_on","")%>.png" id="imgmaechulprofitper1">
	</td>
	<td onClick="jstrSort('maechulprofit2'); return false;" style="cursor:hand;">
		�������2<br>(��޾ױ���)
		<img src="/images/list_lineup<%=CHKIIF(vSorting="maechulprofit2D","_bot","_top")%><%=CHKIIF(instr(vSorting,"maechulprofit2")>0,"_on","")%>.png" id="imgmaechulprofit2">
	</td>
	<td onClick="jstrSort('maechulprofitper2'); return false;" style="cursor:hand;">
		������2
		<img src="/images/list_lineup<%=CHKIIF(vSorting="maechulprofitper2D","_bot","_top")%><%=CHKIIF(instr(vSorting,"maechulprofitper2")>0,"_on","")%>.png" id="imgmaechulprofitper2">
	</td>
	<td onClick="jstrSort('upchejungsan'); return false;" style="cursor:hand;">
		��ü<br>�����
		<img src="/images/list_lineup<%=CHKIIF(vSorting="upchejungsanD","_bot","_top")%><%=CHKIIF(instr(vSorting,"upchejungsan")>0,"_on","")%>.png" id="imgupchejungsan">
	</td>
	<td onClick="jstrSort('reducedpricenotexistsupchejungsan'); return false;" style="cursor:hand;">
		<b>ȸ�����</b>
		<img src="/images/list_lineup<%=CHKIIF(vSorting="reducedpricenotexistsupchejungsanD","_bot","_top")%><%=CHKIIF(instr(vSorting,"reducedpricenotexistsupchejungsan")>0,"_on","")%>.png" id="imgreducedpricenotexistsupchejungsan">
	</td>
</tr>

<% if cStatistic.FTotalCount > 0 then %>
	<% For i = 0 To cStatistic.FTotalCount -1 %>
	<tr bgcolor="#FFFFFF">
		<%IF sVType = 2 then%>
			<td align="center"><%= cStatistic.FItemList(i).Fddate %></td>
		<%END IF%>

		<td align="center"><%= get_academy_sitename(cStatistic.FItemList(i).fsitename) %></td>
		<td align="center">
			<a href="#" onclick="frontitemlink('<%= cStatistic.FItemList(i).fsitename %>','<%= cStatistic.FItemList(i).FitemID %>'); return false;">
			<%= cStatistic.FItemList(i).FitemID %></a>
		</td>
		<td align="center">
			<a href="#" onclick="frontitemlink('<%= cStatistic.FItemList(i).fsitename %>','<%= cStatistic.FItemList(i).FitemID %>'); return false;">
			<%=cStatistic.FItemList(i).FMakerID%></a>
		</td>
		<td align="center"><%= CDbl(cStatistic.FItemList(i).FItemNO) %></td>

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
		<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FupcheJungsan,0) %></td>
		<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= FormatNumber(cStatistic.FItemList(i).FReducedPrice - cStatistic.FItemList(i).FupcheJungsan,0) %></b></td>
	</tr>
	<%
	vTot_ItemNO						= vTot_ItemNO + CLng(FormatNumber(cStatistic.FItemList(i).FItemNO,0))
	vTot_ItemcostCouponNotApplied	= vTot_ItemcostCouponNotApplied + CLng(FormatNumber(cStatistic.FItemList(i).fcouponNotAsigncost,0))
	vTot_ItemCost					= vTot_ItemCost + CLng(FormatNumber(cStatistic.FItemList(i).FItemCost,0))
	vTot_BonusCouponPrice			= vTot_BonusCouponPrice + CDbl(FormatNumber(cStatistic.FItemList(i).FItemCost-cStatistic.FItemList(i).FReducedPrice,0))
	vTot_ReducedPrice				= vTot_ReducedPrice + CDbl(FormatNumber(cStatistic.FItemList(i).FReducedPrice,0))
	vTot_BuyCash					= vTot_BuyCash + CLng(FormatNumber(cStatistic.FItemList(i).FBuyCash,0))
	vTot_MaechulProfit				= vTot_MaechulProfit + CLng(FormatNumber(cStatistic.FItemList(i).FMaechulProfit,0))
	vTot_MaechulProfit2				= vTot_MaechulProfit2 + CDbl(FormatNumber(cStatistic.FItemList(i).FReducedPrice-cStatistic.FItemList(i).FBuyCash,0))
	vTot_upcheJungsan				= vTot_upcheJungsan + CDbl(FormatNumber(cStatistic.FItemList(i).FupcheJungsan,0))
	Next
	
	vTot_MaechulProfitPer = Round(((vTot_ItemCost - vTot_BuyCash)/CHKIIF(vTot_ItemCost=0,1,vTot_ItemCost))*100,2)
	vTot_MaechulProfitPer2 = Round(((vTot_ReducedPrice - vTot_BuyCash)/CHKIIF(vTot_ReducedPrice=0,1,vTot_ReducedPrice))*100,2)
	%>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td align="center">�Ѱ�</td>
		<td align="center" colspan=2></td>

		<% if sVType="2" then %>
			<td align="center"></td>
		<% end if %>

		<td align="center"><%=vTot_ItemNO%></td>

		<% if (NOT C_InspectorUser) then %>
			<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_ItemcostCouponNotApplied,0)%></td>
			<td align="right" style="padding-right:5px;"><b><%=FormatNumber(vTot_ItemCost,0)%></b></td>
			<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_BonusCouponPrice,0)%></td>
	    <% end if %>

		<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_ReducedPrice,0)%></td>
		<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_BuyCash,0)%></td>
		<td align="right" style="padding-right:5px;"><b><%=FormatNumber(vTot_MaechulProfit,0)%></b></td>
		<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer%>%</td>
		<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_MaechulProfit2,0)%></td>
		<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer2%>%</td>
		<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_upcheJungsan,0)%></td>
		<td align="right" style="padding-right:5px;"><b><%=FormatNumber(vTot_ReducedPrice - vTot_upcheJungsan,0)%></b></td>
	</tr>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="25" align="center">
	       	<%'sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	</tr>
<% ELSE %>
	<tr  align="center" bgcolor="#FFFFFF">
		<td colspan="25">��ϵ� ������ �����ϴ�.</td>
	</tr>
<% end if %>

</table>

<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>

<%
Set cStatistic = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->