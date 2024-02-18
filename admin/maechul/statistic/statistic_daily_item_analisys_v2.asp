<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �¶��� ���ڵ� ���
' Hieditor : 2016.01.20 ������ ����
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_dw.asp" -->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->

<%
	Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, vIsBanPum, vPurchasetype, v6Ago, vmakerid
	dim sellchnl, inc3pl, vSorting, dispCate, maxDepth
	Dim mwdiv, chkShowGubun,itemid, showsuply
	Dim incStockAvg, isSendGift
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= request("sitename")
	vDateGijun	= NullFillWith(request("date_gijun"),"regdate")  ''beasongdate  :�����=>�ֹ��� 2018/05/28  by eastone
	vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",-13,now())))
	vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",-13,now())))
	vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",-13,now())))
	vEYear		= NullFillWith(request("eyear"),Year(now))
	vEMonth		= NullFillWith(request("emonth"),Month(now))
	vEDay		= NullFillWith(request("eday"),Day(now))
	vIsBanPum	= NullFillWith(request("isBanpum"),"all")
	vPurchasetype = request("purchasetype")
	v6Ago		= NullFillWith(request("is6ago"),"")
	dispCate	= requestCheckVar(request("disp"),20)
	maxDepth = "1"
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	vmakerid    = NullFillWith(request("makerid"),"")
	mwdiv       = NullFillWith(request("mwdiv"),"")
	itemid      = requestCheckvar(request("itemid"),255)
	inc3pl      = request("inc3pl")
	chkShowGubun = request("chkShowGubun")
	vSorting	= NullFillWith(request("sorting"),"yyyymmddD")
	showsuply   = requestCheckvar(request("showsuply"),10)
	incStockAvg = requestCheckvar(request("incStockAvg"),10)
	isSendGift	= requestCheckvar(request("isSendGift"),1)

	Dim vTot_countOrder, vTot_ItemNO, vTot_OrgitemCost, vTot_ItemcostCouponNotApplied, vTot_ItemCost, vTot_BuyCash, vTot_MaechulProfit, vTot_MaechulProfitPer
	Dim vTot_BonusCouponPrice, vTot_ReducedPrice, vTot_MaechulProfit2, vTot_MaechulProfitPer2
	dim vTot_upcheJungsan, vTot_avgipgoPrice, vTot_overValueStockPrice

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

	Set cStatistic = New cStaticTotalClass_list
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectIsBanPum = vIsBanPum
	cStatistic.FRectPurchasetype = vPurchasetype
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	''cStatistic.FRect6MonthAgo = v6Ago
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.FRectMwDiv = mwdiv
	cStatistic.FRectMakerid = vmakerid
	cStatistic.FRectInc3pl = inc3pl  ''2014/01/15 �߰�
	cStatistic.FRectChkShowGubun = chkShowGubun					'// 2015-10-22, skyer9
	cStatistic.FRectIncStockAvgPrc = (incStockAvg<>"") ''true '' ��ո��԰� ���� ��������.
	cStatistic.FRectItemid   = itemid  '/2016-03-18 �߰�
	cStatistic.FRectDispCate = dispCate
	cStatistic.FRectSort = vSorting
	cStatistic.FRectBySuplyPrice = CHKIIF(showsuply="on",1,0)
	cStatistic.FRectIsSendGift = isSendGift
	cStatistic.fStatistic_daily_item()
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

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

function searchSubmit(){
	//if((frm.syear.value == <%=Year(v6MonthDate)%> && frm.smonth.value < <%=Month(v6MonthDate)%>) && (frm.is6ago.checked == false))
	//{
	//	alert("6�������� �����ʹ� 6�������������͸� üũ�ϼž� �����մϴ�.");
	//}
	//else
	//{
		if ((CheckDateValid(frm.syear.value, frm.smonth.value, frm.sday.value) == true) && (CheckDateValid(frm.eyear.value, frm.emonth.value, frm.eday.value) == true)) {
			frm.submit();
		}
	//}
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="sorting" value="<%= vsorting %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">�˻�</td>
	<td align="left">
		<table class="a" border="0">
		<tr>
			<td>
				* �Ⱓ :
				<select name="date_gijun" class="select">
					<option value="regdate" <%=CHKIIF(vDateGijun="regdate","selected","")%>>�ֹ���</option>
					<option value="ipkumdate" <%=CHKIIF(vDateGijun="ipkumdate","selected","")%>>������</option>
					<option value="beasongdate" <%=CHKIIF(vDateGijun="beasongdate","selected","")%>>��ǰ�����</option>
					<option value="jfixeddt" <%=CHKIIF(vDateGijun="jfixeddt","selected","")%>>����Ȯ����</option>
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
					For i=Year(now) To 2001 Step -1
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
					'Response.Write "<input type=""checkbox"" name=""is6ago"" value=""o"" "
					'If v6Ago = "o" Then
					'	Response.Write "checked"
					'End If
					'Response.Write ">6��������������"
				%>
				&nbsp;
                * ä�α��� :
                <% drawSellChannelComboBox "sellchnl",sellchnl %>
        	    &nbsp;
				* �ֹ����� :
				<select name="isBanpum" class="select">
					<option value="all" <%=CHKIIF(vIsBanPum="all","selected","")%>>��ǰ����</option>
					<option value="<>" <%=CHKIIF(vIsBanPum="<>","selected","")%>>��ǰ����</option>
					<option value="=" <%=CHKIIF(vIsBanPum="=","selected","")%>>��ǰ�Ǹ�</option>
				</select>
			</td>
		</tr>
		<tr>
		    <td>
				* ���Ա��� :
				<% Call DrawBrandMWUPCombo("mwdiv",mwdiv) %>
        	    &nbsp;
        	    * �������� : 
				<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
        	    &nbsp;
				* ����Ʈ:
				&nbsp;
				* �귣�� : <% drawSelectBoxDesigner "makerid",vmakerid %>
				&nbsp;
				* ����ī�װ� :
				<!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->
			</td>
		 </tr>
		 <tr>
			<td>
				* ����Ʈ���� : <% Call Drawsitename("sitename", vSiteName) %>
				&nbsp;
				* ����ó :
        	    <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
        	    &nbsp;
		        <label><input type="checkbox" name="chkShowGubun" value="Y" <% if (chkShowGubun = "Y") then %>checked<% end if %> > ä�α���,���Ա��� ǥ��</label>
		        <!--&nbsp;* ��ǰ�ڵ� : <textarea rows="3" cols="10" name="itemid" id="itemid"><%'=replace(itemid,",",chr(10))%></textarea>-->
			    &nbsp;
			    <label><input type="checkbox" name="showsuply" value="on" <%= CHKIIF(showsuply="on","checked","") %> >���ް��� ǥ��</label>
			    &nbsp;&nbsp;
			    <label><input type="checkbox" name="incStockAvg" <%=CHKIIF(incStockAvg<>"","checked","")%>>��ո��԰�����</label>
				&nbsp;&nbsp;
			    <label><input type="checkbox" name="isSendGift" value="Y" <%=CHKIIF(isSendGift<>"","checked","")%>>�����ֹ��� ����</label>
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
* �˻� �Ⱓ�� ������� ����� �������ϴ�. �׷��� �˻� ��ư�� Ŭ���� �� �ƹ� ������ ����δٰ� ���� �˻���ư�� Ŭ������ ������.<br>
* ��ۺ� ���� ����
<br>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
	<% if (chkShowGubun = "Y") then %>
		<td align="center" onClick="jstrSort('beadaldiv'); return false;" style="cursor:hand;">
			ä��<br>����
			<img src="/images/list_lineup<%=CHKIIF(vSorting="beadaldivD","_bot","_top")%><%=CHKIIF(instr(vSorting,"beadaldiv")>0,"_on","")%>.png" id="imgbeadaldiv">
		</td>
		<td align="center" onClick="jstrSort('omwdiv'); return false;" style="cursor:hand;">
			����<br>����
			<img src="/images/list_lineup<%=CHKIIF(vSorting="omwdivD","_bot","_top")%><%=CHKIIF(instr(vSorting,"omwdiv")>0,"_on","")%>.png" id="imgomwdiv">
		</td>
	<% end if %>

	<td align="center" colspan="2" onClick="jstrSort('yyyymmdd'); return false;" style="cursor:hand;">
		�Ⱓ
		<img src="/images/list_lineup<%=CHKIIF(vSorting="yyyymmddD","_bot","_top")%><%=CHKIIF(instr(vSorting,"yyyymmdd")>0,"_on","")%>.png" id="imgyyyymmdd">
	</td>
    <td align="center" onClick="jstrSort('countOrder'); return false;" style="cursor:hand;">
    	�ֹ���
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="countOrderD","_bot","_top")%><%=CHKIIF(instr(vSorting,"countOrder")>0,"_on","")%>.png" id="imgcountOrder">
    </td>
    <td align="center" onClick="jstrSort('itemno'); return false;" style="cursor:hand;">
    	�Ǹż���
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="itemnoD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemno")>0,"_on","")%>.png" id="imgitemno">
    </td>

    <% if (NOT C_InspectorUser) then %>
	    <td align="center" onClick="jstrSort('orgitemcost'); return false;" style="cursor:hand;">
	    	�Һ��ڰ�[��ǰ]
	    	<img src="/images/list_lineup<%=CHKIIF(vSorting="orgitemcostD","_bot","_top")%><%=CHKIIF(instr(vSorting,"orgitemcost")>0,"_on","")%>.png" id="imgorgitemcost">
	    </td>
	    <td align="center" onClick="jstrSort('itemcostcouponnotapplied'); return false;" style="cursor:hand;">
	    	�ǸŰ�[��ǰ]<br>(��������)
	    	<img src="/images/list_lineup<%=CHKIIF(vSorting="itemcostcouponnotappliedD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemcostcouponnotapplied")>0,"_on","")%>.png" id="imgitemcostcouponnotapplied">
	    </td>
	    <td align="center" onClick="jstrSort('itemcost1'); return false;" style="cursor:hand;">
	    	�����Ѿ�[��ǰ]<br>(��ǰ��������)
	    	<img src="/images/list_lineup<%=CHKIIF(vSorting="itemcost1D","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemcost1")>0,"_on","")%>.png" id="imgitemcost1">
	    </td>
	    <td align="center" onClick="jstrSort('itemcostnotreducedprice'); return false;" style="cursor:hand;">
	    	���ʽ�����<br>����[��ǰ]
	    	<img src="/images/list_lineup<%=CHKIIF(vSorting="itemcostnotreducedpriceD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemcostnotreducedprice")>0,"_on","")%>.png" id="imgitemcostnotreducedprice">
	    </td>
    <% end if %>

    <td align="center" onClick="jstrSort('reducedPrice'); return false;" style="cursor:hand;">
    	��޾�
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="reducedPriceD","_bot","_top")%><%=CHKIIF(instr(vSorting,"reducedPrice")>0,"_on","")%>.png" id="imgreducedPrice">
    </td>
	<td align="center" onClick="jstrSort('upchejungsan1'); return false;" style="cursor:hand;">
		��ü<br>�����
		<img src="/images/list_lineup<%=CHKIIF(vSorting="upchejungsan1D","_bot","_top")%><%=CHKIIF(instr(vSorting,"upchejungsan1")>0,"_on","")%>.png" id="imgupchejungsan1">
	</td>
	<td align="center" onClick="jstrSort('reducedpricenotupchejungsan'); return false;" style="cursor:hand;">
		<b>ȸ�����</b>
		<img src="/images/list_lineup<%=CHKIIF(vSorting="reducedpricenotupchejungsanD","_bot","_top")%><%=CHKIIF(instr(vSorting,"reducedpricenotupchejungsan")>0,"_on","")%>.png" id="imgreducedpricenotupchejungsan">
	</td>
	<td align="center" onClick="jstrSort('avgipgoprice'); return false;" style="cursor:hand;">
		���<br>���԰�
		<img src="/images/list_lineup<%=CHKIIF(vSorting="avgipgopriceD","_bot","_top")%><%=CHKIIF(instr(vSorting,"avgipgoprice")>0,"_on","")%>.png" id="imgavgipgoprice">
	</td>
	<td align="center" onClick="jstrSort('overvaluestockprice'); return false;" style="cursor:hand;">
		���<br>����
		<img src="/images/list_lineup<%=CHKIIF(vSorting="overvaluestockpriceD","_bot","_top")%><%=CHKIIF(instr(vSorting,"overvaluestockprice")>0,"_on","")%>.png" id="imgovervaluestockprice">
	</td>
    <td align="center" onClick="jstrSort('buycash'); return false;" style="cursor:hand;">
    	�����Ѿ�[��ǰ]<% if (NOT C_InspectorUser) then %><br>(��ǰ��������)<% end if %>
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="buycashD","_bot","_top")%><%=CHKIIF(instr(vSorting,"buycash")>0,"_on","")%>.png" id="imgbuycash">
    </td>
    <td align="center" onClick="jstrSort('maechulprofit1'); return false;" style="cursor:hand;">
    	<b>�������</b>
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="maechulprofit1D","_bot","_top")%><%=CHKIIF(instr(vSorting,"maechulprofit1")>0,"_on","")%>.png" id="imgmaechulprofit1">
    </td>
    <td align="center">������</td>
    <td align="center" onClick="jstrSort('maechulprofit2'); return false;" style="cursor:hand;">
    	�������2<br>(��޾ױ���)
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="maechulprofit2D","_bot","_top")%><%=CHKIIF(instr(vSorting,"maechulprofit2")>0,"_on","")%>.png" id="imgmaechulprofit2">
    </td>
    <td align="center">������</td>
    <td align="center">���</td>
</tr>
<% if cStatistic.FTotalCount > 0 then %>
<%
	For i = 0 To cStatistic.FTotalCount -1
%>
<tr bgcolor="#FFFFFF">
	<% if (chkShowGubun = "Y") then %>
	<td align="center"><%= getSellChannelName(cStatistic.flist(i).Fbeadaldiv) %></td>
	<td align="center"><%= cStatistic.flist(i).Fomwdiv %></td>
	<% end if %>
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
	<td align="center"><%= CDbl(cStatistic.FList(i).FcountOrder) %></td>
	<td align="center"><%= CDbl(cStatistic.FList(i).FItemNO) %></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FOrgitemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FItemcostCouponNotApplied) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FItemCost) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FItemCost-cStatistic.FList(i).FReducedPrice) %></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice) %></td>

	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FupcheJungsan) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><b><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice - cStatistic.FList(i).FupcheJungsan) %></b></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FavgipgoPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FoverValueStockPrice) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FBuyCash) %></td>
	<td align="right" style="padding-right:5px;"><b><%= NullOrCurrFormat(cStatistic.FList(i).FMaechulProfit) %></b></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).FMaechulProfitPer %>%</td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).FReducedPrice-cStatistic.FList(i).FBuyCash) %></td>
	<td align="right" style="padding-right:5px;"><%= cStatistic.FList(i).FMaechulProfitPer2 %>%</td>
	<td align="center" >[<a href="/admin/upchejungsan/upcheselllist.asp?datetype=jumunil&yyyy1=<%=Year(cStatistic.FList(i).FRegdate)%>&mm1=<%=TwoNumber(Month(cStatistic.FList(i).FRegdate))%>&dd1=<%=TwoNumber(Day(cStatistic.FList(i).FRegdate))%>&yyyy2=<%=Year(cStatistic.FList(i).FRegdate)%>&mm2=<%=TwoNumber(Month(cStatistic.FList(i).FRegdate))%>&dd2=<%=TwoNumber(Day(cStatistic.FList(i).FRegdate))%>&disp=<%=dispCate%>&delivertype=all&inc3pl=<%= inc3pl %>&isSendGift=<%=isSendGift%>" target="_blank">��</a>]</td>
</tr>
<%
	vTot_countOrder					= vTot_countOrder + CLng(NullOrCurrFormat(cStatistic.FList(i).FcountOrder))
	vTot_ItemNO						= vTot_ItemNO + CLng(NullOrCurrFormat(cStatistic.FList(i).FItemNO))
	vTot_OrgitemCost				= vTot_OrgitemCost + CDbl(NullOrCurrFormat(cStatistic.FList(i).FOrgitemCost))
	vTot_ItemcostCouponNotApplied	= vTot_ItemcostCouponNotApplied + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemcostCouponNotApplied))
	vTot_ItemCost					= vTot_ItemCost + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemCost))
	vTot_BonusCouponPrice			= vTot_BonusCouponPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FItemCost-cStatistic.FList(i).FReducedPrice))
	vTot_ReducedPrice				= vTot_ReducedPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FReducedPrice))
	vTot_BuyCash					= vTot_BuyCash + CDbl(NullOrCurrFormat(cStatistic.FList(i).FBuyCash))
	vTot_MaechulProfit				= vTot_MaechulProfit + CDbl(NullOrCurrFormat(cStatistic.FList(i).FMaechulProfit))
	vTot_MaechulProfit2				= vTot_MaechulProfit2 + CDbl(NullOrCurrFormat(cStatistic.FList(i).FReducedPrice-cStatistic.FList(i).FBuyCash))

	vTot_upcheJungsan				= vTot_upcheJungsan + CDbl(NullOrCurrFormat(cStatistic.FList(i).FupcheJungsan))
	vTot_avgipgoPrice				= vTot_avgipgoPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FavgipgoPrice))
	vTot_overValueStockPrice		= vTot_overValueStockPrice + CDbl(NullOrCurrFormat(cStatistic.FList(i).FoverValueStockPrice))

	Next

	vTot_MaechulProfitPer = Round(((vTot_ItemCost - vTot_BuyCash)/CHKIIF(vTot_ItemCost=0,1,vTot_ItemCost))*100,2)
	vTot_MaechulProfitPer2 = Round(((vTot_ReducedPrice - vTot_BuyCash)/CHKIIF(vTot_ReducedPrice=0,1,vTot_ReducedPrice))*100,2)
%>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<% if (chkShowGubun = "Y") then %>
	<td colspan="2"></td>
	<% end if %>
	<td align="center" colspan="2">�Ѱ�</td>
	<td align="center"><%=vTot_countOrder%></td>
	<td align="center"><%=vTot_ItemNO%></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_OrgitemCost)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_ItemcostCouponNotApplied)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_ItemCost)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_BonusCouponPrice)%></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_ReducedPrice)%></td>

	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_upcheJungsan)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_ReducedPrice - vTot_upcheJungsan)%></b></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_avgipgoPrice)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_overValueStockPrice)%></td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_BuyCash)%></td>
	<td align="right" style="padding-right:5px;"><b><%=NullOrCurrFormat(vTot_MaechulProfit)%></b></td>
	<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer%>%</td>
	<td align="right" style="padding-right:5px;"><%=NullOrCurrFormat(vTot_MaechulProfit2)%></td>
	<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer2%>%</td>
	<td></td>
</tr>
<% ELSE %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="26">��ϵ� ������ �����ϴ�.</td>
</tr>
<%END IF%>
</table>

<%
Set cStatistic = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
