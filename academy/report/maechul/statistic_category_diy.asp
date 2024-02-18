<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  핑거스 카테고리별매출
' History : 2016.06.10 한용민 생성
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
	cStatistic.FRectIncStockAvgPrc = true '' 평균매입가 포함 쿼리여부.
	cStatistic.FRectSort = vSorting

	if (categbn="M") then
	    cStatistic.fStatistic_diy_category()
	else
	    cStatistic.FRectdispCate = dispCate
        cStatistic.FRectmaxDepth = maxdepth   
    	cStatistic.fStatistic_diy_DispCategory  ''2013/10/17 추가
    end if

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>
<script type="text/javascript">

function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}

function popCateSellDetail(cdl,cdm,cds,dispcate){
    window.open("/academy/report/maechul/statistic_category_diy.asp?menupos=<%= menupos %>&date_gijun=<%=vDateGijun%>&syear=<%=vSYear%>&smonth=<%=vSMonth%>&sday=<%=vSDay%>&eyear=<%=vEYear%>&emonth=<%=vEMonth%>&eday=<%=vEDay%>&cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&disp="+dispcate,'','');
}

function MonthDiff(d1, d2) {
	d1 = d1.split("-");
	d2 = d2.split("-");

	d1 = new Date(d1[0], d1[1] - 1, d1[2]);
	d2 = new Date(d2[0], d2[1] - 1, d2[2]);

	var d1Y = d1.getFullYear();
	var d2Y = d2.getFullYear();
	var d1M = d1.getMonth();
	var d2M = d2.getMonth();

	return (d2M+12*d2Y)-(d1M+12*d1Y);
}

function searchSubmit(){
	if ((CheckDateValid(frm.syear.value, frm.smonth.value, frm.sday.value) == true) && (CheckDateValid(frm.eyear.value, frm.emonth.value, frm.eday.value) == true)) {
		$("#btnSubmit").prop("disabled", true);
		frm.submit();
	}
}

function jsChangeDepth(ivalue){
    var dispDepth  = "<%=maxDepth%>";
    var strDisp = 0;
   
    if(ivalue < dispDepth){ 
        if (ivalue == 0){
            strDisp = "";
        }else{ 
         strDisp = "<%=dispCate%>".substring(0,ivalue*3);
        }
    
        document.all.disp.value =strDisp ;
    }
    searchSubmit(); 
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

function downloadexcel(){
    document.frm.target = "view"; 
    document.frm.action = "/academy/report/maechul/statistic_category_diy_excel.asp";  
	document.frm.submit();
    document.frm.target = ""; 
    document.frm.action = "";  
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="sorting" value="<%= vsorting %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				* 기간 : 
				<select name="date_gijun" class="select">
					<option value="regdate" <%=CHKIIF(vDateGijun="regdate","selected","")%>>주문일</option>
					<option value="ipkumdate" <%=CHKIIF(vDateGijun="ipkumdate","selected","")%>>결제일</option>
					<option value="beasongdate" <%=CHKIIF(vDateGijun="beasongdate","selected","")%>>상품출고일</option>
				</select>
				<% DrawDateBoxdynamic vSYear,"syear",vEYear,"eyear",vSMonth,"smonth",vEMonth,"emonth",vSDay,"sday",vEDay,"eday" %>
			</td>
		</tr>
		<tr>
			<td height="25">
				* 채널구분
            	<% drawSelectBox_SellChannel "sellchnl", sellchnl, "" %>
                &nbsp;
                * 주문구분 :
				<select name="isBanpum" class="select">
					<option value="all" <%=CHKIIF(vIsBanPum="all","selected","")%>>반품포함</option>
					<option value="<>" <%=CHKIIF(vIsBanPum="<>","selected","")%>>반품제외</option>
					<option value="=" <%=CHKIIF(vIsBanPum="=","selected","")%>>반품건만</option>
				</select>
				&nbsp;
				* 매입구분 :
				<% Call DrawBrandMWUPCombo("mwdiv",mwdiv) %>
			</td>
		</tr>
		<tr>
		    <td>
				* 브랜드 : <input type="text" class="text" name="ebrand" value="<%=vBrandID%>" size="20"> <input type="button" class="button" value="IDSearch" onclick="jsSearchBrandID(this.form.name,'ebrand');">
				&nbsp;
				<input type="radio" name="categbn" value="M" <%=CHKIIF(categbn="M","checked","")%> >관리카테고리
				<input type="radio" name="categbn" value="D" <%=CHKIIF(categbn="D","checked","")%> >전시카테고리
				<select name="selDepth" class="select"  onChange="jsChangeDepth(this.value);" <%if categbn = "M" then%>disabled<%end if%>>
				    <option value="0" <%if maxDepth ="0" then%>selected<%end if%>>대(1 Depth)</option>
				    <option value="1" <%if maxDepth ="1" then%>selected<%end if%>>중(2 Depth)</option>
				    <option value="2" <%if maxDepth ="2" then%>selected<%end if%>>소(3 Depth)</option>
				    <option value="3" <%if maxDepth ="3" then%>selected<%end if%>>세(4 Depth)</option>
				</select> 

				<%if categbn = "M" then %>
					<!-- #include virtual="/academy/comm/CategorySelectBox.asp"-->
				<% end if%>
				<%if maxDepth > 0 and categbn = "D" then %>
					<!-- #include virtual="/academy/comm/dispCateSelectBoxDepth.asp"-->
				<% end if%>
			</td>
		</tr>
	    </table>
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" id="btnSubmit" class="button_s" value="검색" onClick="javascript:searchSubmit();" ></td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		* 검색 기간이 길어지면 상당히 느려집니다. 그러니 검색 버튼을 클릭한 뒤 아무 반응이 없어보인다고 재차 검색버튼을 클릭하지 마세요.
	</td>
	<td align="right">	
		<input type="button" onclick="downloadexcel();" value="엑셀다운로드" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		검색결과 : <b><%= cStatistic.FTotalCount %></b>
	</td> 
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>"  align="center">
	<td align="center" onClick="jstrSort('categoryname'); return false;" style="cursor:hand;">
		<%=CateGubun(vCateGubun)%>카테고리
		<img src="/images/list_lineup<%=CHKIIF(vSorting="categorynameD","_bot","_top")%><%=CHKIIF(instr(vSorting,"categoryname")>0,"_on","")%>.png" id="imgcategoryname">
	</td>
	<td onClick="jstrSort('itemno'); return false;" style="cursor:hand;">
		상품수량
		<img src="/images/list_lineup<%=CHKIIF(vSorting="itemnoD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemno")>0,"_on","")%>.png" id="imgitemno">
	</td>

    <% if (NOT C_InspectorUser) then %>
    	<td onClick="jstrSort('couponnotasigncost'); return false;" style="cursor:hand;">
    		판매가[상품]<br>(할인적용)
    		<img src="/images/list_lineup<%=CHKIIF(vSorting="couponnotasigncostD","_bot","_top")%><%=CHKIIF(instr(vSorting,"couponnotasigncost")>0,"_on","")%>.png" id="imgcouponnotasigncost">
    	</td>
    	<td onClick="jstrSort('itemcost'); return false;" style="cursor:hand;">
    		<b>구매총액[상품]<br>(상품쿠폰적용)</b>
    		<img src="/images/list_lineup<%=CHKIIF(vSorting="itemcostD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemcost")>0,"_on","")%>.png" id="imgitemcost">
    	</td>
    	<td onClick="jstrSort('itemCostnotexistsbonus'); return false;" style="cursor:hand;">
    		<b>보너스쿠폰<br>사용액[상품]</b>
    		<img src="/images/list_lineup<%=CHKIIF(vSorting="itemCostnotexistsbonusD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemCostnotexistsbonus")>0,"_on","")%>.png" id="imgitemCostnotexistsbonus">
    	</td>
    <% end if %>

    <td onClick="jstrSort('reducedprice'); return false;" style="cursor:hand;">
    	취급액
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="reducedpriceD","_bot","_top")%><%=CHKIIF(instr(vSorting,"reducedprice")>0,"_on","")%>.png" id="imgreducedprice">
    </td>
    <td onClick="jstrSort('buycash'); return false;" style="cursor:hand;">
    	매입총액[상품]<% if (NOT C_InspectorUser) then %><br>(상품쿠폰적용)<% end if %>
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="buycashD","_bot","_top")%><%=CHKIIF(instr(vSorting,"buycash")>0,"_on","")%>.png" id="imgbuycash">
    </td>
    <td onClick="jstrSort('maechulprofit1'); return false;" style="cursor:hand;">
    	<b>매출수익</b>
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="maechulprofit1D","_bot","_top")%><%=CHKIIF(instr(vSorting,"maechulprofit1")>0,"_on","")%>.png" id="imgmaechulprofit1">
    </td>
    <td onClick="jstrSort('maechulprofitper1'); return false;" style="cursor:hand;">
    	수익율1
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="maechulprofitper1D","_bot","_top")%><%=CHKIIF(instr(vSorting,"maechulprofitper1")>0,"_on","")%>.png" id="imgmaechulprofitper1">
    </td>
    <td onClick="jstrSort('maechulprofit2'); return false;" style="cursor:hand;">
    	매출수익2<br>(취급액기준)
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="maechulprofit2D","_bot","_top")%><%=CHKIIF(instr(vSorting,"maechulprofit2")>0,"_on","")%>.png" id="imgmaechulprofit2">
    </td>
    <td onClick="jstrSort('maechulprofitper2'); return false;" style="cursor:hand;">
    	수익율2
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="maechulprofitper2D","_bot","_top")%><%=CHKIIF(instr(vSorting,"maechulprofitper2")>0,"_on","")%>.png" id="imgmaechulprofitper2">
    </td>
    <td>
    	카테고리별<br>매출 비중
    </td>
	<td onClick="jstrSort('upchejungsan'); return false;" style="cursor:hand;" align="center">
		업체<br>정산액
		<img src="/images/list_lineup<%=CHKIIF(vSorting="upchejungsanD","_bot","_top")%><%=CHKIIF(instr(vSorting,"upchejungsan")>0,"_on","")%>.png" id="imgupchejungsan">
	</td>
	<td onClick="jstrSort('reducedpricenotexistsupchejungsan'); return false;" style="cursor:hand;" align="center">
		<b>회계매출</b>
		<img src="/images/list_lineup<%=CHKIIF(vSorting="reducedpricenotexistsupchejungsanD","_bot","_top")%><%=CHKIIF(instr(vSorting,"reducedpricenotexistsupchejungsan")>0,"_on","")%>.png" id="imgreducedpricenotexistsupchejungsan">
	</td>
    <td align="center">
    	비고
    </td>
</tr>
<% if cStatistic.FTotalCount > 0 then %>
<%
For i = 0 To cStatistic.FTotalCount -1
%>
<tr bgcolor="#FFFFFF">
	<td  style="padding-left:5px;">
		<%= cStatistic.FItemList(i).FCategoryName %>&nbsp;
		<%  linkcate = ""
			If vCateGubun = "L" Then
				Response.Write "<a href="""&vParam&"&cdl="&cStatistic.FItemList(i).FCateL&"""><font color=""gray"">[중]</font></a>"
				IF (cStatistic.FItemList(i).FCateL="999") then
				    '' Response.Write " <a href=""javascript:popCateSellDetail('"&cStatistic.FItemList(i).FCateL&"','','','')"">(상세)</a>"
				end if
				if categbn = "D" then
				    linkcate = "&disp1="&cStatistic.FItemList(i).FCateL
				else    
				    linkcate = "&cdl="&cStatistic.FItemList(i).FCateL
				end if
			ElseIf vCateGubun = "M" Then
				Response.Write "<a href="""&vParam&"""><font color=""gray"">[대]</font></a>"
				Response.Write "<a href="""&vParam&"&cdl="&cStatistic.FItemList(i).FCateL&"&cdm="&cStatistic.FItemList(i).FCateM&"""><font color=""gray"">[소]</font></a>"
				IF (cStatistic.FItemList(i).FCateM="") then
				    Response.Write " <a href=""javascript:popCateSellDetail('"&cStatistic.FItemList(i).FCateL&"','','','')"">(상세)</a>"
				end if
				if categbn = "D" then
				    linkcate = "&disp1="&cStatistic.FItemList(i).FCateL&"&disp2="&cStatistic.FItemList(i).FCateM
			    else    
				    linkcate = "&cdl="&cStatistic.FItemList(i).FCateL&"&cdm="&cStatistic.FItemList(i).FCateM
			    end if
				
			ElseIf vCateGubun = "S" Then
				Response.Write "<a href="""&vParam&"""><font color=""gray"">[대]</font></a>"
				Response.Write "<a href="""&vParam&"&cdl="&cStatistic.FItemList(i).FCateL&"""><font color=""gray"">[중]</font></a>"
				if (categbn="D") then
                Response.Write "<a href="""&vParam&"&cdl="&cStatistic.FItemList(i).FCateL&"&cdm="&cStatistic.FItemList(i).FCateM&"&cds="&cStatistic.FItemList(i).FCateS&"""><font color=""gray"">[세]</font></a>"
                    linkcate = "&disp1="&cStatistic.FItemList(i).FCateL&"&disp2="&cStatistic.FItemList(i).FCateM&"&disp3="&cStatistic.FItemList(i).FCateS
                else
                    linkcate = "&cdl="&cStatistic.FItemList(i).FCateL&"&cdm="&cStatistic.FItemList(i).FCateM&"&cds="&cStatistic.FItemList(i).FCateS
                end if 
            ElseIf vCateGubun = "X" Then
				Response.Write "<a href="""&vParam&"""><font color=""gray"">[대]</font></a>"
				Response.Write "<a href="""&vParam&"&cdl="&cStatistic.FItemList(i).FCateL&"""><font color=""gray"">[중]</font></a>"
                Response.Write "<a href="""&vParam&"&cdl="&cStatistic.FItemList(i).FCateL&"&cdm="&cStatistic.FItemList(i).FCateM&"""><font color=""gray"">[소]</font></a>"
                'Response.Write " <a href=""javascript:popCateSellDetail('"&cStatistic.FItemList(i).FCateL&"','"&cStatistic.FItemList(i).FCateM&"','"&cStatistic.FItemList(i).FCateS&"','"&cStatistic.FItemList(i).FCateX&"')"">(상세)</a>"
             End IF
              linkdispcate =  "&disp="&cStatistic.FItemList(i).FDispCateCode 
			if cStatistic.FTotItemCost ="" or cStatistic.FTotItemCost = 0 then
				vCateMRate = 0
			else
				vCateMRate = (cStatistic.FItemList(i).FItemCost/cStatistic.FTotItemCost)*100
			end if
	' Response.Write " <a href=""javascript:popCateSellDetail('"&cStatistic.FItemList(i).FCateL&"','"&cStatistic.FItemList(i).FCateM&"','"&cStatistic.FItemList(i).FCateS&"','"&cStatistic.FItemList(i).FCateX&"')"">(상세)</a>"
		%>
	</td>
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
	<td align="center"><a href="/academy/report/maechul/statistic_item.asp?menupos=<%= menupos %>&sitename=diyitem&date_gijun=<%=vDateGijun%>&syear=<%=vSYear%>&smonth=<%=vSMonth%>&sday=<%=vSDay%>&eyear=<%=vEYear%>&emonth=<%=vEMonth%>&eday=<%=vEDay%><%=linkcate&linkdispcate%>" target="_blank">[상품상세]</a></td>
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
	<td align="center">총계</td>
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
	<td></td>
</tr>
<% ELSE %>
	<tr  align="center" bgcolor="#FFFFFF">
		<td colspan="25">등록된 내용이 없습니다.</td>
	</tr>
<% end if %>
</table>
<script>
<!--
FusionCharts.ready(function () {
    var revenueChart = new FusionCharts({
        type: 'doughnut2d',
        renderAt: 'chart-container',
        width: '450',
        height: '450',
        dataFormat: 'json',
        dataSource: {
            "chart": {
                "caption": "카테고리별매출",
                "subCaption": "<%=vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay) & " ~ " & vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay) %>",
                "numberPrefix": "",
                "paletteColors": "#0075c2,#1aaf5d,#f2c500,#f45b00,#8e0000",
                "bgColor": "#ffffff",
                "showBorder": "0",
                "use3DLighting": "0",
				"formatNumberScale": "0",
                "showShadow": "0",
                "enableSmartLabels": "0",
                "startingAngle": "310",
                "showLabels": "0",
                "showPercentValues": "1",
                "showLegend": "1",
                "legendShadow": "0",
                "legendBorderAlpha": "0",
                "defaultCenterLabel": "Total revenue: <%=FormatNumber(vTot_ItemCost,0)%>",
                "centerLabel": "Revenue from $label: $value",
                "centerLabelBold": "1",
                "showTooltip": "1",
                "decimals": "0",
                "captionFontSize": "14",
                "subcaptionFontSize": "10",
                "subcaptionFontBold": "0"
            },
            "data": [
				<%
				if cStatistic.FTotalCount > 0 then
					For i = 0 To cStatistic.FTotalCount -1
						Response.Write "{" & vbCrLf
						Response.Write """label"": """ & cStatistic.FItemList(i).FCategoryName & """," & vbCrLf
						Response.Write """value"": """ & cStatistic.FItemList(i).FItemCost & """" & vbCrLf
						Response.Write "}"
						If i <> cStatistic.FTotalCount-1 Then
							Response.Write ","
						End If
						Response.Write vbCrLf
					Next
				End If
				%>
            ]
        }
    }).render();
});
//-->
</script>
<div id="chart-container">FusionCharts will render here</div>
<% Set cStatistic = Nothing

Function CateGubun(g)
	If g = "L" Then
		CateGubun = "대"
	ElseIf vCateGubun = "M" Then
		CateGubun = "중"
	ElseIf vCateGubun = "S" Then
		CateGubun = "소"
	End IF
End Function
%>
<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
