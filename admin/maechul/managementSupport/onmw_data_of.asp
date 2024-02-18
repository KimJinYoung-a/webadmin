<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp" -->
<%

'// 기타매출 서머리 재작성
'// [db_datamart].[dbo].[ten_Datamart_ManagementSupportTeam_Daily_totalsale_ETC] '2013-03','2013-03'

const jErrShow = true
	dim bancancle,accountdiv,sitename,ipkumdatesucc, vPurchasetype, vatinclude
	dim yyyy1,yyyy2,mm1,mm2
	dim i ,defaultdate,defaultdate1 , olddata
    dim mdgbn, targetGbn, dlvdiv, vbizsec
    dim supptype
    
	sitename = request("sitename")
	accountdiv = request("accountdiv")
	vPurchasetype = request("purchasetype")
	bancancle = NullFillWith(request("bancancle"), "1")
	vatinclude = request("vatinclude")
    supptype = request("supptype")
    
	defaultdate1 = dateadd("d",-60,year(now) & "-" &month(now) & "-" & day(now))		'날짜값이 없을때 기본값으로 60이전까지 검색
	yyyy1 = NullFillWith(request("yyyy1"), left(defaultdate1,4))
	mm1 = NullFillWith(request("mm1"), mid(defaultdate1,6,2))
	yyyy2 = NullFillWith(request("yyyy2"), year(now))
	mm2 = NullFillWith(request("mm2"), month(now))
	mm2 = TwoNumber(mm2)
    mdgbn = NullFillWith(request("mdgbn"),"m")
    targetGbn = NullFillWith(request("targetGbn"),"")
    dlvdiv = NullFillWith(request("dlvdiv"),"")
    vbizsec = NullFillWith(request("bizsec"),"")

	dim Omaechul_list
	set Omaechul_list = new cManagementSupportMaechul_list
	Omaechul_list.FRectStartdate = yyyy1 & "-" & mm1 & "-01"
	Omaechul_list.FRectEndDate = CStr(DateAdd("d",-1,DateAdd("m",1,yyyy2 & "-" & mm2 & "-01")))
	Omaechul_list.frectbancancle = bancancle
	Omaechul_list.frectaccountdiv = accountdiv
	Omaechul_list.frectsitename = sitename
	Omaechul_list.frectipkumdatesucc = ipkumdatesucc
	Omaechul_list.frectpurchasetype = vPurchasetype
	Omaechul_list.frectvatinclude = vatinclude

	Omaechul_list.FRectOnOff = targetGbn
	Omaechul_list.FRectDLVdiv = dlvdiv
	Omaechul_list.frectGroupByMwDiv="on"
	Omaechul_list.frectGroupByMonth=mdgbn
	Omaechul_list.FRectBizSectionCd=vbizsec
    Omaechul_list.FRectSupptype = supptype
	Omaechul_list.fmaechul_listOFByGbn()

	Dim vSum_TotItemNo, vSum_TotReducedPrice, vSum_TotBuycash, vSum_TotBuycashCouponNotApplied
	Dim vSum_TotOrgitemcost, vSum_TotItemcostCouponNotApplied,  vSum_TotItemcost
	Dim vSum_TotReducePrice, vSum_SpendCouponSum, vSum_MaechulItem
	Dim vSum_SpendMileSum
	Dim vSum_jPrice,vSum_jPriceEtc,vSum_jPriceEtcChulgo
    Dim vSum_HanDlePrice , vSum_CalcuMeachul, vSum_CalcuMeachulNoVat, vSum_ErrJungsan

%>
<h3>수정중</h3>
<!-- 검색 시작 -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">상품출고일 / 날짜 <% DrawYMYMBox yyyy1,mm1,yyyy2,mm2 %>
			&nbsp;&nbsp;&nbsp;
			<!--
			<input type="radio" name="mdgbn" value="m" <%= CHKIIF(mdgbn="m","checked","") %> >월별
			<input type="radio" name="mdgbn" value="d" <%= CHKIIF(mdgbn="d","checked","") %> disabled >일별
			-->
			* 기본 매출부서 :
			<% Call DrawBizSectionGain("F,T","bizsec", vbizsec,"") %>

			&nbsp;&nbsp;&nbsp;
			상품귀속
			<select name="targetGbn">
			<option value="" <%=CHKIIF(targetGbn="","selected","") %> >전체
			<option value="OF" <%=CHKIIF(targetGbn="OF","selected","") %> >오프라인(OF)
			<option value="IT" <%=CHKIIF(targetGbn="IT","selected","") %> >아이띵소_오프라인(IT)
			</select>

			&nbsp;&nbsp;&nbsp;
			매입(정산)구분
			<select name="dlvdiv">
			<option value="" <%=CHKIIF(dlvdiv="","selected","") %> >전체
			<option value="s" <%=CHKIIF(dlvdiv="s","selected","") %> >업체특정제외
			<option value="d" <%=CHKIIF(dlvdiv="d","selected","") %> >업체특정
			<option value="B000" <%=CHKIIF(dlvdiv="M","selected","") %> >미지정
			<option value="B011" <%=CHKIIF(dlvdiv="W","selected","") %> >특정판매
			<option value="B012" <%=CHKIIF(dlvdiv="U","selected","") %> >업체특정
			<option value="B013" <%=CHKIIF(dlvdiv="Y","selected","") %> >출고특정
			<option value="B021" <%=CHKIIF(dlvdiv="Z","selected","") %> >오프매입
			<option value="B022" <%=CHKIIF(dlvdiv="Z","selected","") %> >매장매입
			<option value="B023" <%=CHKIIF(dlvdiv="Z","selected","") %> >가맹점매입
	        <option value="B031" <%=CHKIIF(dlvdiv="Z","selected","") %> >출고매입
	        <option value="B032" <%=CHKIIF(dlvdiv="Z","selected","") %> >센터매입
	        <option value="B999" <%=CHKIIF(dlvdiv="Z","selected","") %> >기타보정
			</select>
			</td>
		</tr>
    	<tr>
    		<td height="25">

	    	<input type=radio name="bancancle" value="1" <% if bancancle="1" then  response.write "checked" %>>반품포함
	    	<input type=radio name="bancancle" value="2" <% if bancancle="2" then  response.write "checked" %>>반품건만
	    	<input type=radio name="bancancle" value="3" <% if bancancle="3" then  response.write "checked" %>>반품제외
	    	<!--
	    	/ 결제구분 <select name="accountdiv">
	    		<option value="" <% if accountdiv = "" then response.write "selected" %>>전체</option>
	    		<option value="7" <% if accountdiv = "7" then response.write "selected" %>>무통장</option>
				<option value="14" <% if accountdiv = "14" then response.write "selected" %>>편의점결제</option>
	    		<option value="20" <% if accountdiv = "20" then response.write "selected" %>>실시간</option>
	    		<option value="50" <% if accountdiv = "50" then response.write "selected" %>>외부몰</option>
	    		<option value="80" <% if accountdiv = "80" then response.write "selected" %>>올엣</option>
	    		<option value="100" <% if accountdiv = "100" then response.write "selected" %>>신용카드</option>
	    	</select>
	    	-->
	    	&nbsp;&nbsp;&nbsp;
	    	/ 과세구분 <select name="vatinclude">
	    	    <option value="" <% if vatinclude = "" then response.write "selected" %>>전체</option>
	    		<option value="Y" <% if vatinclude = "Y" then response.write "selected" %>>과세</option>
	    		<option value="N" <% if vatinclude = "N" then response.write "selected" %>>면세</option>
	    	</select>
	    	&nbsp;&nbsp;&nbsp;

	    	<!--* 출고처구분 : <% 'drawSelectBoxOffShopdiv_off "sitename", sitename, "1,3,7,9", "", " " %>-->
			* 매출처 : <% NewdrawSelectBoxShopAll "sitename", sitename %>
	    	&nbsp;&nbsp;&nbsp;
	    	매출형태 : <% drawPartnerCommCodeBox true,"selljungsantype","purchasetype",vPurchasetype,"" %>
	    	
	    	&nbsp;&nbsp;&nbsp;
	    	/   <input type="radio" name="supptype" value="S" <%= CHKIIF(supptype="S","checked","") %> > 공급가액
	    	    <input type="radio" name="supptype" value="" <%= CHKIIF(supptype="","checked","") %> > 합계금액
	    	</td>
	    </tr>
	    </table>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:submit();">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<br>
<table width="100%" class="a">
<tr bgcolor="#FFFFFF">
    <td>
        매출(M) : 직영 : 업체특정 = 취급액-매입가(수수료), 그외정산조건 = 취급액, 아이띵소 = 취급액,
    </td>
</tr>
</table>

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td rowspan="2" align="center" width="70" >출고<%=CHKIIF(mdgbn="m","월","일") %></td>
	<td rowspan="2" align="center" width="90" >매장</td>
	<td rowspan="2" align="center" width="90" >계정과목</td>
	<td rowspan="2" align="center" width="60" >매출<br>부서</td>
	<td rowspan="2" align="center" width="40" >상품<br>귀속</td>
	<td rowspan="2" align="center" width="40" >매입<br>구분</td>
    <td rowspan="2" align="center" width="50" >총상품<br>갯수</td>
    <td rowspan="2" align="center" >소비자가<br>A</td>
    <td rowspan="2" align="center" >할인금액<br>B</td>
    <td rowspan="2" align="center" >판매가(할인가)<br>C=A-B</td>
    <td rowspan="2" align="center" >POS할인<br>D</td>
    <td rowspan="2" align="center" >구매총액<br>E=C-D</td>
    <td align="center" colspan="2">보너스쿠폰<br>정율쿠폰(F)=E-환불액(reducePrice)<br>정액쿠폰(G)</td>
    <td rowspan="2" align="center" >취급액<br>(H)=E-F-G</td>
    <td rowspan="2" align="center" width="5" ></td>
    <!--<td rowspan="2" align="center" >마일리지<br>사용안분</td>-->
    <td rowspan="2" align="center" >취급액원가<br>(주문시매입가)<br>(S)</td>
    <td rowspan="2" align="center" >취급액<br>원가율(%)<br>S/H</td>
    <td rowspan="2" align="center" >매출(M)</td>
    <td rowspan="2" align="center" >매출<br>(vat제외)</td>
    <td rowspan="2" align="center" width="5" ></td>
    <td rowspan="2" align="center" >정산액<br>(J1)</td>
    <td rowspan="2" align="center" >기타정산<br>(반품배송비등)</td>
    <td rowspan="2" align="center" >기타출고정산<br>(판촉,로스등)</td>
    <% if (jErrShow) then %>
    <td rowspan="2" align="center" >정산오차<br>(S-J1)</td>
    <% end if %>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>정율쿠폰(F)</td>
	<td>정액쿠폰(G)<br>안분</td>
</tr>
<%
Dim vYear, vMonth, vDay
For i = 0 To Omaechul_list.ftotalcount -1
	vYear	= Year(Omaechul_list.flist(i).fbaesongdate)
	vMonth	= TwoNumber(Month(Omaechul_list.flist(i).fbaesongdate))
	vDay	= TwoNumber(Day(Omaechul_list.flist(i).fbaesongdate))
%>
<tr align="center" bgcolor="#FFFFFF">
    <td align="center">
    <% IF(mdgbn="m") then %>
        <%= Omaechul_list.flist(i).fbaesongdate %>
    <% else %>
    	<% if right(FormatDateTime(Omaechul_list.flist(i).fbaesongdate,1),3) = "토요일" then %>
    		<font color="blue"><%= Omaechul_list.flist(i).fbaesongdate %></font>
    	<% elseif right(FormatDateTime(Omaechul_list.flist(i).fbaesongdate,1),3) = "일요일" then %>
    		<font color="red"><%= Omaechul_list.flist(i).fbaesongdate %></font>
    	<% else %>
    		<%= Omaechul_list.flist(i).fbaesongdate %>
    	<% end if %>
    <% end if %>
	</td>
	<td align="center"><%= Omaechul_list.flist(i).fsitename %></td>
	<td align="center"><%= Omaechul_list.flist(i).fsellTypeName %></td>
	<td align="center"><%= Omaechul_list.flist(i).fsellBizCdName %></td>
	<td align="center"><%= Omaechul_list.flist(i).getItemGubunName %></td>

	<td align="center"><%= Omaechul_list.flist(i).getMwGubunName %></td>

    <td align="center"><%= Replace(Omaechul_list.flist(i).ftot_itemno,"-","") %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_orgitemcost) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_orgitemcost - Omaechul_list.flist(i).ftot_itemcostCouponNotApplied) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcostCouponNotApplied) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcostCouponNotApplied - Omaechul_list.flist(i).ftot_itemcost) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcost) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_itemcost-Omaechul_list.flist(i).ftot_reducedPrice) %></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_DivSpendCouponSum) %></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).getHanDlePrice) %></td>
	<td align="center" >
	<!--
	[<a href="/admin/upchejungsan/upcheselllist.asp?datetype=chulgoil&yyyy1=<%=vYear%>&mm1=<%=vMonth%>&dd1=<%=vDay%>&yyyy2=<%=vYear%>&mm2=<%=vMonth%>&dd2=<%=vDay%>&delivertype=all" target="_blank">상세</a>]
	-->
	</td>
	<!--<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_DivSpendMileSum) %></td>-->
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftot_buycash) %></td>
	<td align="center">
	<% if (Omaechul_list.flist(i).getHanDlePrice<>0) then %>
	<%= CLNG(Omaechul_list.flist(i).ftot_buycash/Omaechul_list.flist(i).getHanDlePrice*100*100)/100 %>
	<% else %>
	-
	<% end if %>
	</td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).getCalcuMeachul) %></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).getCalcuMeachulNoVat) %></td>
	<td align="right"></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).fjPrice) %></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).fjPriceEtc) %></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).fjPriceEtcChulgo) %></td>
	<% if (jErrShow) then %>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).getErrJungsan) %></td>
	<% end if %>
</tr>
<%
	vSum_TotItemNo 						= vSum_TotItemNo + Omaechul_list.flist(i).ftot_itemno
	vSum_TotReducedPrice 				= vSum_TotReducedPrice + Omaechul_list.flist(i).ftot_reducedPrice
	vSum_TotBuycash 					= vSum_TotBuycash + Omaechul_list.flist(i).ftot_buycash
	vSum_TotBuycashCouponNotApplied 	= vSum_TotBuycashCouponNotApplied + Omaechul_list.flist(i).ftot_buycashCouponNotApplied
	vSum_TotOrgitemcost 				= vSum_TotOrgitemcost + Omaechul_list.flist(i).ftot_orgitemcost
	vSum_TotItemcostCouponNotApplied 	= vSum_TotItemcostCouponNotApplied + Omaechul_list.flist(i).ftot_itemcostCouponNotApplied
	vSum_TotItemcost 					= vSum_TotItemcost + Omaechul_list.flist(i).ftot_itemcost
	vSum_TotReducePrice					= vSum_TotReducePrice + Omaechul_list.flist(i).ftot_reducedPrice
	vSum_SpendCouponSum					= vSum_SpendCouponSum + Omaechul_list.flist(i).ftot_DivSpendCouponSum
	vSum_MaechulItem					= vSum_MaechulItem + (Omaechul_list.flist(i).ftot_itemcost-(Omaechul_list.flist(i).ftot_itemcost-Omaechul_list.flist(i).ftot_reducedPrice)-Omaechul_list.flist(i).ftot_DivSpendCouponSum)

	vSum_SpendMileSum					= vSum_SpendMileSum + Omaechul_list.flist(i).ftot_DivSpendMileSum

	vSum_jPrice                         = vSum_jPrice + Omaechul_list.flist(i).fjPrice
	vSum_jPriceEtc                      = vSum_jPriceEtc + Omaechul_list.flist(i).fjPriceEtc
	vSum_jPriceEtcChulgo                = vSum_jPriceEtcChulgo + Omaechul_list.flist(i).fjPriceEtcChulgo

	vSum_HanDlePrice                    = vSum_HanDlePrice + Omaechul_list.flist(i).getHanDlePrice
	vSum_CalcuMeachul                   = vSum_CalcuMeachul + Omaechul_list.flist(i).getCalcuMeachul
	vSum_CalcuMeachulNoVat              = vSum_CalcuMeachulNoVat + Omaechul_list.flist(i).getCalcuMeachulNoVat
	vSum_ErrJungsan                     = vSum_ErrJungsan + Omaechul_list.flist(i).getErrJungsan
Next
%>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td align="center" rowspan="2">
	총계
	</td>
	<td rowspan="2" align="center"  ></td>
	<td rowspan="2" align="center"  ></td>
	<td rowspan="2" align="center"  ></td>
	<td rowspan="2" align="center"  ></td>
	<td rowspan="2" align="center"  ></td>
	<td rowspan="2" align="center"  ><%= Replace(vSum_TotItemNo,"-","") %></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_TotOrgitemcost) %></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_TotOrgitemcost - vSum_TotItemcostCouponNotApplied) %></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_TotItemcostCouponNotApplied) %></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_TotItemcostCouponNotApplied - vSum_TotItemcost) %></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_TotItemcost) %></td>

	<td align="right"><%= NullOrCurrFormat(vSum_TotItemcost-vSum_TotReducePrice) %></td>
	<td align="right"><%= NullOrCurrFormat(vSum_SpendCouponSum) %></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_HanDlePrice) %></td>
	<td rowspan="2" ></td>
	<!--<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_SpendMileSum) %></td>-->
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_TotBuycash) %></td>
	<td rowspan="2" align="center">
	<% if (vSum_HanDlePrice<>0) then %>
	<%= CLNG(vSum_TotBuycash/vSum_HanDlePrice*100*100)/100 %>
	<% else %>
	-
	<% end if %>
	</td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_CalcuMeachul) %></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_CalcuMeachulNoVat) %></td>
	<td rowspan="2" align="right"></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_jPrice) %></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_jPriceEtc) %></td>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_jPriceEtcChulgo) %></td>
	<% if (jErrShow) then %>
	<td rowspan="2" align="right"><%= NullOrCurrFormat(vSum_ErrJungsan) %></td>
	<% end if %>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
    <td colspan="2"><%= NullOrCurrFormat(vSum_TotItemcost-vSum_TotReducePrice+vSum_SpendCouponSum)  %></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" rowspan="2">
	점유율
	</td>
	<td align="center" rowspan="2"></td>
	<td align="center" rowspan="2"></td>
	<td align="center" rowspan="2"></td>
	<td align="center" rowspan="2"></td>
	<td align="center" rowspan="2"></td>
	<td align="center" rowspan="2"></td>
	<td align="right" rowspan="2">소비가대비=&gt</td>
	<td align="center">
	<% if vSum_TotOrgitemcost<>0 then %>
	    <%= CLNG((vSum_TotOrgitemcost-vSum_TotItemcostCouponNotApplied)/vSum_TotOrgitemcost*100*100)/100 %> %
	<% end if %>
	</td>

	<td align="right" rowspan="2">판매가대비=&gt</td>
	<td align="center">
	<% if vSum_TotItemcostCouponNotApplied<>0 then %>
	    <%= CLNG((vSum_TotItemcostCouponNotApplied-vSum_TotItemcost)/vSum_TotItemcostCouponNotApplied*100*100)/100 %> %
	<% end if %>
	</td>

	<td align="right" rowspan="2"></td>

	<td align="right" colspan="2" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<!--<td align="right" rowspan="2"></td>-->
	<td align="right" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<td align="right" rowspan="2"></td>
	<% if (jErrShow) then %>
	<td align="right" rowspan="2"></td>
	<% end if %>
</tr>
<tr align="center" bgcolor="#FFFFFF">
    <td >
    <% if (vSum_TotOrgitemcost)<>0 then %>
        <%= CLNG(((vSum_TotOrgitemcost)-(vSum_TotItemcostCouponNotApplied))/(vSum_TotOrgitemcost)*100*100)/100 %> %
    <% end if %>
    </td>
    <td >
    <% if (vSum_TotItemcostCouponNotApplied)<>0 then %>
        <%= CLNG(((vSum_TotItemcostCouponNotApplied)-(vSum_TotItemcost))/(vSum_TotItemcostCouponNotApplied)*100*100)/100 %> %
    <% end if %>
    </td>
</tr>
</table>

<% set Omaechul_list = nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
