<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  핑거스 매출집계- 관심등록전환매출
' History : 2016.10.06 한용민 생성
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
dim iCurrPage,iPageSize,iTotalPage,iTotCnt, dispCate,vBrandID ,itemid
dim  vTotwww_ItemNO,vTotwww_ItemCost,vTotwww_MaechulProfit,vTotwww_BuyCash,vTotma_ItemNO,vTotma_ItemCost,vTotma_MaechulProfit
dim vTotma_BuyCash,vTotout_ItemNO,vTotout_ItemCost,vTotout_MaechulProfit	,vTotout_BuyCash			
dim vTotwww_MaechulProfitPer ,vTotma_MaechulProfitPer ,vTotout_MaechulProfitPer 
Dim vTot_OrderCnt, vTot_ItemNO, vTot_ItemcostCouponNotApplied, vTot_ItemCost, vTot_BuyCash, vTot_MaechulProfit
Dim vTot_MaechulProfitPer, vTot_BonusCouponPrice, vTot_ReducedPrice, vTot_MaechulProfit2, vTot_MaechulProfitPer2
dim vTot_upcheJungsan, lec_cdl, lec_cdm, chkImg
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
	vSorting	= NullFillWith(RequestCheckvar(request("sorting"),32),"itemsellcntD")
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
	lec_cdl = RequestCheckvar(request("lec_cdl"),3)
	lec_cdm = RequestCheckvar(request("lec_cdm"),3)
	chkImg		= requestCheckvar(request("chkImg"),1)
  	if itemid <> "" then
		if checkNotValidHTML(itemid) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
if chkImg ="" then chkImg = 0	
if iCurrPage = "" then iCurrPage = 1
if vSiteName = "" then vSiteName = "diyitem"

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
	cStatistic.FPageSize = iPageSize
	cStatistic.FCurrPage = iCurrPage
	cStatistic.FRectIncStockAvgPrc = true '' 평균매입가 포함 쿼리여부.
	cStatistic.fStatistic_wish()

	iTotCnt = cStatistic.FResultCount
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
	
'response.end
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function downloadexcel(){
    document.frm.target = "view";
    document.frm.action = "/academy/report/maechul/statistic_wish_excel.asp";
	document.frm.submit();
    document.frm.target = "";
    document.frm.action = "";
}

function searchSubmit(){
    document.frm.target = "_self"; 
    document.frm.action = "/academy/report/maechul/statistic_wish.asp";  
	
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
		alert('구분자가 없습니다.');
	}
	
	var popwin = window.open('<%= wwwFingers %>'+linkurl,'addreg','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function frontbrandlink(sitename, makerid){
	var linkurl;

	if (sitename=='diyitem'){
		linkurl = '/corner/corner_good_detail.asp?lecturer_id=' + makerid
	} else if (sitename=='academy'){
		linkurl = '/corner/corner_good_detail.asp?lecturer_id=' + makerid
	}else{
		alert('구분자가 없습니다.');
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

function fnWishUserList(sitename,itemid){
	var popwin = window.open('statistic_wish_userlist.asp?sitename='+sitename+'&itemid='+itemid,'addreg','width=200,height=300,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>

<!-- 검색 시작 -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>"> 
<input type="hidden" name="iC" value="">
<input type="hidden" name="sorting" value="<%= vsorting %>">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"  >
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>" rowspan=4>검색 조건</td>
	<td align="left">
		* 기간:
		<select name="date_gijun" class="select">
			<option value="regdate" <%=CHKIIF(vDateGijun="regdate","selected","")%>>주문일</option>
			<option value="ipkumdate" <%=CHKIIF(vDateGijun="ipkumdate","selected","")%>>결제일</option>
			<option value="beasongdate" <%=CHKIIF(vDateGijun="beasongdate","selected","")%>>출고일</option>
		</select>
		&nbsp;
		<% DrawDateBoxdynamic vSYear,"syear",vEYear,"eyear",vSMonth,"smonth",vEMonth,"emonth",vSDay,"sday",vEDay,"eday" %>
		* 채널구분 : <% drawSelectBox_SellChannel "sellchnl", sellchnl, "" %>
		&nbsp;
		* 주문구분 :
		<select name="isBanpum" class="select">
			<option value="all" <%=CHKIIF(vIsBanPum="all","selected","")%>>반품포함</option>
			<option value="<>" <%=CHKIIF(vIsBanPum="<>","selected","")%>>반품제외</option>
			<option value="=" <%=CHKIIF(vIsBanPum="=","selected","")%>>반품건만</option>
		</select>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan=4><input type="button" id="btnSubmit" class="button_s" value="검색" onClick="searchSubmit();"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<table class="a" cellpadding="0" border=0 cellspacing="0" width="100%">
		<tr>
			<td colspan="2" width=400>
				* 브랜드 : <input type="text" class="text" name="ebrand" value="<%=vBrandID%>" size="20"> <input type="button" class="button" value="IDSearch" onclick="jsSearchBrandID(this.form.name,'ebrand');">
			    &nbsp;
			    * 상품코드 :
			</td>
			<td rowspan="2" align="left"  >
				<textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
			</td>
		</tr> 
		<tr>
			<td width=400>
				<input type="checkbox" name="chkImg" value="1" <%if chkImg = 1 then%>checked<%end if%>>상품이미지 보기
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<% if vSiteName="academy" then %>
			* 매입구분 : <% Call DrawBrandMWUPCombo("mwdiv",mwdiv) %>
		<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 사이트구분(필수값) : <% drawradio_academy_sitename "sitename", vSiteName, "", "N" %>
		<% if vSiteName="academy" then %>
			* 카테고리 :  <% DrawSelectBoxLecCategoryLarge "lec_cdl" ,  lec_cdl  , "N" %>

			<% if lec_cdl <> "" Then %>
				* 중카테고리 : <% call DrawSelectBoxLecCategoryMid("lec_cdm", lec_cdl, lec_cdm, "N") %>
			<% end if %>
		<% elseif vSiteName="diyitem" then %>
			* 기능<!-- #include virtual="/academy/comm/CategorySelectBox.asp"-->
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
			       		// 내용 넣기 
			       		$("#lyrDispCtBox").empty().html(message);
			       		$("#oDispCate").val(dc);
					}
				});
			}
			</script>
			* 전시카테고리 : <span id="lyrDispCtBox"></span>
			<input type="hidden" name="disp" id="oDispCate" value="<%=dispCate%>">
		<% end if %>
	</td>
</tr>
</table>
<!-- 검색 끝 -->
<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left"></td>
	<td align="right">
		<input type="button" onclick="downloadexcel();" value="엑셀다운로드" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

</form>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		검색결과 : <b><%=iTotCnt%></b>
		&nbsp;
		페이지 : <b><%= iCurrPage %> / <%=iTotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td></td>

	<% IF chkImg = 1 then %>
		<td></td>
	<% END IF %>

	<td></td>
    <td></td>
    <td></td>
	<td>A</td>
	<td>B</td>
    <!--<td>C</td>-->
    <td>D</td>
    <td>E</td>
    <td>F</td>
    <td>G</td>
    <td>H</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>상품코드</td>

	<% IF chkImg = 1 then %>
		<td>이미지</td>
	<% END IF %>

	<td>브랜드</td>
    <td>전시카테고리</td>
    <td>상품명</td>
	<td onClick="jstrSort('sellcash'); return false;" style="cursor:hand;">
		판매가
		<img src="/images/list_lineup<%=CHKIIF(vSorting="sellcashD","_bot","_top")%><%=CHKIIF(instr(vSorting,"sellcash")>0,"_on","")%>.png" id="imgsellcash">
	</td>
	<td onClick="jstrSort('buycash'); return false;" style="cursor:hand;">
		매입가
		<img src="/images/list_lineup<%=CHKIIF(vSorting="buycashD","_bot","_top")%><%=CHKIIF(instr(vSorting,"buycash")>0,"_on","")%>.png" id="imgbuycash">
	</td>
    <!--<td onClick="jstrSort('totwishcnt'); return false;" style="cursor:hand;">
    	총담은수
    	<img src="/images/list_lineup<% '=CHKIIF(vSorting="totwishcntD","_bot","_top")%><% '=CHKIIF(instr(vSorting,"totwishcnt")>0,"_on","")%>.png" id="imgtotwishcnt">
    	<Br>D+E
    </td>-->
    <td onClick="jstrSort('itemsellcnt'); return false;" style="cursor:hand;">
    	판매전환수
    	<br>(판매건수)
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="itemsellcntD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemsellcnt")>0,"_on","")%>.png" id="imgitemsellcnt">
    </td>
    <td onClick="jstrSort('itemwishcnt'); return false;" >
    	위시
    	<br>담긴건수
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="itemwishcntD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemwishcnt")>0,"_on","")%>.png" id="imgitemwishcnt">
    </td>
    <td onClick="jstrSort('itemsellconversrate'); return false;" >
    	판매전환율
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="itemsellconversrateD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemsellconversrate")>0,"_on","")%>.png" id="imgitemsellconversrate">
    </td>
    <td onClick="jstrSort('itemsellsum'); return false;" style="cursor:hand;">
    	전체매출
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="itemsellsumD","_bot","_top")%><%=CHKIIF(instr(vSorting,"itemsellsum")>0,"_on","")%>.png" id="imgitemsellsum">
    </td>
    <td onClick="jstrSort('recentfavcount'); return false;" style="cursor:hand;">
    	최근위시수1일
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="recentfavcountD","_bot","_top")%><%=CHKIIF(instr(vSorting,"recentfavcount")>0,"_on","")%>.png" id="imgrecentfavcount">
    </td>
</tr>
<% if cStatistic.FTotalCount>0 then %>
	<%
	dim tot_totwishcnt, tot_itemsellcnt, tot_itemwishcnt, tot_itemsellconversrate, tot_itemsellsum, tot_recentfavcount

	For i = 0 To cStatistic.FTotalCount -1

	'tot_totwishcnt = tot_totwishcnt + cStatistic.FItemList(i).ftotwishcnt
	tot_itemsellcnt = tot_itemsellcnt + cStatistic.FItemList(i).fitemsellcnt
	tot_itemwishcnt = tot_itemwishcnt + cStatistic.FItemList(i).fitemwishcnt
	tot_itemsellconversrate = tot_itemsellconversrate + cStatistic.FItemList(i).fitemsellconversrate
	tot_itemsellsum = tot_itemsellsum + cStatistic.FItemList(i).fitemsellsum
	tot_recentfavcount = tot_recentfavcount + cStatistic.FItemList(i).frecentfavcount
	%>
	<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFFF';>
		<td>
			<a href="#" onclick="frontitemlink('<%= vSiteName %>','<%= cStatistic.FItemList(i).FitemID %>'); return false;">
			<%= cStatistic.FItemList(i).FitemID %></a>
		</td>

		<% IF chkImg = 1 then %>
			<td><img src="<%= cStatistic.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
		<% END IF %>

		<td>
			<a href="#" onclick="frontbrandlink('<%= vSiteName %>','<%= cStatistic.FItemList(i).fmakerid %>'); return false;">
			<%= cStatistic.FItemList(i).FMakerID %></a>
		</td>
		<td align="left">
			<% if cStatistic.FItemList(i).Fcode_large_nm<>"" then %>
				<%= Replace(cStatistic.FItemList(i).Fcode_large_nm,"^^"," >> ") %>
			<% end if %>

			<% if cStatistic.FItemList(i).Fcode_mid_nm <> "" then %>
				>> <%= cStatistic.FItemList(i).Fcode_mid_nm %>
			<% end if %>
		</td>
		<td align="left"><%= cStatistic.FItemList(i).fitemname %></td>
		<td align="right"><%= CurrFormat(cStatistic.FItemList(i).fsellcash) %></td>
		<td align="right"><%= CurrFormat(cStatistic.FItemList(i).fbuycash) %></td>
		<!--<td align="right"><%'= CurrFormat(cStatistic.FItemList(i).ftotwishcnt) %></td>-->
		<td align="right"><%= CurrFormat(cStatistic.FItemList(i).fitemsellcnt) %></td>
		<td align="right" onclick="fnWishUserList('<%=vSiteName%>',<%= cStatistic.FItemList(i).FitemID %>);"><%= CurrFormat(cStatistic.FItemList(i).fitemwishcnt) %></td>
		<td align="right"><%= round(CurrFormat(cStatistic.FItemList(i).fitemsellconversrate),1) %>%</td>
		<td align="right"><%= CurrFormat(cStatistic.FItemList(i).fitemsellsum) %></td>
		<td align="right"><%= CurrFormat(cStatistic.FItemList(i).frecentfavcount) %></td>
	</tr>
	<% Next %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="<% IF chkImg = 1 then %>7<% else %>6<% end if %>">총계</td>
		<!--<td align="right"><%= CurrFormat(tot_totwishcnt) %></td>-->
		<td align="right"><%= CurrFormat(tot_itemsellcnt) %></td>
		<td align="right"><%= CurrFormat(tot_itemwishcnt) %></td>
		<td align="right"><%= round(CurrFormat(tot_itemsellconversrate/cStatistic.FTotalCount),1) %>%</td>
		<td align="right"><%= CurrFormat(tot_itemsellsum) %></td>
		<td align="right"><%= CurrFormat(tot_recentfavcount) %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan="25">
			<%'sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center">검색결과가 없습니다.</td>
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