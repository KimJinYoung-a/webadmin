<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim yyyy1,mm1, yyyy2, mm2, chkdate, chknotfinish
dim research, page
yyyy1       = request("yyyy1")
mm1         = request("mm1")
yyyy2       = request("yyyy2")
mm2         = request("mm2")
chkdate     = request("chkdate")
chknotfinish= request("chknotfinish")
research    = request("research")
page        = request("page")

if (research="") and (chkdate="") then chkdate="on"
if (page="") then page=1

dim stdt, eddt, StartYYYYMM, EndYYYYMM
if (yyyy1="") then
	stdt = dateserial(year(Now),month(now)-6,1)
	yyyy1 = Left(CStr(stdt),4)
	mm1 = Mid(CStr(stdt),6,2)

	eddt = dateadd("d",dateserial(year(Now),month(now)+1,1),-1)
	yyyy2 = Left(CStr(eddt),4)
	mm2 = Mid(CStr(eddt),6,2)
end if


StartYYYYMM = yyyy1 + "-" + mm1
EndYYYYMM   = yyyy2 + "-" + mm2


dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FRectFixStateExiste = chknotfinish
if (chkdate="on") then
    ojungsan.FRectStartYYYYMM = StartYYYYMM
    ojungsan.FRectEndYYYYMM   = EndYYYYMM
end if
ojungsan.JungsanSummary0



dim i,p_yyyymm,subtotalFlag
dim sumub, summe, sumwi, sumet, sumdlv,sumbuytot, sum_notconfirmsum, sum_confirmsum, sum_ipkumsum, sum_fixedthissum, sum_fixednextsum
dim allsumub, allsumme, allsumwi, allsumet, allsumdlv, allsumbuytot, allsum_notconfirmsum, allsum_confirmsum, allsum_ipkumsum, allsum_fixedthissum, allsum_fixednextsum

dim sumub_sell, summe_sell, sumwi_sell, sumet_sell, sumdlv_sell, sumselltot
dim allsumub_sell, allsumme_sell, allsumwi_sell, allsumet_sell, allsumdlv_sell, allsumselltot

dim ipsum
%>
<script language='javascript'>
function reSearch(){
	var arr = '';
	if (frm2.ck_dummi1.checked) frm.ck_dummi1.value="on";
	if (frm2.ck_dummi2.checked) frm.ck_dummi2.value="on";
	frm.submit();
}

function CheckEnabled(frm,comp){
    if (comp.name=='chkdate'){
        frm.chknotfinish.checked=false;
    }else{
        frm.chkdate.checked=false;
    }
}

</script>


<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
   	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	<input type="checkbox" name="chkdate" <% if chkdate="on" then response.write "checked" %> onClick="CheckEnabled(frm,this);">&nbsp;기간검색 : <% DrawYMYMBox yyyy1,mm1, yyyy2,mm2 %> (정산 대상월)
	        	&nbsp;&nbsp;
	        	<input type="checkbox" name="chknotfinish" <% if chknotfinish="on" then response.write "checked" %> onClick="CheckEnabled(frm,this);">&nbsp;처리완료 정산월 제외
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->



<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#DDDDFF">
		<td rowspan="2" width="90">대상월</td>
		<td rowspan="2" width="60">정산일</td>
		<td colspan="5">계약별 구분</td>
		<td rowspan="2" width="100">정산총액</td>
		<td colspan="4">입금진행현황</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="90">업체배송</td>
		<td width="90">매입</td>
		<td width="90">위탁</td>
		<td width="90">기타</td>
		<td width="90">배송비</td>
		<td width="90">확정이전금액</td>
		<td width="90">확정금액<br>(금월결제)</td>
		<td width="90">확정금액<br>(익월결제)</td>
		<td width="90">입금완료금액</td>
	</tr>
	<% for i=0 to ojungsan.FresultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
	  <td><%= ojungsan.FItemList(i).Fyyyymm %></td>
	  <td><%= ojungsan.FItemList(i).Fjungsan_date %></td>
	  <td align="right"><%= FormatNumber(ojungsan.FItemList(i).Fuptot,0) %></td>
	  <td align="right"><%= FormatNumber(ojungsan.FItemList(i).Fmetot,0) %></td>
	  <td align="right"><%= FormatNumber(ojungsan.FItemList(i).Fwitot,0) %></td>
	  <td align="right"><%= FormatNumber(ojungsan.FItemList(i).Fettot,0) %></td>
	  <td align="right"><%= FormatNumber(ojungsan.FItemList(i).Fdlvtot,0) %></td>
	  <td align="right">
	    <%= FormatNumber(ojungsan.FItemList(i).getTotSum,0) %>
	    <% if ojungsan.FItemList(i).getTotSum<>(ojungsan.FItemList(i).Ftotflag_notconfirmsum + ojungsan.FItemList(i).Ffixedthissum + ojungsan.FItemList(i).Ffixednextsum + ojungsan.FItemList(i).Ftotflag_ipkumsum) then %>
	    <font color="blue"><%= FormatNumber((ojungsan.FItemList(i).Ftotflag_notconfirmsum + ojungsan.FItemList(i).Ffixedthissum + ojungsan.FItemList(i).Ffixednextsum + ojungsan.FItemList(i).Ftotflag_ipkumsum),0) %></font>
	    <% end if %>
	  </td>
	  <td align="right"><%= FormatNumber(ojungsan.FItemList(i).Ftotflag_notconfirmsum,0) %></td>
	  <td align="right"><%= FormatNumber(ojungsan.FItemList(i).Ffixedthissum,0) %></td>
	  <td align="right"><%= FormatNumber(ojungsan.FItemList(i).Ffixednextsum,0) %></td>
	  <td align="right"><%= FormatNumber(ojungsan.FItemList(i).Ftotflag_ipkumsum,0) %></td>
	</tr>
	<%
	'' sub total
	p_yyyymm = ojungsan.FItemList(i).FYYYYMM


	sumub           = sumub + ojungsan.FItemList(i).Fuptot
	summe           = summe + ojungsan.FItemList(i).Fmetot
	sumwi           = sumwi + ojungsan.FItemList(i).Fwitot
	sumet           = sumet + ojungsan.FItemList(i).Fettot
	sumdlv          = sumdlv + ojungsan.FItemList(i).Fdlvtot

	sumbuytot       = sumbuytot + ojungsan.FItemList(i).getTotSum
	sum_notconfirmsum   = sum_notconfirmsum + ojungsan.FItemList(i).Ftotflag_notconfirmsum
	sum_confirmsum      = sum_confirmsum + ojungsan.FItemList(i).Ftotflag_confirmsum

	sum_fixedthissum    = sum_fixedthissum + ojungsan.FItemList(i).Ffixedthissum
	sum_fixednextsum    = sum_fixednextsum + ojungsan.FItemList(i).Ffixednextsum

	sum_ipkumsum        = sum_ipkumsum + ojungsan.FItemList(i).Ftotflag_ipkumsum


	sumub_sell = sumub_sell + ojungsan.FItemList(i).Fupselltot
	summe_sell = summe_sell + ojungsan.FItemList(i).Fmeselltot
	sumwi_sell = sumwi_sell + ojungsan.FItemList(i).Fwiselltot
	sumet_sell = sumet_sell + ojungsan.FItemList(i).Fetselltot
	sumdlv_sell = sumdlv_sell + ojungsan.FItemList(i).Fdlvselltot
	sumselltot = sumselltot + ojungsan.FItemList(i).getTotSellcashSum


	subtotalFlag=false

	if (i=ojungsan.FResultCount-1) then
	    subtotalFlag=true
	elseif (ojungsan.FItemList(i+1).FYYYYMM<>p_yyyymm) then
	    subtotalFlag=true
	else
	    subtotalFlag=false
	end if

	%>




	<% if (subtotalFlag) then %>
	<tr align="center" bgcolor="#F3F3FF">
	  <td ><b><%= ojungsan.FItemList(i).Fyyyymm %><b></td>
	  <td>합계</td>
	  <td align="right"><%= FormatNumber(sumub,0) %></td>
	  <td align="right"><%= FormatNumber(summe,0) %></td>
	  <td align="right"><%= FormatNumber(sumwi,0) %></td>
	  <td align="right"><%= FormatNumber(sumet,0) %></td>
	  <td align="right"><%= FormatNumber(sumdlv,0) %></td>
	  <td align="right"><b><%= FormatNumber(sumbuytot,0) %></b></td>
	  <td align="right"><b><%= FormatNumber(sum_notconfirmsum,0) %></b></td>
	  <td align="right"><b><%= FormatNumber(sum_fixedthissum,0) %></b></td>
	  <td align="right"><b><%= FormatNumber(sum_fixednextsum,0) %></b></td>
	  <td align="right"><b><%= FormatNumber(sum_ipkumsum,0) %></b></td>
	</tr>
	<tr align="center" bgcolor="#F3F3FF">
	  <td ></td>
	  <td ></td>
	  <td align="right"><%= FormatNumber(sumub_sell,0) %></td>
	  <td align="right"><%= FormatNumber(summe_sell,0) %></td>
	  <td align="right"><%= FormatNumber(sumwi_sell,0) %></td>
	  <td align="right"><%= FormatNumber(sumet_sell,0) %></td>
	  <td align="right"><%= FormatNumber(sumdlv_sell,0) %></td>
	  <td align="right"><%= FormatNumber(sumselltot,0) %>.</td>
	  <td></td>
	  <td></td>
	  <td></td>
	  <td></td>
	</tr>
	<tr align="center" bgcolor="#D3D3FF">
	  <td ></td>
	  <td ></td>
	  <td align="right"><% if sumub_sell<>0 then response.write CLng((sumub_sell-sumub)/sumub_sell*10000)/100 %> %</td>
	  <td align="right"><% if summe_sell<>0 then response.write CLng((summe_sell-summe)/summe_sell*10000)/100 %> %</td>
	  <td align="right"><% if sumwi_sell<>0 then response.write CLng((sumwi_sell-sumwi)/sumwi_sell*10000)/100 %> %</td>
	  <td align="right"><% if sumet_sell<>0 then response.write CLng((sumet_sell-sumet)/sumet_sell*10000)/100 %> %</td>
	  <td align="right"><% if sumdlv_sell<>0 then response.write CLng((sumdlv_sell-sumdlv)/sumdlv_sell*10000)/100 %> %</td>
	  <td align="right"><% if sumselltot<>0 then response.write CLng((sumselltot-sumbuytot)/sumselltot*10000)/100 %> %</td>
	  <td></td>
	  <td></td>
	  <td></td>
	  <td></td>
	</tr>
		<%
		allsumub = allsumub + sumub
		allsumme = allsumme + summe
		allsumwi = allsumwi + sumwi
		allsumet= allsumet + sumet
		allsumdlv= allsumdlv + sumdlv
		allsumbuytot = allsumbuytot + sumbuytot

		allsum_notconfirmsum = allsum_notconfirmsum + sum_notconfirmsum
		allsum_confirmsum = allsum_confirmsum + sum_confirmsum

		allsum_fixedthissum = allsum_fixedthissum + sum_fixedthissum
		allsum_fixednextsum = allsum_fixednextsum + sum_fixednextsum
		allsum_ipkumsum     = allsum_ipkumsum + sum_ipkumsum


		allsumub_sell = allsumub_sell + sumub_sell
		allsumme_sell = allsumme_sell + summe_sell
		allsumwi_sell = allsumwi_sell + sumwi_sell
		allsumet_sell = allsumet_sell + sumet_sell
		allsumdlv_sell = allsumdlv_sell + sumdlv_sell
		allsumselltot = allsumselltot + sumselltot


		sumub = 0
		summe = 0
		sumwi = 0
		sumet = 0
		sumdlv = 0
		sumbuytot = 0
		sum_notconfirmsum   = 0
		sum_confirmsum      = 0

		sum_fixedthissum    = 0
		sum_fixednextsum    = 0
		sum_ipkumsum        = 0

		sumub_sell = 0
		summe_sell = 0
		sumwi_sell = 0
		sumet_sell = 0
		sumdlv_sell = 0
		sumselltot = 0

		%>
	<% end if %>
<% next %>
	<tr bgcolor="#FFFFFF">
	  <td align="center"><b>Total</b></td>
	  <td ></td>
	  <td align="right"><%= FormatNumber(allsumub,0) %></td>
	  <td align="right"><%= FormatNumber(allsumme,0) %></td>
	  <td align="right"><%= FormatNumber(allsumwi,0) %></td>
	  <td align="right"><%= FormatNumber(allsumet,0) %></td>
	  <td align="right"><%= FormatNumber(allsumdlv,0) %></td>
	  <td align="right"><b><%= FormatNumber(allsumbuytot,0) %></b></td>
	  <td align="right"><b><%= FormatNumber(allsum_notconfirmsum,0) %></b></td>
	  <td align="right"><b><%= FormatNumber(allsum_fixedthissum,0) %></b></td>
	  <td align="right"><b><%= FormatNumber(allsum_fixednextsum,0) %></b></td>
	  <td align="right"><b><%= FormatNumber(allsum_ipkumsum,0) %></b></td>
	</tr>
	<tr align="center" bgcolor="#D3D3FF">
	  <td></td>
	  <td ></td>
	  <td align="right"><% if allsumub_sell<>0 then response.write CLng((allsumub_sell-allsumub)/allsumub_sell*10000)/100 %> %</td>
	  <td align="right"><% if allsumme_sell<>0 then response.write CLng((allsumme_sell-allsumme)/allsumme_sell*10000)/100 %> %</td>
	  <td align="right"><% if allsumwi_sell<>0 then response.write CLng((allsumwi_sell-allsumwi)/allsumwi_sell*10000)/100 %> %</td>
	  <td align="right"><% if allsumet_sell<>0 then response.write CLng((allsumet_sell-allsumet)/allsumet_sell*10000)/100 %> %</td>
	  <td align="right"><% if allsumdlv_sell<>0 then response.write CLng((allsumdlv_sell-allsumdlv)/allsumdlv_sell*10000)/100 %> %</td>
	  <td align="right"><% if allsumselltot<>0 then response.write CLng((allsumselltot-allsumbuytot)/allsumselltot*10000)/100 %> %</td>
	  <td></td>
	  <td></td>
	  <td></td>
	  <td></td>
	</tr>
</table>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right" bgcolor="F4F4F4">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<%
set ojungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->