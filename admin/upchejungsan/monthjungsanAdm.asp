<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/jungsan/jungsanTaxCls.asp"-->
<%
dim makerid, yyyy1,mm1
makerid = requestCheckvar(request("makerid"),32)
yyyy1   = requestCheckvar(request("yyyy1"),10)
mm1     = requestCheckvar(request("mm1"),10)

if (yyyy1="") then
    yyyy1 = LEFT(dateadd("m",-1,now()),4)
    mm1 = MID(dateadd("m",-1,now()),6,2)
end if

dim ojungsanTaxCC
set ojungsanTaxCC = new CUpcheJungsanTax
ojungsanTaxCC.FRectMakerid = makerid
ojungsanTaxCC.FRectYYYYMM = yyyy1+"-"+mm1
ojungsanTaxCC.FRectJGubun = "CC"
ojungsanTaxCC.getMonthUpcheJungsanList

dim ojungsanTaxCE
set ojungsanTaxCE = new CUpcheJungsanTax
ojungsanTaxCE.FRectMakerid = makerid
ojungsanTaxCE.FRectYYYYMM = yyyy1+"-"+mm1
ojungsanTaxCE.FRectJGubun = "CE"
ojungsanTaxCE.getMonthUpcheJungsanList

dim ojungsanTaxMM
set ojungsanTaxMM = new CUpcheJungsanTax
ojungsanTaxMM.FRectMakerid = makerid
ojungsanTaxMM.FRectYYYYMM = yyyy1+"-"+mm1
ojungsanTaxMM.FRectJGubun = "MM"
ojungsanTaxMM.getMonthUpcheJungsanList

dim i
%>
<script language='javascript'>
function PopDetail(iidx,tg,igroupid){
    var uri = 'jungsandetailsumONAdm.asp?id=' + iidx + '&groupid='+igroupid;
    if (tg=="OF") uri = 'jungsandetailsumOFAdm.asp?idx=' + iidx + '&groupid='+igroupid;
	var popwin = window.open(uri+'&makerid=<%=makerid%>','PopDetail','width=1280, height=800, scrollbars=1, resizable=yes');
	popwin.focus();
}

function PopTaxRegPrdCommission(makerid, yyyy1, mm1, onoffGubun, jidx) {
	<% 'var popwin = window.open("popTaxRegAdmin.asp?makerid=" + makerid + "&yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&onoffGubun=" + onoffGubun + "&jidx="+jidx,"PopTaxRegPrdCommission","width=640 height=700 scrollbars=yes resizable=yes"); %>
    var popwin = window.open("/admin/upchejungsan/popTaxRegAdminapi.asp?makerid=" + makerid + "&yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&onoffGubun=" + onoffGubun + "&jidx="+jidx,"PopTaxRegPrdCommission","width=1024 height=768 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopTaxPrintReDirect(itax_no){
	var popwinsub = window.open("red_taxprint.asp?tax_no=" + itax_no ,"taxview","width=800,height=700,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popwinsub.focus();
}

function PopConfirm(mnupos,iidx){
	//var popwin = window.open('jungsanmaster.asp?id=' + iidx + '&menupos=' + mnupos,'popshowdetail','width=900, height=540, scrollbars=1');
	//popwin.focus();
}

function PopTaxReg(v){
	//var popwin = window.open("poptaxreg.asp?id=" + v,"poptaxreg","width=640 height=700 scrollbars=yes resizable=yes");
	//popwin.focus();
}

function PopTaxRegOff(v){
	//var popwin = window.open("poptaxregoff.asp?idx=" + v,"poptaxregoff1","width=640 height=680 scrollbars=yes resizable=yes");
	//popwin.focus();
}
<% if (ojungsanTaxCC.FresultCount>0) then %>
//alert('2014년 1월 정산부터 수수료 정산분에 대해서는\n\n텐바이텐에서 계산서를 발행 하오니\n\n이세로 등을 통해 따로 발행하지 말아 주시길 바랍니다.');
<% end if %>
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		정산 대상 년월 :&nbsp;<% DrawYMBox yyyy1,mm1 %>
		&nbsp;
		브랜드ID : <% drawSelectBoxDesignerwithName "makerid",makerid  %>&nbsp;&nbsp;
		<!--
		<span ><strong>수정중 지급대상액 (배송비), 추가정산액 (기타) 분리 관련</strong></span>
        -->
	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
</form>
</table>
<p>

<% if (ojungsanTaxCC.FresultCount<1) and (ojungsanTaxCE.FresultCount<1) and (ojungsanTaxMM.FresultCount<1) then %>
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td align="left"><strong>* 정산 내역</strong></td>
</tr>
<tr height="30">
    <td align="center" bgcolor="#FFFFFF"> 검색 결과가 존재 하지 않습니다.</td>
</tr>
</table>
<% else %>


<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="15" align="left"><strong>* 수수료 정산 내역</strong> <font color=red>(수수료 세금 계산서는 <b>텐바이텐</b>에서 <b>일괄 발행</b>합니다.)</font></td>
</tr>
<% if (ojungsanTaxCC.FresultCount>0) then %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td width="60" >정산월</td>
    <td width="60" >매출처</td>
    <td width="50" >과세<br>구분</td>
    <td width="90" >브랜드ID</td>
    <td width="180" >정산내역</td>
    <td width="90" >협력사 매출액<br>상품</td>
    <td width="80" >수수료</td>
    <td width="80" >협력사 매출액<br>배송비</td>
    <td width="100">지급대상액<br>(상품)</td>
  	<td width="80">지급대상액<br>(배송비)</td>
  	<td width="80">추가정산액<br>(기타)</td>
  	<td width="80">지급예정액</td>
    <!--td width="60" >지급예정일</td-->
    <td width="90" >계산서상태</td>
    <td width="80" >계산서</td>
    <td >상세조회</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" >
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td>(A)</td>
    <td>(B)</td>
    <td>(C)</td>
    <td>(D)</td>
    <td>(E)</td>
    <td>(F)</td>
    <td>(G)</td>
    <td>(H)</td>
    <td>(I)</td>
    <td>(J)</td>
</tr>
<% for i=0 to ojungsanTaxCC.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td><%=ojungsanTaxCC.FItemList(i).Fyyyymm%></td>
    <td><%=ojungsanTaxCC.FItemList(i).getTargetNm%></td>
    <td><%=ojungsanTaxCC.FItemList(i).getItemVatTypeName%></td>
    <td align="left"><%=ojungsanTaxCC.FItemList(i).Fmakerid%></td>
    <td align="left"><%=ojungsanTaxCC.FItemList(i).Ftitle%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxCC.FItemList(i).FPrdMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxCC.FItemList(i).FPrdCommissionSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxCC.FItemList(i).FdlvMeachulsum + ojungsanTaxCC.FItemList(i).FetMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxCC.FItemList(i).FprdJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxCC.FItemList(i).FdlvJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxCC.FItemList(i).FetJungsanSum,0) %></td>
    <td align="right"><%=FormatNumber(ojungsanTaxCC.FItemList(i).getToTalJungsanSum,0)%></td>
    <!--td><%= ojungsanTaxCC.FItemList(i).getMayIpkumdateStr %></td -->
    <td><%=ojungsanTaxCC.FItemList(i).GetTaxEvalStateName%></td>
    <td>
        <% if ojungsanTaxCC.FItemList(i).IsElecTaxExists then %>
      	<a href="javascript:PopTaxPrintReDirect('<%= ojungsanTaxCC.FItemList(i).Fneotaxno %>');">출력
      	<img src="/images/icon_print02.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
		<% else %>
      	<!--<a href="javascript:PopTaxRegPrdCommission('<%'=ojungsanTaxCC.FItemList(i).Fmakerid %>', '<%'= yyyy1 %>', '<%'= mm1 %>', '<%'= ojungsanTaxCC.FItemList(i).FtargetGbn %>','<%'= ojungsanTaxCC.FItemList(i).Fid %>');">발행-->
      	<!--<img src="/images/icon_new.gif" width="14" height="14" border="0" align="absbottom">-->
      	<!--</a>-->
      	<% end if %>
    </td>
    <td>
    <a href="javascript:PopDetail('<%= ojungsanTaxCC.FItemList(i).FId %>','<%= ojungsanTaxCC.FItemList(i).FtargetGbn%>','<%= ojungsanTaxCC.FItemList(i).Fgroupid%>');">보기<img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom"></a>
    </td>
</tr>
<% next %>
<% else %>
<tr align="center" bgcolor="#FFFFFF" height="30">
<td align="center" colspan="14">검색결과가 없습니다.</td>
</tr>
<% end if %>
</table>
<p><br>


<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="14" align="left"><strong>* 기타 정산 내역</strong> <font color=red>(기타 정산 세금 계산서는 <b>텐바이텐</b>에서 <b>일괄 발행</b>합니다.)</font></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td width="60" >정산월</td>
    <td width="60" >매출처</td>
    <td width="50" >과세<br>구분</td>
    <td width="90" >브랜드ID</td>
    <td width="180" >정산내역</td>
    <td width="90" ></td>
    <td width="80" >프로모션<br>(협력사 부담)</td>
    <td width="80" ></td>
    <td width="100"></td>
  	<td width="80"></td>
  	<td width="80">지급예정액</td>
    <td width="90" >계산서상태</td>
    <td width="80" >계산서</td>
    <td >상세조회</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" >
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td>(B)</td>
    <td></td>
    <td></td>
    <td></td>
    <td>(F)</td>
    <td>(G)</td>
    <td>(H)</td>
    <td>(I)</td>
</tr>
<% if (ojungsanTaxCE.FresultCount>0) then %>
<% for i=0 to ojungsanTaxCE.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td><%=ojungsanTaxCE.FItemList(i).Fyyyymm%></td>
    <td><%=ojungsanTaxCE.FItemList(i).getTargetNm%></td>
    <td><%=ojungsanTaxCE.FItemList(i).getItemVatTypeName%></td>
    <td align="left"><%=ojungsanTaxCE.FItemList(i).Fmakerid%></td>
    <td align="left"><%=ojungsanTaxCE.FItemList(i).Ftitle%></td>
    <td align="right"></td>
    <td align="right"><%=FormatNumber(ojungsanTaxCE.FItemList(i).FPrdCommissionSum,0)%></td>
    <td align="right"></td>
    <td align="right"></td>
    <td align="right"></td>
    <td align="right"><%=FormatNumber(ojungsanTaxCE.FItemList(i).getToTalJungsanSum,0)%></td>
    <td><%=ojungsanTaxCE.FItemList(i).GetTaxEvalStateName%></td>
    <td>
        <% if ojungsanTaxCE.FItemList(i).IsElecTaxExists then %>
      	<a href="javascript:PopTaxPrintReDirect('<%= ojungsanTaxCE.FItemList(i).Fneotaxno %>');">출력
      	<img src="/images/icon_print02.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
		<% else %>
      	<!--<a href="javascript:PopTaxRegPrdCommission('<%'=ojungsanTaxCE.FItemList(i).Fmakerid %>', '<%'= yyyy1 %>', '<%'= mm1 %>', '<%'= ojungsanTaxCE.FItemList(i).FtargetGbn %>','<%'= ojungsanTaxCE.FItemList(i).Fid %>');">발행-->
      	<!--<img src="/images/icon_new.gif" width="14" height="14" border="0" align="absbottom">-->
      	<!--</a>-->
      	<% end if %>
    </td>
    <td>
    <a href="javascript:PopDetail('<%= ojungsanTaxCE.FItemList(i).FId %>','<%= ojungsanTaxCE.FItemList(i).FtargetGbn%>','<%= ojungsanTaxCE.FItemList(i).Fgroupid%>');">보기<img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom"></a>
    </td>
</tr>
<% next %>
<% else %>
<tr align="center" bgcolor="#FFFFFF" height="30">
<td align="center" colspan="14">검색결과가 없습니다.</td>
</tr>
<% end if %>
</table>
<p><br>



<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="13" align="left"><strong>* 매입 정산 내역</strong> (협력사에서 텐바이텐으로 발행해 주셔야 합니다.) (롯데닷컴 판매 내역 및 가맹점 판매 내역은 매입정산으로 처리 됩니다.)</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td width="60" >정산월</td>
    <td width="60" >매출처</td>
    <td width="50" >과세<br>구분</td>
    <td width="90" >브랜드ID</td>
    <td width="170" >정산내역</td>
    <td width="90" >협력사<br>상품공급액</td>
    <td width="80" >배송비/기타</td>
    <td width="100">지급대상액<br>(상품)</td>
  	<td width="80">지급대상액<br>(배송비/기타)</td>
  	<td width="80">협력사매출액<br>(지급예정액)</td>
    <!--td width="60" >지급예정일</td-->
    <td width="90" >계산서상태</td>
    <td width="80" >계산서</td>
    <td >상세조회</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" >
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td>(a)</td>
    <td>(b)</td>
    <td>(c)</td>
    <td>(d)</td>
    <td>(e)</td>
    <td>(f)</td>
    <td>(g)</td>
    <td>(h)</td>
</tr>
<% if (ojungsanTaxMM.FresultCount>0) then %>
<% for i=0 to ojungsanTaxMM.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td><%=ojungsanTaxMM.FItemList(i).Fyyyymm%></td>
    <td><%=ojungsanTaxMM.FItemList(i).getTargetNm%></td>
    <td><%=ojungsanTaxMM.FItemList(i).getTaxtypeName%></td>
    <td align="left"><%=ojungsanTaxMM.FItemList(i).Fmakerid%></td>
    <td align="left"><%=ojungsanTaxMM.FItemList(i).Ftitle%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxMM.FItemList(i).FPrdMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxMM.FItemList(i).FdlvMeachulsum + ojungsanTaxMM.FItemList(i).FetMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxMM.FItemList(i).FprdJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxMM.FItemList(i).FdlvJungsanSum + ojungsanTaxMM.FItemList(i).FetJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxMM.FItemList(i).getToTalJungsanSum,0)%></td>
    <!--td><%= ojungsanTaxMM.FItemList(i).getMayIpkumdateStr %></td-->
    <td><%=ojungsanTaxMM.FItemList(i).GetTaxEvalStateName%></td>
    <td>
        <% if ojungsanTaxMM.FItemList(i).IsElecTaxExists then %>
      	<a href="javascript:PopTaxPrintReDirect('<%= ojungsanTaxMM.FItemList(i).Fneotaxno %>');">출력
      	<img src="/images/icon_print02.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	<% elseif ojungsanTaxMM.FItemList(i).IsCommissionTax then %>
      	</a>
      	<% elseif ojungsanTaxMM.FItemList(i).IsElecTaxCase then %>
      	<!--
      	<a href="javascript:PopTaxReg<%=CHKIIF(ojungsanTaxMM.FItemList(i).FtargetGbn="OF","Off","")%>('<%= ojungsanTaxMM.FItemList(i).FId %>');">발행
      	<img src="/images/icon_new.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	-->
      	<% elseif ojungsanTaxMM.FItemList(i).IsElecFreeTaxCase then %>
      	<!--
      	<a href="javascript:PopTaxReg('<%= ojungsanTaxMM.FItemList(i).FId %>');">발행
      	<img src="/images/icon_new.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	-->
      	<% elseif ojungsanTaxMM.FItemList(i).IsElecSimpleBillCase then %>
      	<!--
      	<a href="javascript:PopConfirm('<%= menupos %>','<%= ojungsanTaxMM.FItemList(i).FId %>');">정산확인
      	<img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	-->
      	<% end if %>
    </td>
    <td>
    <a href="javascript:PopDetail('<%= ojungsanTaxMM.FItemList(i).FId %>','<%= ojungsanTaxMM.FItemList(i).FtargetGbn%>','<%= ojungsanTaxMM.FItemList(i).Fgroupid %>');">보기<img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom"></a>
    </td>
</tr>
<% next %>
<% else %>
<tr align="center" bgcolor="#FFFFFF" height="30">
<td align="center" colspan="13">검색결과가 없습니다.</td>
</tr>
<% end if %>
</table>

<% end if %>


<p><br><br>

<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="4" align="left">■ 수수료정산내역</td>
</tr>
</table>

<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td colspan="2" width="240">구분</td>
    <td>내 용</td>
    <td width="200">기타</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td width="200">고객판매금액(협력사매출액)</td>
    <td width="40" align="center" >(A)</td>
    <td >협력사가 텐바이텐 사이트를 통해 판매한 매출액(부가세신고시 매출신고금액)</td>
    <td width="220">계산서 발행하지 않음</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>수수료</td>
    <td align="center">(B)</td>
    <td>판매대행수수료 ㈜텐바이텐 매출액(텐바이텐>>협력사로 세금계산서 발행)</td>
    <td>세금계산서 발행</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>배송비/기타판매금액</td>
    <td align="center">(C)</td>
    <td>텐바이텐으로 입금된 배송비 + 기타정산분	</td>
    <td>계산서 발행하지 않음</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>지급대상액(상품)</td>
    <td align="center">(D)</td>
    <td>(D)=(A)-(B) 상품판매에 금액에 대한 매출액-수수료</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>지급대상액(배송비/기타)</td>
    <td align="center">(E)</td>
    <td>텐바이텐에서 협력사로 지급해야할 배송비 + 기타정산분</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>지급예정액</td>
    <td align="center">(F)</td>
    <td>(F)=(A)-(B)+(E) 협력사로 지급할 총액(업체정산액)</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>계산서상태</td>
    <td align="center">(G)</td>
    <td>협력사 검토 및 확인 >> 정산확정 세금계산서	익월 5일 일괄발행됨</td>
    <td>익월 5일 일괄발행됨</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>계산서</td>
    <td align="center">(H)</td>
    <td>텐바이텐>>협력사 발행된 세금계산서 실물 확인 및 출력</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>상세조회</td>
    <td align="center">(I)</td>
    <td>정산에 대한 상세내역 조회</td>
    <td></td>
</tr>
</table>
<p>

<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="4" align="left">■ 기타정산내역</td>
</tr>
</table>

<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td colspan="2" width="240">구분</td>
    <td>내 용</td>
    <td width="200">기타</td>
</tr>

<tr align="left" bgcolor="#FFFFFF">
    <td>프로모션 (협력사 부담)</td>
    <td align="center">(B)</td>
    <td>업체 부담 프로모션 비용</td>
    <td>세금계산서 발행</td>
</tr>

<tr align="left" bgcolor="#FFFFFF">
    <td>지급예정액</td>
    <td align="center">(F)</td>
    <td>정산에서 차감할 금액</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>계산서상태</td>
    <td align="center">(G)</td>
    <td>협력사 검토 및 확인 >> 정산확정 세금계산서	익월 5일 일괄발행됨</td>
    <td>익월 5일 일괄발행됨</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>계산서</td>
    <td align="center">(H)</td>
    <td>텐바이텐>>협력사 발행된 세금계산서 실물 확인 및 출력</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>상세조회</td>
    <td align="center">(I)</td>
    <td>정산에 대한 상세내역 조회</td>
    <td></td>
</tr>
</table>
<p>


<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="4" align="left">■ 매입정산내역</td>
</tr>
</table>

<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td colspan="2" width="240">구분</td>
    <td>내 용</td>
    <td width="200">기타</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td width="200">협력사 상품공급액</td>
    <td width="40" align="center" >(a)</td>
    <td >협력사에서 텐바이텐으로 공급한 상품가액</td>
    <td width="220"></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>배송비/기타</td>
    <td align="center">(b)</td>
    <td>텐바이텐으로 입금된 배송비 + 기타정산분</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>지급대상액(상품)</td>
    <td align="center">(c)</td>
    <td>상품공급에 대한 정산액</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>지급대상액(배송비/기타)</td>
    <td align="center">(d)</td>
    <td>텐바이텐에서 협력사로 지급해야할 배송비 + 기타정산분</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>협력사매출액(지급예정액)</td>
    <td align="center">(e)</td>
    <td>(e)=(c)+(d) 협력사로 지급할 총액(업체정산액)</td>
    <td>세금계산서 발행(협력사>>텐바이텐)</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>계산서상태</td>
    <td align="center">(f)</td>
    <td>업체확인 후 세금계산서 발행시 : 정산확정 / 미발행시 : 업체확인대기</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>계산서</td>
    <td align="center">(g)</td>
    <td>협력사>>텐바이텐 발행된 세금계산서 실물 확인 및 출력</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>상세조회</td>
    <td align="center">(h)</td>
    <td>정산에 대한 상세내역 조회</td>
    <td></td>
</tr>
</table>
<p>

<%
set ojungsanTaxCC = Nothing
set ojungsanTaxCE = Nothing
set ojungsanTaxMM = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
