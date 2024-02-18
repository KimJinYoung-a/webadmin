<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/jungsan/jungsanTaxCls.asp"-->
<%
dim makerid, yyyy1,mm1, jgubun, targetGbn, groupid, page, finishflag, taxtype, jungsan_date, jacctcd
dim noincTen
dim searchType, searchText

page    		= requestCheckvar(request("page"),10)
makerid 		= requestCheckvar(request("makerid"),32)
yyyy1   		= requestCheckvar(request("yyyy1"),10)
mm1     		= requestCheckvar(request("mm1"),10)
jgubun  		= requestCheckvar(request("jgubun"),10)
targetGbn		= requestCheckvar(request("targetGbn"),10)
groupid  		= requestCheckvar(request("groupid"),10)
finishflag 		= requestCheckvar(request("finishflag"),10)
taxtype   		= requestCheckvar(request("taxtype"),10)
jungsan_date 	= requestCheckvar(request("jungsan_date"),10)
jacctcd    		= requestCheckvar(request("jacctcd"),10)
noincTen    	= requestCheckvar(request("noincTen"),10)
searchType 		= requestCheckVar(request("searchType"), 32)
searchText 		= requestCheckVar(request("searchText"), 32)

if (page="") then page=1

if (yyyy1="") then
    yyyy1 = LEFT(dateadd("m",-1,now()),4)
    mm1 = MID(dateadd("m",-1,now()),6,2)
end if

if (jgubun="") then
    jgubun = "MM"
end if

dim ojungsanTax
set ojungsanTax = new CUpcheJungsanTax
ojungsanTax.FPageSize = 30
ojungsanTax.FCurrPage = page
ojungsanTax.FRectMakerid = makerid
ojungsanTax.FRectYYYYMM = yyyy1+"-"+mm1
ojungsanTax.FRectJGubun = jgubun
ojungsanTax.FRectTargetGbn = targetGbn
ojungsanTax.FRectGroupid = groupid
ojungsanTax.FRectFinishFlag = finishflag
ojungsanTax.FRectTaxType = taxtype
ojungsanTax.FRectJungsanDate = jungsan_date
ojungsanTax.FRectjacctcd = jacctcd
ojungsanTax.FRectNotIncTen = noincTen
ojungsanTax.FRectSearchType = searchType
ojungsanTax.FRectSearchText = searchText
ojungsanTax.getMonthUpcheJungsanListAdmAll


dim i
%>
<script language='javascript'>

function NextPage(page){
    var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}

function PopDetail(iidx,tg,makerid){
    var uri = 'jungsandetailsumONAdm.asp?id=' + iidx;
    if (tg=="OF") uri = 'jungsandetailsumOFAdm.asp?idx=' + iidx;
	var popwin = window.open(uri+'&makerid='+makerid,'PopDetail','width=1300, height=900, scrollbars=1, resizable=yes');
	popwin.focus();
}

function PopTaxRegPrdCommission(makerid, yyyy1, mm1, onoffGubun, jidx) {
	var popwin = window.open("popTaxRegAdmin.asp?makerid=" + makerid + "&yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&onoffGubun=" + onoffGubun + "&jidx="+jidx,"PopTaxRegPrdCommission","width=640 height=700 scrollbars=yes resizable=yes");
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

function XLDown(){
	var frm = document.frm;
	var page = frm.xlpage.value;
	if (page*0 != 0) { page = 1; }

    var paramURL = 'monthjungsanAdmAllXL.asp?yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&makerid=<%=makerid%>&jgubun=<%=jgubun%>&targetGbn=<%=targetGbn%>&groupid=<%=groupid%>&finishflag=<%=finishflag%>&taxtype=<%=taxtype%>&jacctcd=<%=jacctcd%>&noincTen=<%=noincTen%>&page=' + page;

    var popwin = window.open(paramURL,'monthjungsanAdmAllXL','width=100,height=100,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function CSetcXLDown(){
    var paramURL = 'monthcsjungsanXL.asp?yyyy1=<%=yyyy1%>&mm1=<%=mm1%>';
    var popwin2 = window.open(paramURL,'monthcsjungsanXL','width=100,height=100,scrollbars=yes,resizable=yes');
    popwin2.focus();
}

<% if (ojungsanTax.FresultCount>0) then %>
//alert('2014년 1월 정산부터 수수료 정산분에 대해서는\n\n텐바이텐에서 계산서를 발행 하오니\n\n이세로 등을 통해 따로 발행하지 말아 주시길 바랍니다.');
<% end if %>
</script>
<H1>수정중</H1>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		정산 대상 년월 :&nbsp;<% DrawYMBox yyyy1,mm1 %>
		&nbsp;
		브랜드ID : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
		&nbsp;&nbsp;
        업체(그룹코드) : <input type="text" class="text" name="groupid" value="<%= groupid %>" size="12" >
        &nbsp;&nbsp;
        계정과목코드 : <input type="text" class="text" name="jacctcd" value="<%= jacctcd %>" size="7" >
        &nbsp;&nbsp;
        <input type="checkbox" name="noincTen" <%=CHKIIF(noincTen="on","checked","")%> >(텐바이텐211-87-00620 사업자 제외)
		<% If True or (jgubun = "CC") Then %>
		<input type="button" value="CS기타정산XL다운" onClick="CSetcXLDown()" class="button">
		<% End If %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
        정산방식구분 :
        <% drawSelectBoxJGubun "jgubun",jgubun %>
        &nbsp;&nbsp;
		과세구분
		<select name="taxtype" >
		<option value="">전체
		<option value="01" <%= CHKIIF(taxtype="01","selected","") %> >과세
		<option value="02" <%= CHKIIF(taxtype="02","selected","") %> >면세
		<option value="03" <%= CHKIIF(taxtype="03","selected","") %> >간이
		</select>


        &nbsp;
        매출처 구분 :
        <select name="targetGbn" >
		<option value="">전체
		<option value="ON" <%= CHKIIF(targetGbn="ON","selected","") %> >ON
		<option value="OF" <%= CHKIIF(targetGbn="OF","selected","") %> >OF
		<option value="AC" <%= CHKIIF(targetGbn="AC","selected","") %> >AC
		</select>
		&nbsp;
		상태
		<select name="finishflag" >
		<option value="">전체
		<option value="0" <%= CHKIIF(finishflag="0","selected","") %> >수정중
		<option value="1" <%= CHKIIF(finishflag="1","selected","") %> >업체확인대기
		<option value="2" <%= CHKIIF(finishflag="2","selected","") %> >업체확인완료
		<option value="3" <%= CHKIIF(finishflag="3","selected","") %> >정산확정
		<option value="7" <%= CHKIIF(finishflag="7","selected","") %> >입금완료
		</select>
        &nbsp;
        정산일 :
        <select name="jungsan_date">
        <option value="" <% if jungsan_date="" then response.write "selected" %> >선택
        <option value="15일" <% if jungsan_date="15일" then response.write "selected" %> >15일
        <option value="말일" <% if jungsan_date="말일" then response.write "selected" %> >말일
        <option value="수시" <% if jungsan_date="수시" then response.write "selected" %> >수시
        </select>

		&nbsp;&nbsp;&nbsp;
		<input type="text" class="text" name="xlpage" value="1" size="1">페이지
		<input type="button" value="XL다운(2500건)" onClick="XLDown()">
		&nbsp;&nbsp;&nbsp;
		검색조건:
		<select class="select" name="searchType">
			<option></option>
			<option value="socname" <% if (searchType = "socname") then %>selected<% end if %> >업체명</option>
			<option value="socno" <% if (searchType = "socno") then %>selected<% end if %> >사업자번호</option>
		</select>
		&nbsp;
		<input type="text" class="text" name=searchText value="<%= searchText %>" size="15" maxlength="20">
    </td>
</tr>
</form>
</table>
<p>


<% if (jgubun="CC") then %>
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="21" align="left"><strong>* 수수료 정산 내역</strong> <font color=red>(수수료 세금 계산서는 <b>텐바이텐</b>에서 <b>일괄 발행</b>합니다.)</font></td>
    <td colspan="2" align="right">총 <%=ojungsanTax.FTotalcount%> 건 <%=page%> / <%=ojungsanTax.FTotalpage%></td>
</tr>
<!--

-->
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td width="50" rowspan="2">정산월</td>
    <td width="50" rowspan="2">매출처</td>
    <td width="40" rowspan="2">과세<br>구분</td>
    <td width="50" rowspan="2">그룹코드</td>
    <td width="50" rowspan="2">ERPCode</td>
    <td width="90" rowspan="2">회사명</td>
    <td width="90" rowspan="2">사업자번호</td>
    <td width="90" rowspan="2">브랜드ID</td>
    <td width="180" rowspan="2">정산내역</td>
    <td width="90" rowspan="2">계정과목</td>
    <td width="90" rowspan="2">고객판매금액<br>(협력사 매출액)</td>
    <td width="80" rowspan="2">상품판매<br>수수료</td>
    <td width="80" rowspan="2">결제대행<br>수수료</td>
    <td width="80" rowspan="2">배송비<br>(판매금액)</td>
    <td width="100" rowspan="2">지급예정액<br>(상품)</td>
  	<td width="80" rowspan="2">지급예정액<br>(배송비)</td>
  	<td colspan="3" align="center">추가정산액</td>
  	<td width="80" rowspan="2">지급예정액</td>
    <td width="90" rowspan="2">계산서상태</td>
    <td width="80" rowspan="2">계산서</td>
    <td rowspan="2">상세조회</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td width="80" >(추가배송비)</td>
  	<td width="80" >(반품배송비등)</td>
  	<td width="80" >(기타프로모션등)</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td>합계</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td align="left"></td>
    <td align="left"></td>
    <td align="left"></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FPrdMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FPrdCommissionSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FPgCommissionSum,0)%></td><!--결제대행 수수료-->
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FdlvMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FprdJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.getOriginDlvJungsanSum,0)%></td><!-- 지급예정액<br>(배송비)==배송비<br>(판매금액) -->
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.getAddDlvJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.getEtcDlvJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.getPromotionJungsanSum,0)%></td>
    <td align="right">
        <% if ojungsanTax.FSumaryOneItem.getCalcuToTalJungsanSum<>ojungsanTax.FSumaryOneItem.getToTalJungsanSum then %>
        <b><font color=red><%=FormatNumber(ojungsanTax.FSumaryOneItem.getToTalJungsanSum,0)%></font></b><br>(<%=formatNUmber(ojungsanTax.FSumaryOneItem.getCalcuToTalJungsanSum-ojungsanTax.FSumaryOneItem.getToTalJungsanSum,0)%>)
        <% else %>
        <%=FormatNumber(ojungsanTax.FSumaryOneItem.getToTalJungsanSum,0)%>
        <% end if %>
    </td>
    <td></td>
    <td></td>
    <td></td>
</tr>
<% for i=0 to ojungsanTax.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td><%=ojungsanTax.FItemList(i).Fyyyymm%></td>
    <td><%=ojungsanTax.FItemList(i).getTargetNm%></td>
    <td><%=ojungsanTax.FItemList(i).getItemVatTypeName%></td>
    <td ><%=ojungsanTax.FItemList(i).Fgroupid%></td>
    <td ><%=ojungsanTax.FItemList(i).FerpCust_cd%></td>
    <td ><%=ojungsanTax.FItemList(i).Fcompany_name%></td>
    <td ><%=ojungsanTax.FItemList(i).Fcompany_no%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Fmakerid%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Ftitle%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Facc_nm%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FPrdMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FPrdCommissionSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FPgCommissionSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FdlvMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FprdJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).getOriginDlvJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).getAddDlvJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).getEtcDlvJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).getPromotionJungsanSum,0)%></td>
    <td align="right">
        <% if ojungsanTax.FItemList(i).getCalcuToTalJungsanSum<>ojungsanTax.FItemList(i).getToTalJungsanSum then %>
        <b><font color=red><%=FormatNumber(ojungsanTax.FItemList(i).getToTalJungsanSum,0)%></font></b>
        <% else %>
        <%=FormatNumber(ojungsanTax.FItemList(i).getToTalJungsanSum,0)%>
        <% end if %>
    </td>
    <td><%=ojungsanTax.FItemList(i).GetTaxEvalStateName%></td>
    <td>
        <% if ojungsanTax.FItemList(i).IsElecTaxExists then %>
      	<a href="javascript:PopTaxPrintReDirect('<%= ojungsanTax.FItemList(i).Fneotaxno %>');">출력
      	<img src="/images/icon_print02.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
		<% else %>
      	<a href="javascript:PopTaxRegPrdCommission('<%=ojungsanTax.FItemList(i).Fmakerid %>', '<%= yyyy1 %>', '<%= mm1 %>', '<%= ojungsanTax.FItemList(i).FtargetGbn %>','<%= ojungsanTax.FItemList(i).Fid %>');">발행
      	<img src="/images/icon_new.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	<% end if %>
    </td>
    <td>
    <a href="javascript:PopDetail('<%= ojungsanTax.FItemList(i).FId %>','<%= ojungsanTax.FItemList(i).FtargetGbn%>','<%=ojungsanTax.FItemList(i).Fmakerid%>');">보기<img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom"></a>
    </td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="23" align="center">
        <% if ojungsanTax.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ojungsanTax.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ojungsanTax.StartScrollPage to ojungsanTax.FScrollCount + ojungsanTax.StartScrollPage - 1 %>
			<% if i>ojungsanTax.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ojungsanTax.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
    </td>
</tr>
</table>

<% else %>

<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="17" align="left"><strong>* 매입 정산 내역</strong> (협력사에서 텐바이텐으로 발행해 주셔야 합니다.) (롯데닷컴 판매 내역 및 가맹점 판매 내역은 매입정산으로 처리 됩니다.)</td>
    <td colspan="2" align="right">총 <%=ojungsanTax.FTotalcount%> 건 <%=page%> / <%=ojungsanTax.FTotalpage%></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td width="60" >정산월</td>
    <td width="60" >매출처</td>
    <td width="50" >과세<br>구분</td>
    <td width="50" >그룹코드</td>
    <td width="50" >ERPCode</td>
    <td width="90" >회사명</td>
    <td width="90" >사업자번호</td>
    <td width="90" >브랜드ID</td>
    <td width="170" >정산내역</td>
    <td width="90" >계정과목</td>
    <td width="90" >입고분매입<br>(상품공급액)</td>
    <td width="90" >판매분매입<br>(상품공급액)</td>
    <td width="90" >기타매입<br>(강좌)</td>
    <td width="90" >기타매입<br>(배송비)</td>
    <td width="90" >기타출고매입<br>(기타출고)</td>
    <td width="90" >지급예정액<br>(협력사매출액)</td>
    <!--
    <td width="90" >협력사<br>상품공급액</td>
    <td width="80" >배송비/기타</td>
    <td width="100">지급대상액<br>(상품)</td>
  	<td width="80">지급대상액<br>(배송비/기타)</td>
  	<td width="80">협력사매출액<br>(지급예정액)</td>
  	-->
    <td width="90" >계산서상태</td>
    <td width="80" >계산서</td>
    <td >상세조회</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
    <td>합계</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td align="left"></td>
    <td align="left"></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FMSuply,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FSSuply,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FESuply,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FDSuply,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FCSuply,0)%></td>
    <!--
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FPrdMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FdlvMeachulsum + ojungsanTax.FSumaryOneItem.FetMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FprdJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FdlvJungsanSum + ojungsanTax.FSumaryOneItem.FetJungsanSum,0)%></td>
    -->
    <td align="right">
        <% if ojungsanTax.FSumaryOneItem.getCalcuToTalJungsanSum<>ojungsanTax.FSumaryOneItem.getToTalJungsanSum then %>
        <b><font color=red><%=FormatNumber(ojungsanTax.FSumaryOneItem.getToTalJungsanSum,0)%></font></b><br>(<%=formatNUmber(ojungsanTax.FSumaryOneItem.getCalcuToTalJungsanSum-ojungsanTax.FSumaryOneItem.getToTalJungsanSum,0)%>)
        <% else %>
        <%=FormatNumber(ojungsanTax.FSumaryOneItem.getToTalJungsanSum,0)%>
        <% end if %>
    </td>
    <td></td>
    <td></td>
    <td></td>
</tr>
<% for i=0 to ojungsanTax.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td><%=ojungsanTax.FItemList(i).Fyyyymm%></td>
    <td><%=ojungsanTax.FItemList(i).getTargetNm%></td>
    <td><%=ojungsanTax.FItemList(i).getTaxtypeName%></td>
    <td ><%=ojungsanTax.FItemList(i).Fgroupid%></td>
    <td ><%=ojungsanTax.FItemList(i).FerpCust_cd%></td>
    <td ><%=ojungsanTax.FItemList(i).Fcompany_name%></td>
    <td ><%=ojungsanTax.FItemList(i).Fcompany_no%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Fmakerid%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Ftitle%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Facc_nm%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FMSuply,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FSSuply,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FESuply,0)%></td>
   <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FDSuply,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FCSuply,0)%></td>
    <!--
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FPrdMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FdlvMeachulsum + ojungsanTax.FItemList(i).FetMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FprdJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FdlvJungsanSum + ojungsanTax.FItemList(i).FetJungsanSum,0)%></td>
    -->
    <td align="right">
        <% if ojungsanTax.FItemList(i).getCalcuToTalJungsanSum<>ojungsanTax.FItemList(i).getToTalJungsanSum then %>
        <b><font color=red><%=FormatNumber(ojungsanTax.FItemList(i).getToTalJungsanSum,0)%></font></b>
        <% else %>
        <%=FormatNumber(ojungsanTax.FItemList(i).getToTalJungsanSum,0)%>
        <% end if %>
    </td>
    <td><%=ojungsanTax.FItemList(i).GetTaxEvalStateName%></td>
    <td>
        <% if ojungsanTax.FItemList(i).IsElecTaxExists then %>
      	<a href="javascript:PopTaxPrintReDirect('<%= ojungsanTax.FItemList(i).Fneotaxno %>');">출력
      	<img src="/images/icon_print02.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	<% elseif ojungsanTax.FItemList(i).IsCommissionTax then %>
      	</a>
      	<% elseif ojungsanTax.FItemList(i).IsElecTaxCase then %>
      	<!--
      	<a href="javascript:PopTaxReg<%=CHKIIF(ojungsanTax.FItemList(i).FtargetGbn="OF","Off","")%>('<%= ojungsanTax.FItemList(i).FId %>');">발행
      	<img src="/images/icon_new.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	-->
      	<% elseif ojungsanTax.FItemList(i).IsElecFreeTaxCase then %>
      	<!--
      	<a href="javascript:PopTaxReg('<%= ojungsanTax.FItemList(i).FId %>');">발행
      	<img src="/images/icon_new.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	-->
      	<% elseif ojungsanTax.FItemList(i).IsElecSimpleBillCase then %>
      	<!--
      	<a href="javascript:PopConfirm('<%= menupos %>','<%= ojungsanTax.FItemList(i).FId %>');">정산확인
      	<img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	-->
      	<% end if %>
    </td>
    <td>
    <a href="javascript:PopDetail('<%= ojungsanTax.FItemList(i).FId %>','<%= ojungsanTax.FItemList(i).FtargetGbn%>','<%=ojungsanTax.FItemList(i).Fmakerid%>');">보기<img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom"></a>
    </td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="19" align="center">
        <% if ojungsanTax.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ojungsanTax.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ojungsanTax.StartScrollPage to ojungsanTax.FScrollCount + ojungsanTax.StartScrollPage - 1 %>
			<% if i>ojungsanTax.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ojungsanTax.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
    </td>
</tr>
</table>
<% end if %>
<%
set ojungsanTax = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
