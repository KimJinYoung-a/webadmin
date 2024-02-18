<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/3pl/jungsanCls.asp"-->
<%

dim tplcompanyid, yyyy1, mm1, page, finishflag, taxtype, research
research 		= requestCheckVar(request("research"),32)
tplcompanyid 	= requestCheckVar(request("tplcompanyid"),32)
yyyy1    		= requestCheckVar(request("yyyy1"),4)
mm1      		= requestCheckVar(request("mm1"),2)
page     		= requestCheckVar(request("page"),10)
finishflag     	= requestCheckVar(request("finishflag"),10)
taxtype     	= requestCheckVar(request("taxtype"),10)

dim i


dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

if page="" then page=1

dim otpljungsan
set otpljungsan = new CTplJungsan
otpljungsan.FPageSize  = 100
otpljungsan.FCurrPage  = page
otpljungsan.FRectYYYYMM = yyyy1 + "-" + mm1

otpljungsan.FRectTplCompanyID = tplcompanyid
otpljungsan.FRectCancelYN = "N"

otpljungsan.GetTPLJungsanMasterList

%>
<script language='javascript'>
function NextPage(ipage) {
    document.frm.page.value=ipage;
    document.frm.submit();
}

function research(frm) {
	frm.submit();
}

function MakeTplBatchJungsan(frm){
    if (frm.differencekey.value.length<1){
        alert('차수 구분을 선택 하세요.');
        frm.differencekey.focus();
        return;
    }

    if (confirm('정산내역을 작성 하시겠습니까?')){
        var queryurl = 'dotpljungsan.asp?mode=tplbatchprocess&tplcompanyid=' + frm.tplcompanyid.value + '&yyyy1=' + frm.yyyy.value + '&mm1=' + frm.mm.value + '&differencekey=' + frm.differencekey.value;
        var popwin = window.open(queryurl ,'on_jungsan_process','width=200, height=200, scrollbars=yes, resizable=yes');
    }
}

function dellThis(v) {
	var upfrm = document.frmarr;
	var ret = confirm('모든 정산 데이터를 삭제 하시겠습니까?');
	if (ret){
		upfrm.masteridx.value = v;
		upfrm.mode.value = "dellall";
		upfrm.submit();
	}
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
   	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">

   	<tr align="center" bgcolor="#FFFFFF" >
        <td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
        <td align="left">
	        	정산대상년월 : <% DrawYMBox yyyy1,mm1 %>&nbsp;&nbsp;
				업체ID : <input type="text" class="text" name="tplcompanyid" value="<%= tplcompanyid %>" size="20" >&nbsp;&nbsp;
	        </td>
        <td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
    		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
    	</td>
    </tr>
	<tr>
        <td bgcolor="#FFFFFF" >
			상태
			<select class="select" name="finishflag" >
			<option value="">전체
			<option value="0" <%= CHKIIF(finishflag="0","selected","") %> >수정중
			<option value="1" <%= CHKIIF(finishflag="1","selected","") %> >업체확인대기
			<option value="2" <%= CHKIIF(finishflag="2","selected","") %> >업체확인완료
			<option value="3" <%= CHKIIF(finishflag="3","selected","") %> >정산확정
			<option value="7" <%= CHKIIF(finishflag="7","selected","") %> >입금완료
			</select>
			&nbsp;&nbsp;
			계산서과세구분
			<select class="select" name="taxtype" >
			<option value="">전체
			<option value="01" <%= CHKIIF(taxtype="01","selected","") %> >과세
			<option value="02" <%= CHKIIF(taxtype="02","selected","") %> >면세
			<option value="03" <%= CHKIIF(taxtype="03","selected","") %> >원천
			</select>
        </td>
    </tr>
	</form>
</table>
<!-- 표 상단바 끝-->

<p />

작업중!!

<p />

<% if (tplcompanyid<>"") and (yyyy1<>"") and (mm1<>"") then %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="tplbatch" >
<input type="hidden" name="tplcompanyid" value="<%= tplcompanyid %>">
<input type="hidden" name="yyyy" value="<%= yyyy1 %>">
<input type="hidden" name="mm" value="<%= mm1 %>">
<tr bgcolor="#FFFFFF">
    <td>
        <select class="select" name="differencekey">
            <option value="">차수 선택</option>
            <option value="0">0차</option>
            <option value="1">1차</option>
            <option value="2">2차</option>
            <option value="3">3차</option>
            <option value="4">4차</option>
            <option value="5">5차</option>
            <option value="6">6차</option>
            <option value="7">7차</option>
            <option value="8">8차</option>
            <option value="9">9차</option>
        </select>

        <input type="hidden" name="ipchulArr" value="">
        <input type="button" class="button" value=" <%= tplcompanyid %> &nbsp;&nbsp;<%= yyyy1 %>년 <%= mm1 %>월 정산 작성 " onClick="MakeTplBatchJungsan(document.tplbatch);">
    </td>
</form>
</tr>
</table>
<% end if %>

<p />

<form name="frmList" method="post" action="dotpljungsan.asp">
<input type="hidden" name="idxarr" value="">
<input type="hidden" name="mode" value="multistatechange">
<table width="100%" align="center" border="0" cellpadding="1" cellspacing="1" class="a" bgcolor=#BABABA>
    <tr bgcolor="#FFFFFF" height="25">
      <td colspan="30" >
      <%= page %>/<%= otpljungsan.FTotalPage %> page 총 <%=otpljungsan.FTotalCount %>건
      </td>
    </tr>
    <tr align="center" bgcolor="#DDDDFF" height="25">
      <td width="70">정산월</td>
      <td width="30">차수</td>
      <td width="30">과세<br>(계산서)</td>
      <td width="90"><a href="javascript:research(frm,'designer')">업체ID</a></td>
      <td>회사명</td>
      <td width="80">임대비용</td>
      <td width="80">입출고비용</td>
      <td width="80">기타비용</td>
      <td width="80">정산액</td>
      <td width="80">VAT</td>
      <td width="80">청구금액</td>
      <td width="80"><a href="javascript:research(frm,'state')">상태</a></td>
      <td width="70">세금계산서<br>등록일</td>
      <td width="70"><a href="javascript:research(frm,'segum')">세금발행일</a></td>
      <td width="70">입금일</td>
      <td width="20">E</td>
      <td width="20">S</td>
      <td width="50"><a href="javascript:research(frm,'tax')">과세구분</a></td>
      <td width="30">정산</td>
      <td width="30">비고</td>
    </tr>
<% if otpljungsan.FResultCount<1 then %>
    <tr align="center" bgcolor="#FFFFFF">
        <td colspan="30" align="center" height="30">
        <% if research="" then %>
            [검색 버튼을 눌러주세요.]
        <% else %>
            [검색 결과가 없습니다.]
        <% end if %>
        </td>
    </tr>
<% else %>
    <% for i=0 to otpljungsan.FResultCount-1 %>
   <tr align="center" bgcolor="#FFFFFF" height="25">
      <td ><a target=_blank href="nowjungsanmasteredit.asp?id=<%= otpljungsan.FItemList(i).FIdx %>"><%= otpljungsan.FItemList(i).Fyyyymm %>&nbsp;<img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom"></a></td>
      <td ><%= otpljungsan.FItemList(i).Fdifferencekey %></td>
      <td ><font color="<%= otpljungsan.FItemList(i).GetTaxtypeNameColor %>"><%= otpljungsan.FItemList(i).GetSimpleTaxtypeName %></font></td>
      <td ><a href="javascript:PopBrandInfoEdit('<%= Replace(otpljungsan.FItemList(i).Ftplcompanyid, "tpl", "3pl") %>')"><%= otpljungsan.FItemList(i).Ftplcompanyid %></a></td>
      <td align="left"><a href="javascript:PopUpcheInfoEdit('<%= otpljungsan.FItemList(i).FGroupID %>')"><%= otpljungsan.FItemList(i).Fcompany_name %></a></td>
      <td align="right"><a target=_blank href="tpljungsandetail.asp?idx=<%= otpljungsan.FItemList(i).FIdx %>&gubun=cbm"><%= FormatNumber(otpljungsan.FItemList(i).Fst_totalcash,0) %></a></td>
      <td align="right"><a target=_blank href="tpljungsandetail.asp?idx=<%= otpljungsan.FItemList(i).FIdx %>&gubun=ipchul"><%= FormatNumber(otpljungsan.FItemList(i).Fio_totalcash,0) %></a></td>
      <td align="right"><a target=_blank href="tpljungsandetail.asp?idx=<%= otpljungsan.FItemList(i).FIdx %>&gubun=etc"><%= FormatNumber(otpljungsan.FItemList(i).Fet_totalcash,0) %></a></td>

      <td align="right"><%= FormatNumber(0,0) %></td>
      <td align="right"><%= FormatNumber(0,0) %></td>
      <td align="right"><%= FormatNumber(0,0) %></td>

      <td ><font color="<%= otpljungsan.FItemList(i).GetStateColor %>"><%= otpljungsan.FItemList(i).GetStateName %></font>
	  <% if otpljungsan.FItemList(i).Ffinishflag="0" then %>
      <a href="javascript:NextStep('<%= otpljungsan.FItemList(i).FIdx %>');">
     <img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom">
      </a>
      <% end if %>
      </td>
	    <% if IsNULL(otpljungsan.FItemList(i).Ftaxinputdate) then %>
	    <td ></td>
  	    <% else %>
 	    <td ><%= Left(Cstr(otpljungsan.FItemList(i).Ftaxinputdate),10) %></td>
  	    <% end if %>
      <% if isNull(otpljungsan.FItemList(i).Ftaxregdate) then %>
      <td ></td>
      <% else %>
      <td ><%= Left(Cstr(otpljungsan.FItemList(i).Ftaxregdate),10) %></td>
      <% end if %>
      <% if isNull(otpljungsan.FItemList(i).Fipkumdate) then %>
      <td ></td>
      <% else %>
      <td ><%= Left(Cstr(otpljungsan.FItemList(i).Fipkumdate),10) %></td>
      <% end if %>
      <td ><a href="javascript:PopCSMailSend('<%= otpljungsan.FItemList(i).Fjungsan_email %>','','');"><% if otpljungsan.FItemList(i).Fjungsan_email<>"" then response.write "E" %></a></td>
      <td ><a href="javascript:PopCSSMSSend('<%= otpljungsan.FItemList(i).Fjungsan_hp %>','','','');"><% if otpljungsan.FItemList(i).Fjungsan_hp<>"" then response.write "S" %></a></td>
      <td ><%= otpljungsan.FItemList(i).Fjungsan_gubun %></td>
      <td ><%= otpljungsan.FItemList(i).Fjungsan_date %></td>
      <% if otpljungsan.FItemList(i).Ffinishflag="0" then %>
      	<td ><a href="javascript:dellThis('<%= otpljungsan.FItemList(i).FIdx %>')">x</a></td>
      <% else %>
        <td >
            <% if Not IsNULL(otpljungsan.FItemList(i).FTaxLinkidx) then %>
      	        <img src="/images/icon_print02.gif" width="14" height="14" border=0 onclick="window.open('http://www.bill36524.com/popupBillTax.jsp?NO_TAX=<%= otpljungsan.FItemList(i).Fneotaxno %>&NO_BIZ_NO=2118700620')" style="cursor:hand">
      	   <% else %>
      	        <%= otpljungsan.FItemList(i).Fbillsitecode %>
      	    <% end if %>

      	    <a href="/admin/upchejungsan/monthjungsanAdm.asp?makerid=<%= otpljungsan.FItemList(i).Fdesignerid %>&yyyy1=<%= LEFT(otpljungsan.FItemList(i).Fyyyymm,4) %>&mm1=<%= right(otpljungsan.FItemList(i).Fyyyymm,2) %>" target="_blank">POP</a>
        </td>
      <% end if %>
    </tr>
    <% next %>
<% end if %>
</table>
</form>

<form name="frmarr" method="post" action="dotpljungsan.asp">
<input type="hidden" name="masteridx" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="rd_state" value="">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_TPLclose.asp" -->
