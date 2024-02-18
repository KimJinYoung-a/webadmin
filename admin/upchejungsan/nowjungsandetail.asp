<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->

<%
dim id,gubun,rectorder
dim yyyymm
dim IsCommissionTax ''수수료 정산 여부
id      = requestCheckvar(request("id"),10)
gubun   = requestCheckvar(request("gubun"),16)
rectorder = requestCheckvar(request("rectorder"),16)

if rectorder="" then rectorder="orderserial"

if gubun="" then gubun="upche"


dim sqlStr
'dim isLecture
'sqlStr = "select top 1 m.id,m.designerid,c.userdiv "
'sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master m,"
'sqlStr = sqlStr + " [db_user].[dbo].tbl_user_c c"
'sqlStr = sqlStr + " where m.id="  & id
'sqlStr = sqlStr + " and m.designerid=c.userid"
'
'rsget.Open sqlStr,dbget,1
'if Not rsget.Eof then
'    isLecture = rsget("userdiv")="14"
'end if
'rsget.close

dim isAcademyJungsan

dim ojungsan, ojungsanmaster
set ojungsanmaster = new CUpcheJungsan
ojungsanmaster.FRectId = id
ojungsanmaster.JungsanMasterList

if (ojungsanmaster.FResultCount<1) then
    dbget.Close(): response.end
end if

IsCommissionTax = ojungsanmaster.FItemList(0).IsCommissionTax
isAcademyJungsan = ojungsanmaster.FItemList(0).FtargetGbn="AC"

if (isAcademyJungsan) and (Not IsCommissionTax) and (gubun="upche") then
    gubun="lecture"
end if

set ojungsan = new CUpcheJungsan
ojungsan.FRectid = id
ojungsan.FRectgubun = gubun

'ojungsan.FREctSitename = "N10x10"


if (gubun<>"") and (gubun<>"witakconfirm") then
    if (isAcademyJungsan) then
        ojungsan.JungsanDetailListLectureSum
    else
	    ojungsan.JungsanDetailListWitakSum
	end if
end if

dim i, suplysum, suplytotalsum, selltotalsum
suplysum = 0
suplytotalsum = 0
selltotalsum  = 0


yyyymm = ojungsanmaster.FItemList(0).FYYYYmm

dim duplicated
%>
<script language='javascript'>
function reOrder(comp){
	document.frm.rectorder.value=comp.value;
	document.frm.submit();
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function DelDetail(frm){
	var ret = confirm('선택 내역을 삭제 하시겠습니까?');
	if (ret){
		frm.mode.value="deldetail";
		frm.submit();
	}
}

function ModiDetail(frm){
	var ret = confirm('선택 내역을 수정 하시겠습니까?');
	if (ret){
		frm.mode.value="modidetail";
		frm.submit();
	}
}

function savememo(frm){
	var ret = confirm('메모를 저장하시겠습니까?');
	if (ret){
		frm.mode.value = "memoedit";
		frm.submit();
	}
}

function addEtcList(iid,igubun){
	window.open('popetclistadd.asp?id=' + iid + '&gubun=' + igubun,'popetc','width=700, height=150, location=no,menubar=no,resizable=yes,scrollbars=no,status=no,toolbar=no');
}

function Char2Zero(v){
	if (isNaN(v)){
		return 0
	}else{
		return v ;
	}
}

function ReCalcu(frm){
	var ireal

	ireal = Char2Zero(frm.realjaego.value) * 1 ;
	frm.tmpsysjaego.value = Char2Zero(frm.prejaego.value) * 1 + Char2Zero(frm.ipgono.value) * 1	- Char2Zero(frm.chulgono.value) * 1 - Char2Zero(frm.sellno.value) * 1;
	frm.ocha.value = Char2Zero(frm.tmpsysjaego.value) * 1 - ireal;
	frm.jungsanno.value = Char2Zero(frm.chulgono.value) * 1 + Char2Zero(frm.sellno.value) * 1 + Char2Zero(frm.ocha.value) * 1;

}

function ReSearch(frm){
	//if (frm.gubun[2].checked){
	//	frm.action="nowjungsandetailwitak.asp"
	//}else{
	//	frm.action="nowjungsandetail.asp"
	//}
	frm.submit();
}

function popBatchDetailEdit(id,gubun,itemid,itemoption,sellcash,suplycash,itemname,itemoptionname){
	var popwin = window.open('','jungsandetailedit','width=600,height=200,scrollbars=yes,resizable=yes');
	popwin.focus();

	bufFrm.target="jungsandetailedit";
	bufFrm.id.value = id;
	bufFrm.gubun.value = gubun;
	bufFrm.itemid.value = itemid;
	bufFrm.itemoption.value = itemoption;
	bufFrm.sellcash.value = sellcash;
	bufFrm.suplycash.value = suplycash;
	bufFrm.itemname.value = itemname;
	bufFrm.itemoptionname.value = itemoptionname;

	bufFrm.submit();
}

</script>

<form name="bufFrm" method=post action="popjungsandetailedit.asp">
<input type="hidden" name="id" value=''>
<input type="hidden" name="gubun" value=''>
<input type="hidden" name="itemid" value=''>
<input type="hidden" name="itemoption" value=''>
<input type="hidden" name="sellcash" value=''>
<input type="hidden" name="suplycash" value=''>
<input type="hidden" name="itemname" value=''>
<input type="hidden" name="itemoptionname" value=''>
</form>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="id" value="<%= id %>">
	<input type="hidden" name="rectorder" value="<%= rectorder %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	브랜드ID : <b><%= ojungsanmaster.FItemList(i).Fdesignerid %></b>
        	&nbsp;
			<input type="radio" name="gubun" value="upche" <% if gubun="upche" then response.write "checked" %> >업체배송
			<input type="radio" name="gubun" value="maeip" <% if gubun="maeip" then response.write "checked" %> >매입
			<input type="radio" name="gubun" value="witaksell" <% if gubun="witaksell" then response.write "checked" %> >위탁판매
			<input type="radio" name="gubun" value="witakchulgo" <% if gubun="witakchulgo" then response.write "checked" %> >기타출고
			<input type="radio" name="gubun" value="DL" <% if gubun="DL" then response.write "checked" %> >배송비
			<input type="radio" name="gubun" value="DT" <% if gubun="DT" then response.write "checked" %> >추가배송비
			<input type="radio" name="gubun" value="DP" <% if gubun="DP" then response.write "checked" %> >기타(프로모션)
			<input type="radio" name="gubun" value="DE" <% if gubun="DE" then response.write "checked" %> >기타(보정)
		<input type="radio" name="gubun" value="lecture" <% if gubun="lecture" then response.write "checked" %> >강좌
        </td>
        <td align="right">
        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->

<% if gubun<>"witakconfirm" then %>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	<img src="/images/icon_arrow_down.gif" align="absbottom">
			<font color="red"><strong>아이템 합계</strong></font>
			(내역을 수정하면 합계에 적용됩니다.)
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">상품코드</td>
		<td>상품명</td>
		<td width="80">옵션명</td>
		<td width="40">수량</td>
		<td width="50"><font color="#AAAAAA">판매가(현재)</font></td>
		<td width="50"><font color="#AAAAAA">공급가(현재)</font></td>
		<td width="50"><font color="#AAAAAA">옵션(현재)</font></td>
		<td width="50"><font color="#AAAAAA">옵션공급가(현재)</font></td>
		<td width="80">판매가</td>
		<td width="60">현재가대비<br>할인</td>
		<td width="80">실매출</td>
		<td width="80">수수료</td>
		<td width="80">PG수수료</td>
		<td width="80">공급가</td>
		<td width="40">마진</td>
		<td width="80">공급가합계</td>
		<!-- td width="40">일괄<br>수정</td -->
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    suplysum =0
    suplysum = suplysum + ojungsan.FItemList(i).Fsuplycash * ojungsan.FItemList(i).FItemNo
    suplytotalsum = suplytotalsum + suplysum
    selltotalsum = selltotalsum + ojungsan.FItemList(i).Fsellcash * ojungsan.FItemList(i).FItemNo

    duplicated = ojungsan.CheckDuplicated(i)
    %>

	<% if duplicated then %>
	<tr bgcolor="#EEEEEE">
	<% else %>
    <tr bgcolor="#FFFFFF">
    <% end if %>
      <td ><%= ojungsan.FItemList(i).FItemID %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= ojungsan.FItemList(i).FItemNo %></td>
      <% if ojungsan.FItemList(i).FOrgsellcash+ojungsan.FItemList(i).FOrgOptaddprice<>ojungsan.FItemList(i).Fsellcash then %>
          <td align="right">
          <font color="#FF0000"><b><%= FormatNumber(ojungsan.FItemList(i).FOrgsellcash,0) %></b></font>
          </td>
      <% else %>
          <td align="right">
          <% if Not IsNULL(ojungsan.FItemList(i).FOrgsellcash) then %>
          <font color="#AAAAAA"><%= FormatNumber(ojungsan.FItemList(i).FOrgsellcash,0) %></font>
          <% end if %>
          </td>
      <% end if %>

      <% if ojungsan.FItemList(i).FOrgsuplycash+ojungsan.FItemList(i).FOrgOptaddbuyprice<>ojungsan.FItemList(i).Fsuplycash then %>
          <td align="right">
          		<font color="#FF0000"><b><%= FormatNumber(ojungsan.FItemList(i).FOrgsuplycash,0) %></b></font>
          </td>
      <% else %>
          <td align="right">
          <% if Not IsNULL(ojungsan.FItemList(i).FOrgsuplycash) then %>
    	      <font color="#AAAAAA"><%= FormatNumber(ojungsan.FItemList(i).FOrgsuplycash,0) %></font>
    	  <% end if %>
          </td>
      <% end if %>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).FOrgOptaddprice,0) %></td>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).FOrgOptaddbuyprice,0) %></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsellcash,0) %></font></td>
      <td align="center">
        <% if ojungsan.FItemList(i).FOrgsellcash<>0 then %>
        <%= 100-CLng((ojungsan.FItemList(i).Fsellcash-ojungsan.FItemList(i).FOrgOptaddprice)/ojungsan.FItemList(i).FOrgsellcash*100) %>
        <% end if %>
      </td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Freducedprice) %>"><%= FormatNumber(ojungsan.FItemList(i).Freducedprice,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fcommission) %>"><%= FormatNumber(ojungsan.FItemList(i).Fcommission,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).FPgcommission) %>"><%= FormatNumber(ojungsan.FItemList(i).FPgcommission,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="center">
      <% if ojungsan.FItemList(i).Fsellcash<>0 then %>
      	<%= 100-CLng(ojungsan.FItemList(i).Fsuplycash/ojungsan.FItemList(i).Fsellcash*100*100)/100 %>
      <% end if %>
      </td>
      <td align="right"><font color="<%= MinusFont(suplysum) %>"><%= FormatNumber(suplysum,0) %></font></td>
      <!-- td><input type="button" value="수정" onclick="popBatchDetailEdit('<%= id %>','<%= gubun %>','<%= ojungsan.FItemList(i).FItemID %>','<%= ojungsan.FItemList(i).FItemOption %>','<%= ojungsan.FItemList(i).Fsellcash %>','<%= ojungsan.FItemList(i).Fsuplycash %>','<%= replace(ojungsan.FItemList(i).Fitemname,"'","||39||") %>','<%= replace(ojungsan.FItemList(i).Fitemoptionname,"'","||39||") %>');"></td -->
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td colspan="9"></td>

      <td align="right"><%= FormatNumber(selltotalsum,0) %></td>
      <td colspan="5"></td>
      <td align="right"><b><%= FormatNumber(suplytotalsum,0) %></b></td>
      <!-- td></td -->
    </tr>
</table>



<%
ojungsan.FRectOrder= rectorder
ojungsan.JungsanDetailList
%>
<br>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	<img src="/images/icon_arrow_down.gif" align="absbottom">
			<% if gubun="upche" then %>
			<font color="red"><strong>업체배송 정산메모</strong></font>(정산 차액및 추가 삭제한 내역에 대한 설명을 입력하세요)
			<% elseif gubun="maeip" then %>
			<font color="red"><strong>매입 정산메모</strong></font>(정산 차액및 추가 삭제한 내역에 대한 설명을 입력하세요)
			<% elseif gubun="witakchulgo" then %>
			<font color="red"><strong>위탁 출고 정산메모</strong></font>(정산 차액및 추가 삭제한 내역에 대한 설명을 입력하세요)
			<% elseif gubun="witakoffshop" then %>
			<font color="red"><strong>위탁 오프라인 판매 정산메모</strong></font>(정산 차액및 추가 삭제한 내역에 대한 설명을 입력하세요)
			<% elseif gubun="witaksell" then %>
			<font color="red"><strong>위탁 판매 정산메모</strong></font>(정산 차액및 추가 삭제한 내역에 대한 설명을 입력하세요)
			<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->



<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="memofrm" method="post" action="dodesignerjungsan.asp">
	<input type="hidden" name="idx" value="<%= ojungsanmaster.FItemList(0).FID %>">
	<input type="hidden" name="gubun" value="<%= gubun %>">
	<input type="hidden" name="mode" value="memoedit">
	<tr align="center" bgcolor="#FFFFFF">
		<td>
			<textarea name="tx_memo" cols="90" rows="2"><%= ojungsanmaster.FItemList(0).Fub_comment %></textarea>
			<input type="button" value="메모저장" onclick="savememo(memofrm)">
		</td>
	</tr>
	</form>
</table>

<br>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	<img src="/images/icon_arrow_down.gif" align="absbottom">
			<% if gubun="upche" then %>
			<font color="red"><strong>업체배송내역</strong>(최대 5,000건표시)</font>
			<% elseif gubun="maeip" then %>
			<font color="red"><strong>매입 입고내역</strong></font>
			<% elseif gubun="witakchulgo" then %>
			<font color="red"><strong>위탁 출고내역</strong></font>(정산에 포함됨)
			<% elseif gubun="witaksell" then %>
			<font color="red"><strong>위탁 판매내역</strong></font>(정산에 포함됨 : 최대 5,000건표시)
			<% elseif gubun="witakoffshop" then %>
			<font color="red"><strong>위탁 오프라인 판매내역</strong></font>(정산에 포함됨)
			<% end if %>

			<select name="rectorder" onchange="reOrder(this)">
				<option value="orderserial" <% if rectorder="orderserial" then response.write "selected" %> >주문번호순
				<option value="itemid" <% if rectorder="itemid" then response.write "selected" %> >아이템순
			</select>
        </td>
        <td align="right">
        	<input type="button" value="기타내역추가" onclick="addEtcList(<%= ojungsanmaster.FItemList(0).FID %>,'<%= gubun %>')">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<% if (gubun="maeip") then %>
<td width="80">입출코드</td>
<% else %>
<td width="80">주문번호</td>
<% end if %>
<td width="50">판매채널</td>
<td width="50">구매자</td>
<td width="50">수령인</td>
<td width="120">아이템명</td>
<td width="80">옵션명</td>
<td width="40">수량</td>
<td width="50">판매가</td>
<td width="50">실매출</td>
<td width="50">수수료</td>
<td width="50">PG수수료</td>
<td width="50">공급가</td>
<td width="30">마진</td>
<td width="30">삭제</td>
<td width="30">수정</td>
</tr>
<% for i=0 to ojungsan.FResultCount-1 %>
<form name="frmBuyPrcSell_<%= i %>" method="post" action="dodesignerjungsan.asp">
<input type="hidden" name="idx" value="<%= ojungsan.FItemList(i).Fid %>">
<input type="hidden" name="midx" value="<%= id %>">
<input type="hidden" name="gubun" value="<%= gubun %>">
<input type="hidden" name="mode" value="">
<tr bgcolor="#FFFFFF">
<td ><%= ojungsan.FItemList(i).Fmastercode %></td>
<td ><%= ojungsan.FItemList(i).FSitename %></td>
<td ><%= ojungsan.FItemList(i).FBuyname %></td>
<td ><%= ojungsan.FItemList(i).FReqname %></td>
<td ><%= ojungsan.FItemList(i).FItemName %></td>
<td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
<td ><input type="text" size="3" name="itemno" value="<%= ojungsan.FItemList(i).FItemNo %>" style="text-align:center"></td>
<td ><input type="text" size="5" name="sellcash" value="<%= ojungsan.FItemList(i).Fsellcash %>" style="text-align:right"></td>
<td ><input type="text" size="5" name="reducedprice" value="<%= ojungsan.FItemList(i).Freducedprice %>" <%=CHKIIF(NOT IsCommissionTax,"readonly class='text_ro'","")%> style="text-align:right"></td>
<td ><input type="text" size="5" name="commission" value="<%= ojungsan.FItemList(i).Fcommission %>" <%=CHKIIF(TRUE or NOT IsCommissionTax,"readonly class='text_ro'","")%> style="text-align:right"></td>
<td ><input type="text" size="5" name="pgcommission" value="<%= ojungsan.FItemList(i).Fpgcommission %>" <%=CHKIIF(TRUE or NOT IsCommissionTax,"readonly class='text_ro'","")%> style="text-align:right"></td>
<td ><input type="text" size="5" name="suplycash" value="<%= ojungsan.FItemList(i).Fsuplycash %>" style="text-align:right"></td>
<td >
<%if ojungsan.FItemList(i).Fsellcash<>0 then %>
<%= 100-ojungsan.FItemList(i).Fsuplycash/ojungsan.FItemList(i).Fsellcash*100 %>
<% end if %>
</td>
<td ><a href="javascript:DelDetail(frmBuyPrcSell_<%= i %>)">삭제</a></td>
<td ><a href="javascript:ModiDetail(frmBuyPrcSell_<%= i %>)">수정</a></td>
</tr>
</form>
<%
'' 버퍼구성제한 초과시 아래 주석제거 
if (i mod 1000)=0 then 
    response.flush
end if 
%>
<% next %>
</table>

<br>
<% end if %>

<%
set ojungsan = Nothing
set ojungsanmaster = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->