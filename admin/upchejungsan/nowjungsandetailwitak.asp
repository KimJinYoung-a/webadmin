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
dim witakstep,yyyymm
id = request("id")
gubun = request("gubun")
rectorder = request("rectorder")
witakstep = request("witakstep")

if witakstep="" then witakstep="1"
if rectorder="" then rectorder="itemid"

if gubun="" then gubun="upche"

if (gubun="upche") or (gubun="maeip") then
	witakstep="1"
end if

dim ojungsan, ojungsanmaster
set ojungsan = new CUpcheJungsan
ojungsan.FRectid = id
ojungsan.FRectgubun = gubun

if witakstep="1" then
	ojungsan.JungsanDetailListSum
end if

dim i, suplysum, suplytotalsum
suplysum = 0
suplytotalsum = 0

set ojungsanmaster = new CUpcheJungsan
ojungsanmaster.FRectId = id
ojungsanmaster.JungsanMasterList

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
	window.open('lib/popetclistadd.asp?id=' + iid + '&gubun=' + igubun,'popetc','width=700, height=300, location=no,menubar=no,resizable=yes,scrollbars=no,status=no,toolbar=no');
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
	if (frm.gubun[2].checked){
		frm.action="nowjungsandetailwitak.asp"
	}else{
		frm.action="nowjungsandetail.asp"
	}
	frm.submit();
}
</script>
<table width="760" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="id" value="<%= id %>">
	<input type="hidden" name="rectorder" value="<%= rectorder %>">
	<input type="hidden" name="witakstep" value="<%= witakstep %>">
	<tr>
		<td class="a" >
		<input type="radio" name="gubun" value="upche" <% if gubun="upche" then response.write "checked" %> >업체배송
		<input type="radio" name="gubun" value="maeip" <% if gubun="maeip" then response.write "checked" %> >매입
		<input type="radio" name="gubun" value="witak" <% if gubun="witak" then response.write "checked" %> >위탁

		<td class="a" align="right">
			<a href="javascript:ReSearch(frm);"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="760" cellspacing="0" class="a">
<tr>
  <td align="right"><a href="/admin/storage/nowjungsanlist.asp?menupos=130">목록</a></td>
</tr>
</table>
<% if gubun="witak" then %>
<table width="760" class="a" >
<tr>
	<td></td>
	<td width="120" align="right">
		<% if witakstep="1" then %>
		<a href="?menupos=<%= menupos %>&id=<%= id %>&rectorder=<%= rectorder %>&gubun=<%= gubun %>&witakstep=1"><b>A.위탁입출고내역</b></a>
		<% else %>
		<a href="?menupos=<%= menupos %>&id=<%= id %>&rectorder=<%= rectorder %>&gubun=<%= gubun %>&witakstep=1">A.위탁입출고내역</a>
		<% end if %>
	</td>
	<td width="120" align="right">
		<% if witakstep="2" then %>
		<a href="?menupos=<%= menupos %>&id=<%= id %>&rectorder=<%= rectorder %>&gubun=<%= gubun %>&witakstep=2"><b>B.위탁정산내역확정</b></a>
		<% else %>
		<a href="?menupos=<%= menupos %>&id=<%= id %>&rectorder=<%= rectorder %>&gubun=<%= gubun %>&witakstep=2">B.위탁정산내역확정</a>
		<% end if %>
	</td>
</tr>
</table>
<% end if %>

<% if witakstep="1" then %>
<div class="a">1.아이템 합계 (내역을 수정하면 합계에 적용됩니다.)</div>
<table width="760" cellspacing="1"  class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
      <td width="40">상품ID</td>
      <td width="200">상품명</td>
      <td width="80">옵션명</td>
      <td width="40">갯수</td>
      <td width="50"><font color="#AAAAAA">판매가 (현재)</font></td>
      <td width="50"><font color="#AAAAAA">공급가 (현재)</font></td>
      <td width="80">판매가</td>
      <td width="80">공급가</td>
      <td width="40">마진</td>
      <td width="80">공급가합계</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    suplysum =0
    suplysum = suplysum + ojungsan.FItemList(i).Fsuplycash * ojungsan.FItemList(i).FItemNo
    suplytotalsum = suplytotalsum + suplysum

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
      <% if ojungsan.FItemList(i).FOrgsellcash<>ojungsan.FItemList(i).Fsellcash then %>
      <td align="right"><font color="#FF0000"><b><%= FormatNumber(ojungsan.FItemList(i).FOrgsellcash,0) %></b></font></td>
      <% else %>
      <td align="right"><font color="#AAAAAA"><%= FormatNumber(ojungsan.FItemList(i).FOrgsellcash,0) %></font></td>
      <% end if %>

      <% if ojungsan.FItemList(i).FOrgsuplycash<>ojungsan.FItemList(i).Fsuplycash then %>
      <td align="right"><font color="#FF0000"><b><%= FormatNumber(ojungsan.FItemList(i).FOrgsuplycash,0) %></b></font></td>
      <% else %>
      <td align="right"><font color="#AAAAAA"><%= FormatNumber(ojungsan.FItemList(i).FOrgsuplycash,0) %></font></td>
      <% end if %>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="center"><%= 100-CLng(ojungsan.FItemList(i).Fsuplycash/ojungsan.FItemList(i).Fsellcash*100) %></td>
      <td align="right"><font color="<%= MinusFont(suplysum) %>"><%= FormatNumber(suplysum,0) %></font></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td colspan="9"></td>
      <td align="right"><%= FormatNumber(suplytotalsum,0) %></td>
    </tr>
</table>
<% end if %>
<%
ojungsan.FRectOrder= rectorder

if witakstep="1" then
ojungsan.JungsanDetailList
%>
<br>
<div class="a">
<% if gubun="upche" then %>
2.업체배송 정산메모(정산 차액및 추가 삭제한 내역에 대한 설명을 입력하세요)
<% elseif gubun="maeip" then %>
2.매입 정산메모(정산 차액및 추가 삭제한 내역에 대한 설명을 입력하세요)
<% else %>
2.위탁 정산메모(정산 차액및 추가 삭제한 내역에 대한 설명을 입력하세요)
<% end if %>
<table width="760" cellspacing="1"  class="a" bgcolor=#3d3d3d>
<form name="memofrm" method="post" action="dodesignerjungsan.asp">
<input type="hidden" name="idx" value="<%= ojungsanmaster.FItemList(0).FID %>">
<input type="hidden" name="gubun" value="<%= gubun %>">
<input type="hidden" name="mode" value="memoedit">
<tr bgcolor="#FFFFFF">
	<td>
		<textarea name="tx_memo" cols="90" rows="7"><%= ojungsanmaster.FItemList(0).Fub_comment %></textarea>
		<input type="button" value="메모저장" onclick="savememo(memofrm)">
	</td>
</tr>
</form>
</table>
</div>
<br>
<table width="760" class="a" border="0">
<tr>
	<td>
	<% if gubun="upche" then %>
	3.업체배송내역
	<% elseif gubun="maeip" then %>
	3.매입 입고내역
	<% else %>
	3.위탁 입고내역
	<% end if %>
	<select name="rectorder" onchange="reOrder(this)">
	<option value="orderserial" <% if rectorder="orderserial" then response.write "selected" %> >주문번호순
	<option value="itemid" <% if rectorder="itemid" then response.write "selected" %> >아이템순
	</select>
	</td>
	<td align="right"><input type="button" value="기타내역추가" onclick="addEtcList(<%= ojungsanmaster.FItemList(0).FID %>,'<%= gubun %>')"></td>
</tr>
</div>
<table width="760" cellspacing="1"  class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
      <% if gubun="upche" then %>
      <td width="80">주문번호</td>
      <% elseif gubun="maeip" then %>
      <td width="80">매입코드</td>
      <% else %>
      <td width="80">위탁코드</td>
      <% end if %>
      <td width="50">구매자</td>
      <td width="50">수령인</td>
      <td width="120">아이템명</td>
      <td width="80">옵션명</td>
      <td width="40">갯수</td>
      <td width="50">판매가</td>
      <td width="50">공급가</td>
      <td width="30">삭제</td>
      <td width="30">수정</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <form name="frmBuyPrc_<%= i %>" method="post" action="dodesignerjungsan.asp">
    <input type="hidden" name="idx" value="<%= ojungsan.FItemList(i).Fid %>">
    <input type="hidden" name="midx" value="<%= id %>">
    <input type="hidden" name="gubun" value="<%= gubun %>">
    <input type="hidden" name="mode" value="">
    <tr bgcolor="#FFFFFF">
      <td ><%= ojungsan.FItemList(i).Fmastercode %></td>
      <td ><%= ojungsan.FItemList(i).FBuyname %></td>
      <td ><%= ojungsan.FItemList(i).FReqname %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td ><input type="text" size="3" name="itemno" value="<%= ojungsan.FItemList(i).FItemNo %>"></td>
      <td ><input type="text" size="5" name="sellcash" value="<%= ojungsan.FItemList(i).Fsellcash %>"></td>
      <td ><input type="text" size="5" name="suplycash" value="<%= ojungsan.FItemList(i).Fsuplycash %>"></td>
      <td ><a href="javascript:DelDetail(frmBuyPrc_<%= i %>)">삭제</a></td>
      <td ><a href="javascript:ModiDetail(frmBuyPrc_<%= i %>)">수정</a></td>
    </tr>
    </form>
    <% next %>
</table>
<br>

<% end if %>

<% if gubun="witak" and witakstep="1" then %>
<%
ojungsan.FRectgubun = "witakchulgo"
ojungsan.JungsanDetailList
%>
<table width="760" border="0">
<tr>
<td align="right"><input type="button" value="기타내역추가" onclick="addEtcList(<%= ojungsanmaster.FItemList(0).FID %>,'witakchulgo')"></td>
</tr>
</table>
<table width="760" cellspacing="1"  class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
      <td width="80">출고코드</td>
      <td width="50">구매자</td>
      <td width="50">수령인</td>
      <td width="120">아이템명</td>
      <td width="80">옵션명</td>
      <td width="40">갯수</td>
      <td width="50">판매가</td>
      <td width="50">공급가</td>
      <td width="30">삭제</td>
      <td width="30">수정</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <form name="frmBuyPrc1_<%= i %>" method="post" action="dodesignerjungsan.asp">
    <input type="hidden" name="midx" value="<%= id %>">
    <input type="hidden" name="idx" value="<%= ojungsan.FItemList(i).Fid %>">
    <input type="hidden" name="gubun" value="<%= gubun %>">
    <input type="hidden" name="mode" value="">
    <tr bgcolor="#FFFFFF">
      <td ><%= ojungsan.FItemList(i).Fmastercode %></td>
      <td ><%= ojungsan.FItemList(i).FBuyname %></td>
      <td ><%= ojungsan.FItemList(i).FReqname %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td ><input type="text" size="3" name="itemno" value="<%= ojungsan.FItemList(i).FItemNo %>"></td>
      <td ><input type="text" size="5" name="sellcash" value="<%= ojungsan.FItemList(i).Fsellcash %>"></td>
      <td ><input type="text" size="5" name="suplycash" value="<%= ojungsan.FItemList(i).Fsuplycash %>"></td>
      <td ><a href="javascript:DelDetail(frmBuyPrc1_<%= i %>)">삭제</a></td>
      <td ><a href="javascript:ModiDetail(frmBuyPrc1_<%= i %>)">수정</a></td>
    </tr>
    </form>
    <% next %>
</table>
<% end if %>

<% if gubun="witak" and witakstep<>"1" then %>
<%
ojungsan.FRectid = id
ojungsan.FrectDesigner = ojungsanmaster.FItemList(0).FDesignerid
ojungsan.FRectStartDay = yyyymm + "-" + "01"
ojungsan.FRectEndDay   = CStr(DateSerial(Left(yyyymm,4), CLng(Right(yyyymm,2))+1,1))
ojungsan.FRectYYYYMM   = yyyymm
ojungsan.FRectPreYYYYMM   = Left(CStr(DateSerial(Left(yyyymm,4), CLng(Right(yyyymm,2))-1,1)),7)

'response.write ojungsan.FRectStartDay
'response.write ojungsan.FRectEndDay
ojungsan.GetWitakJungSanByItem


dim i_pjaego, i_rjaego
dim sysjaego, ocha
dim bufipgo, bufchulgo
dim totjungsanno, totjungsansum
%>
<script language='javascript'>
function saveArr(){
	var frm;
	var upfrm = document.frmarr;

	var ret = confirm('저장하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				upfrm.detailidx.value = upfrm.detailidx.value + frm.detailidx.value + "|";
				upfrm.itemid.value = upfrm.itemid.value + frm.itemid.value + "|";
				upfrm.itemoption.value = upfrm.itemoption.value + frm.itemoption.value + "|";
				upfrm.sellcash.value = upfrm.sellcash.value + frm.sellcash.value + "|";
				upfrm.suplycash.value = upfrm.suplycash.value + frm.suplycash.value + "|";
				upfrm.prejaego.value = upfrm.prejaego.value + frm.prejaego.value + "|";
				upfrm.ipgono.value = upfrm.ipgono.value + frm.ipgono.value + "|";
				upfrm.chulgono.value = upfrm.chulgono.value + frm.chulgono.value + "|";
				upfrm.sellno.value = upfrm.sellno.value + frm.sellno.value + "|";
				upfrm.ocha.value = upfrm.ocha.value + frm.ocha.value + "|";
				upfrm.realjaego.value = upfrm.realjaego.value + frm.realjaego.value + "|";
				upfrm.jungsanno.value = upfrm.jungsanno.value + frm.jungsanno.value + "|";
				if (frm.isdelete.checked){
					upfrm.isdelete.value = upfrm.isdelete.value + "Y" + "|";
				}else{
					upfrm.isdelete.value = upfrm.isdelete.value + "N" + "|";
				}

			}
		}

		upfrm.submit();
	}
}

function delArr(){
	var upfrm = document.frmarr;
	var ret = confirm('확정된 데이터를 삭제 하시겠습니까?');
	if (ret){
		upfrm.gubun.value = "witakjungsan_del";
		upfrm.submit();
	}
}
</script>

<table width="1000" cellspacing="1"  class="a" >
<tr>
	<td width="140"><%= yyyymm %> 위탁정산내역</td>
	<% if ojungsan.FWitakInsserted then %>
	<td width="200"><font color="red"> 현 데이터는 확정된 데이타 입니다.</font></td>
	<td><!-- <b><a href="javascript:delArr()">[삭제]</a></b> --></td>
	<% end if %>
	<td align="right"><input type="button" value="내역확정" onclick="saveArr()"></td>
</tr>
</table>

<table width="1000" cellspacing="1"  class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
    	<td width="50">상품번호</td>
    	<td width="120">상품명</td>
    	<td width="80">옵션</td>
    	<td width="50">소비자가(현재)</td>
    	<td width="50">공급가(현재)</td>
    	<td width="50">소비자가(판매)</td>
    	<td width="50">공급가(판매)</td>
    	<td width="30"></td>
    	<td width="50">이월재고 (A)</td>
    	<td width="50">입고수량 (B)</td>
    	<td width="50">출고수량 (C)</td>
    	<td width="50">판매수량 (D)</td>
    	<td width="60">시스템재고 (S=A+B-C-D)</td>
    	<td width="50">오차<br>(E=S-R)</td>
    	<td width="50">실사재고 (R)</td>
    	<td width="50">정산예정수량(C+D+E)</td>
    	<td width="50">정산수량</td>
    	<td width="20">삭제</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    	duplicated = ojungsan.CheckDuplicated(i)
   	%>

    <% if (ojungsan.FItemList(i).FIsUsing<>"Y") and (ojungsan.FItemList(i).FIpGoNo=0) and (ojungsan.FItemList(i).FChulgoNo=0) and (ojungsan.FItemList(i).FsellNo=0) then %>
    <% else %>
        <form name="frmBuyPrc_<%= i %>" method="post" action="">
        <input type="hidden" name="detailidx" value="<%= ojungsan.FItemList(i).Fdetailidx %>">
		<input type="hidden" name="itemid" value="<%= ojungsan.FItemList(i).Fitemid %>">
		<input type="hidden" name="itemoption" value="<%= ojungsan.FItemList(i).Fitemoption %>">
    	<% if duplicated then %>
    		<tr bgcolor="#EEEEEE">
    	<% else %>
		    <% if ojungsan.FItemList(i).FIsUsing<>"Y" then %>
		    <tr bgcolor="#FFFFFF" class="gray">
		    <% else %>
		    <tr bgcolor="#FFFFFF">
		    <% end if %>
		<% end if %>
	    	<td><%= ojungsan.FItemList(i).Fitemid %></td>
	    	<td><%= ojungsan.FItemList(i).Fitemname %></td>
	    	<td><%= ojungsan.FItemList(i).Fitemoptionname %></td>
	    	<td align="right"><%= ojungsan.FItemList(i).FSellcash %></td>
	    	<td align="right"><%= ojungsan.FItemList(i).FSuplycash %></td>
	    	<td align="right"><input type="text" name="sellcash" value="<%= ojungsan.FItemList(i).FSellcash_sell %>" size="6" style="border-width:1; border-color:#AAAAAA; border-style:solid;" ></td>
	    	<td align="right"><input type="text" name="suplycash" value="<%= ojungsan.FItemList(i).FSuplycash_sell %>" size="6" style="border-width:1; border-color:#AAAAAA; border-style:solid;" ></td>
	    	<td align="center"><%= ojungsan.FItemList(i).FPrejaego %></td>

	    	<td align="center"><input type="text" name="prejaego" size=3 value="<%= ojungsan.FItemList(i).FPrejaego %>" style="border-width:1; border-color:#AAAAAA; border-style:solid;" onKeyUp="javascript:ReCalcu(frmBuyPrc_<%= i %>)"></td>
	    	<td align="center"><input type="text" name="ipgono" size=3 value="<%= ojungsan.FItemList(i).FIpgoNo %>" style="border-width:1; border-color:#AAAAAA; border-style:solid;" onKeyUp="javascript:ReCalcu(frmBuyPrc_<%= i %>)"></td>
	    	<td align="center"><input type="text" name="chulgono" size=3 value="<%= ojungsan.FItemList(i).FChulgono %>" style="border-width:1; border-color:#AAAAAA; border-style:solid;" onKeyUp="javascript:ReCalcu(frmBuyPrc_<%= i %>)"></td>
	    	<td align="center"><input type="text" name="sellno" value="<%= ojungsan.FItemList(i).FsellNo %>" size="4" style="border-width:1; border-color:#AAAAAA; border-style:solid;" onKeyUp="javascript:ReCalcu(frmBuyPrc_<%= i %>)"></td>
	    	<td align="center"><input type="text" name="tmpsysjaego" size="3" value="<%= ojungsan.FItemList(i).FsysJaeGo %>" style="border-width:1; border-color:#FFFFFF; border-style:solid; " readonly ></td>
	    	<% if ojungsan.FItemList(i).FOCha<>0 then %>
	    	<td align="center"><input type="text" name="ocha" size="3" value="<%= ojungsan.FItemList(i).FOCha %>" style="border-width:1; border-color:#FFFFFF; border-style:solid; color:#FF0000"></td>
	    	<% else %>
	    	<td align="center"><input type="text" name="ocha" size="3" value="<%= ojungsan.FItemList(i).FOCha %>" style="border-width:1; border-color:#FFFFFF; border-style:solid;"></td>
	    	<% end if %>
	    	<td align="center"><input type="text" name="realjaego" size=3 value="<%= ojungsan.FItemList(i).FRealJaego %>" style="border-width:1; border-color:#AAAAAA; border-style:solid;" onKeyUp="javascript:ReCalcu(frmBuyPrc_<%= i %>)"></td>
	    	<td align="center"><%= ojungsan.FItemList(i).FjungsanNo %></td>
	    	<td align="center"><input type="text" name="jungsanno" size="3" value="<%= ojungsan.FItemList(i).FjungsanNo %>" style="border-width:1; border-color:#AAAAAA; border-style:solid;"></td>
	    	<td align="center"><input type="checkbox" name="isdelete" <% if ojungsan.FItemList(i).FIsDelete="Y" then response.write "checked" %> ></td>
	    	<%
	    		if ojungsan.FItemList(i).FIsDelete<>"Y" then
		    		totjungsanno = totjungsanno + ojungsan.FItemList(i).FjungsanNo
		    		totjungsansum = totjungsansum + ojungsan.FItemList(i).FSuplycash_sell * ojungsan.FItemList(i).FjungsanNo
	    		end if
	    	%>
	    </tr>
	    </form>
	<% end if %>
    <% next %>
    <tr bgcolor="#FFFFFF">
    	<td width="50">총계</td>
    	<td align="right" colspan="17">총 건수 : <%= totjungsanno %> 총 금액 : <%= FormatNumber(totjungsansum,0) %></td>
    </tr>
</table>
<form name="frmarr" method="post" action="dodesignerjungsan.asp">
<input type="hidden" name="mode" value="arrsave">
<input type="hidden" name="gubun" value="witakjungsan">
<input type="hidden" name="idx" value="<%= id %>">
<input type="hidden" name="detailidx" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemoption" value="">
<input type="hidden" name="sellcash" value="">
<input type="hidden" name="suplycash" value="">
<input type="hidden" name="prejaego" value="">
<input type="hidden" name="ipgono" value="">
<input type="hidden" name="chulgono" value="">
<input type="hidden" name="sellno" value="">
<input type="hidden" name="ocha" value="">
<input type="hidden" name="realjaego" value="">
<input type="hidden" name="jungsanno" value="">
<input type="hidden" name="isdelete" value="">

</form>
<% end if %>
<%
set ojungsan = Nothing
set ojungsanmaster = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->