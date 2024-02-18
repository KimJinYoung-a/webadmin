<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim designer, gubun
dim page
dim yyyy1,mm1
dim differencekey

designer    = reQuestCheckVar(request("designer"),32)
gubun       = reQuestCheckVar(request("gubun"),16)
page        = reQuestCheckVar(request("page"),9)
yyyy1       = reQuestCheckVar(request("yyyy1"),4)
mm1             = reQuestCheckVar(request("mm1"),2)
differencekey   = reQuestCheckVar(request("differencekey"),9)

if gubun="" then gubun="upche"
if page="" then page=1

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

dim ojungsan, ojungsanselected
set ojungsan = new CUpcheJungsan
ojungsan.FPageSize = 3000
ojungsan.FCurrPage = page
ojungsan.FRectGubun = gubun
ojungsan.FRectDesigner = designer

''강좌만해당
ojungsan.FRectYYYYMM = yyyy1 + "-" + mm1

if (designer<>"") then
    ojungsan.SearchJungsanList
end if

set ojungsanselected = new CUpcheJungsan


dim i,j,ischeckd
dim checkdate1,checkdate2
checkdate1 = dateserial(yyyy1,mm1+1,1)
checkdate2 = dateserial(yyyy1,mm1,1)

dim iitemlist, precode
%>

<script language='javascript'>
function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function SaveArr(){
	var differencekey = document.frm.differencekey.value;
	var taxtype ="";



	if (differencekey.length<1){
		alert('차수를 입력하세요. 수시정산-1부터시작, 일반-0부터시작');
		document.frm.differencekey.focus();
		return;
	}

	if (document.frm.taxtype[0].checked){
		taxtype = "01";
	}else if (document.frm.taxtype[1].checked){
		taxtype = "02";
	}else if (document.frm.taxtype[2].checked){
		taxtype = "03";
	}else{
		alert('과세구분을 선택 하세요.');
		document.frm.taxtype[0].focus();
		return;
	}

	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		ret = confirm('선택 내역이 없습니다. \r\n\r\n 정산대상 내역으로 저장 하시겠습니까?');
		if (!ret){
			return;
		}else{

		}
	}else{
		ret = confirm('선택 내역을 정산대상 내역으로 저장 하시겠습니까?');
	}

	upfrm.idx.value = "";

	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.idx.value = upfrm.idx.value + frm.idx.value + ",";
				}
			}
		}
		upfrm.differencekey.value=differencekey;
		upfrm.taxtype.value=taxtype;
		upfrm.mode.value="arrsave";
		upfrm.submit();
	}
}



function popOrderDetailEdit(idx){
	var popwin = window.open('/common/orderdetailedit_UTF8.asp?idx=' + idx,'orderdetailedit','width=600,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>


<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	정산대상년월:<% DrawYMBox yyyy1,mm1 %>
				차수 : <input type=text name="differencekey" value="<%= differencekey %>" size=2>
				<input type="radio" name="taxtype" value="01">과세
				<input type="radio" name="taxtype" value="02">면세
				<input type="radio" name="taxtype" value="03">원천징수등
				<br>
				업체:<% drawSelectBoxDesignerwithName "designer",designer  %>
				<input type="radio" name="gubun" value="upche" <% if gubun="upche" then response.write "checked" %> >업체배송
				<input type="radio" name="gubun" value="maeip" <% if Left(gubun,5)="maeip" then response.write "checked" %> >매입
				<input type="radio" name="gubun" value="witak" <% if Left(gubun,5)="witak" then response.write "checked" %> >위탁
		
				<input type="radio" name="gubun" value="lecture" <% if gubun="lecture" then response.write "checked" %> >강좌
	        </td>
	        <td valign="top" align="right">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->


<% if (gubun="upche") then %>
<!-- 업체배송 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="right">
        	<input type="button" value="정산대상내역으로 저장" onclick="SaveArr()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>


<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
    <tr align="center" bgcolor="#DDDDFF">
      <td width="20" ><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
      <td width="80">주문번호</td>
      <td width="60">구매자</td>
      <td>상품명</td>
      <td>옵션</td>
      <td width="40">(현재)<br>과세</td>
      <td width="30">수량</td>
      <td width="60">판매가</td>
      <td width="60">매입가</td>
      <td width="70">주문일</td>
      <td width="70">입금일</td>
      <td width="70">출고일</td>
      <td width="70">진행상태</td>
      <td width="80">송장번호</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    	ischeckd = false
    	if Not IsNull(ojungsan.FItemList(i).FBeasongdate) then
    		ischeckd = (ojungsan.FItemList(i).FCurrState="7") and (Cdate(ojungsan.FItemList(i).FBeasongdate)<checkdate1) and (Cdate(ojungsan.FItemList(i).FBeasongdate)>=checkdate2)
    	end if

    	ischeckd = ischeckd or ((ojungsan.FItemList(i).FJumundiv="9") and (Cdate(ojungsan.FItemList(i).FRegDate)<checkdate1) and (Cdate(ojungsan.FItemList(i).FRegDate)>=checkdate2))
    %>
    <form name="frmBuyPrc_<%= i %>" >
    <input type="hidden" name="idx" value="<%= ojungsan.FItemList(i).FIDX %>">
    <tr bgcolor="#FFFFFF" <% if ischeckd then response.write "class='H'" %> >
      <td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <% if ischeckd then response.write "checked" %> ></td>
      <td ><a href="javascript:popOrderDetailEdit(<%= ojungsan.FItemList(i).FIDX %>);"><%= ojungsan.FItemList(i).Forderserial %></a></td>
      <td align="center"><%= ojungsan.FItemList(i).FBuyname %></td>
      <td><%= ojungsan.FItemList(i).FItemName %></td>
      <td><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= ChkIIF(ojungsan.FItemList(i).Fvatinclude="N","면세","") %></td>
      <td align="center"><%= ojungsan.FItemList(i).FItemNo %></td>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).FSellCash,0) %></td>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).FBuyCash,0) %></td>
      <td align="center"><acronym title="<%= ojungsan.FItemList(i).FRegDate %>"><%= Left(ojungsan.FItemList(i).FRegDate,10) %></acronym></td>
      <td align="center"><acronym title="<%= ojungsan.FItemList(i).FIpkumDate %>"><%= Left(ojungsan.FItemList(i).FIpkumDate,10) %></acronym></td>
      <td align="center"><acronym title="<%= ojungsan.FItemList(i).FBeasongdate %>"><%= Left(ojungsan.FItemList(i).FBeasongdate,10) %></acronym></td>
      <td align="center">
      <% if ojungsan.FItemList(i).FJumundiv="9" then %>
      <font color="#FF33FF"><b>마이너스</b></font>
      <% else %>
      <font color="<%= UpCheBeasongStateColor(ojungsan.FItemList(i).FCurrState) %>"><%= UpCheBeasongState2Name(ojungsan.FItemList(i).FCurrState) %></font>
      <% end if %>
      </td>
      <td align="center"><%= ojungsan.FItemList(i).FUpcheSongjangNo %></td>
    </tr>
    </form>
    <% next %>
</table>
<%
ojungsanselected.FRectdesigner = designer
ojungsanselected.FRectYYYYMM = yyyy1 + "-" + mm1
ojungsanselected.FRectgubun = "upche"
ojungsanselected.FRectdifferencekey = differencekey

if (designer<>"") then
    ojungsanselected.JungsanDetailListByYYYYMM
end if
%>

<table width="100%" cellspacing="1" class="a" >
<tr><td><hr></td></tr>
</table>

<div class="a"><b>*등록된 업체배송 정산 대상 내역</b></div>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="90">주문번호</td>
      <td width="80">구매자</td>
      <td width="80">수령인</td>
      <td>아이템명</td>
      <td>옵션명</td>
      <td width="40">과세</td>
      <td width="40">수량</td>
      <td width="60">판매가</td>
      <td width="60">매입가</td>
      <td width="70">입금확인일</td>
      <td width="70">출고일</td>
      <td width="50">삭제</td>
    </tr>
    <% for i=0 to ojungsanselected.FResultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
      <td ><%= ojungsanselected.FItemList(i).Fmastercode %></td>
      <td ><%= ojungsanselected.FItemList(i).FBuyname %></td>
      <td ><%= ojungsanselected.FItemList(i).FReqname %></td>
      <td align="left"><%= ojungsanselected.FItemList(i).FItemName %></td>
      <td><%= ojungsanselected.FItemList(i).FItemOptionName %></td>
      <td><%= ChkIIF(ojungsanselected.FItemList(i).Fvatinclude="N","면세","") %></td>
      <td ><%= ojungsanselected.FItemList(i).FItemNo %></td>
      <td align="right"><%= FormatNumber(ojungsanselected.FItemList(i).Fsellcash,0) %></td>
      <td align="right"><%= FormatNumber(ojungsanselected.FItemList(i).Fsuplycash,0) %></td>
      <td align="center"></td>
      <td align="center"></td>
      <td >삭제</td>
    </tr>
    <% next %>
</table>


<% elseif Left(gubun,5)="maeip" then %>
<!-- 매입 -->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <% if gubun="maeip" then %>
		<td width="100"><a href="?designer=<%= designer %>&gubun=maeip&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>"><b>1.매입입고내역</b></a></td>
		<% else %>
		<td width="100"><a href="?designer=<%= designer %>&gubun=maeip&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>">1.매입입고내역</a></td>
		<% end if %>
		<% if gubun="maeipchulgo" then %>
		<td width="100"><a href="?designer=<%= designer %>&gubun=maeipchulgo&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>"><b>2.매입출고내역</b></a></td>
		<% else %>
		<td width="100"><a href="?designer=<%= designer %>&gubun=maeipchulgo&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>">2.매입출고내역</a></td>
		<% end if %>
		<td align="right"><input type="button" value="정산대상내역으로저장" onclick="SaveArr()"></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->


<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
    <tr align="center" bgcolor="#DDDDFF">
      <td width="20" ><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
      <td colspan="3">입고코드</td>
      <td width="40">(현재)<br>과세</td>
      <td width="80">총매입가</td>
      <td width="50">수량</td>
      <% if gubun="maeip" then %>
      <td width="80">입고일</td>
      <% else %>
      <td width="80">출고(예정)일</td>
      <% end if %>
      <td width="150">등록일</td>
      <td width="100"></td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    	ischeckd = false
    	if Not IsNull(ojungsan.FItemList(i).FExecuteDate) then
    		ischeckd = ((Cdate(ojungsan.FItemList(i).FExecuteDate)<checkdate1) and (Cdate(ojungsan.FItemList(i).FExecuteDate)>=checkdate2))
    	end if
    %>
    <% if precode<>ojungsan.FItemList(i).FCode then %>
    <tr bgcolor="#FFFFFF">
      <td></td>
      <td colspan="3"><b><%= ojungsan.FItemList(i).FCode %></b></td>
      <td></td>
      <td align="right"><b><%= FormatNumber(ojungsan.FItemList(i).FTotalsuplycash,0) %></b></td>
      <td></td>
      <% if gubun="maeip" then %>
      <td><%= ojungsan.FItemList(i).FExecuteDate %></td>
      <% else %>
      <td><%= ojungsan.FItemList(i).FScheduleDate %></td>
      <% end if %>
      <td><%= ojungsan.FItemList(i).FRegDate %></td>
      <% if gubun="maeip" then %>
      <td ></td>
      <% else %>
      <td ><%= ojungsan.FItemList(i).FDesignerID %></td>
      <% end if %>
    </tr>
    <% end if %>
    <% precode = ojungsan.FItemList(i).FCode %>
    	<form name="frmBuyPrc_<%= i %>" >
    	<input type="hidden" name="idx" value="<%= ojungsan.FItemList(i).FID %>">
    	<tr bgcolor="#FFFFFF" <% if ischeckd then response.write "class='H'" %> >
	      <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <% if ischeckd then response.write "checked" %> ></td>
	      <td width="60"></td>
	      <td><%= ojungsan.FItemList(i).FItemName %></td>
	      <td width="60"><%= ojungsan.FItemList(i).FItemOptionName %></td>
	      <td align="center"><%= ChkIIF(ojungsan.FItemList(i).Fvatinclude="N","면세","") %></td>
	      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).Fsuplycash,0) %></td>
	      <td align="center"><%= ojungsan.FItemList(i).FItemNo %></td>
	      <td></td>
	      <td></td>
	      <td></td>
	    </tr>
	    </form>
    <% next %>
</table>
<%
ojungsanselected.FRectdesigner = designer
ojungsanselected.FRectYYYYMM = yyyy1 + "-" + mm1
ojungsanselected.FRectgubun = "maeip"
ojungsanselected.FRectdifferencekey = differencekey

if (designer<>"") then
    ojungsanselected.JungsanDetailListByYYYYMM
end if
%>
<br>

<table width="100%" cellspacing="1" class="a" >
<tr><td><hr></td></tr>
</table>

<div class="a"><b>*등록된 매입입고내역</b></div>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
	<tr align="center" bgcolor="#DDDDFF">
      <td width="80">입고코드</td>
      <td width="80">구매자</td>
      <td width="80">수령인</td>
      <td>상품명</td>
      <td width="80">옵션명</td>
      <td width="40">(현재)<br>과세</td>
      <td width="40">수량</td>
      <td width="50">판매가</td>
      <td width="50">매입가</td>
      <td width="50">삭제</td>
    </tr>
    <% for i=0 to ojungsanselected.FResultCount-1 %>
    <tr bgcolor="#FFFFFF">
      <td align="center"><%= ojungsanselected.FItemList(i).Fmastercode %></td>
      <td align="center"><%= ojungsanselected.FItemList(i).FBuyname %></td>
      <td align="center"><%= ojungsanselected.FItemList(i).FReqname %></td>
      <td ><%= ojungsanselected.FItemList(i).FItemName %></td>
      <td ><%= ojungsanselected.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= ChkIIF(ojungsanselected.FItemList(i).Fvatinclude="N","면세","") %></td>
      <td align="center"><%= ojungsanselected.FItemList(i).FItemNo %></td>
      <td align="right"><%= ojungsanselected.FItemList(i).Fsellcash %></td>
      <td align="right"><%= ojungsanselected.FItemList(i).Fsuplycash %></td>
      <td align="center">삭제</td>
    </tr>
    <% next %>
</table>

<%
ojungsanselected.FRectgubun = "maeipchulgo"
ojungsanselected.FRectdifferencekey = differencekey

if (designer<>"") then
    ojungsanselected.JungsanDetailListByYYYYMM
end if
%>

<p>

<div class="a"><b>등록된 매입출고내역</b></div>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
	<tr align="center" bgcolor="#DDDDFF">
      <td width="80">출고코드</td>
      <td width="80">구매자</td>
      <td width="80">수령인</td>
      <td>아이템명</td>
      <td width="80">옵션명</td>
      <td width="40">(현재)<br>과세</td>
      <td width="40">수량</td>
      <td width="50">판매가</td>
      <td width="50">매입가</td>
      <td width="50">삭제</td>
    </tr>
    <% for i=0 to ojungsanselected.FResultCount-1 %>
    <tr bgcolor="#FFFFFF">
      <td align="center"><%= ojungsanselected.FItemList(i).Fmastercode %></td>
      <td align="center"><%= ojungsanselected.FItemList(i).FBuyname %></td>
      <td align="center"><%= ojungsanselected.FItemList(i).FReqname %></td>
      <td ><%= ojungsanselected.FItemList(i).FItemName %></td>
      <td ><%= ojungsanselected.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= ChkIIF(ojungsanselected.FItemList(i).Fvatinclude="N","면세","") %></td>
      <td align="center"><%= ojungsanselected.FItemList(i).FItemNo %></td>
      <td align="right"><%= ojungsanselected.FItemList(i).Fsellcash %></td>
      <td align="right"><%= ojungsanselected.FItemList(i).Fsuplycash %></td>
      <td align="center">삭제</td>
    </tr>
    <% next %>
</table>


<% elseif Left(gubun,5)="witak" then %>
<!-- 위탁 -->

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="right">
    	<% if gubun="witak" then %>
		<td width="100"><a href="?designer=<%= designer %>&gubun=witak&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>"><b>1.위탁입고내역</b></a></td>
		<td width="100"><a href="?designer=<%= designer %>&gubun=witakchulgo&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>">2.위탁출고내역</a></td>
		<td width="100"><a href="?designer=<%= designer %>&gubun=witaksell&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>">3.위탁판매내역</a></td>
		<td width="150"><a href="?designer=<%= designer %>&gubun=witakoffshop&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>">4.위탁오프라인판매내역</a></td>
		<% elseif gubun="witakchulgo" then %>
		<td width="100"><a href="?designer=<%= designer %>&gubun=witak&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>">1.위탁입고내역</a></td>
		<td width="100"><a href="?designer=<%= designer %>&gubun=witakchulgo&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>"><b>2.위탁출고내역</b></a></td>
		<td width="100"><a href="?designer=<%= designer %>&gubun=witaksell&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>">3.위탁판매내역</a></td>
		<td width="150"><a href="?designer=<%= designer %>&gubun=witakoffshop&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>">4.위탁오프라인판매내역</a></td>
		<% elseif gubun="witaksell" then %>
		<td width="100"><a href="?designer=<%= designer %>&gubun=witak&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>">1.위탁입고내역</a></td>
		<td width="100"><a href="?designer=<%= designer %>&gubun=witakchulgo&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>">2.위탁출고내역</a></td>
		<td width="100"><a href="?designer=<%= designer %>&gubun=witaksell&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>"><b>3.위탁판매내역</b></a></td>
		<td width="150"><a href="?designer=<%= designer %>&gubun=witakoffshop&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>">4.위탁오프라인판매내역</a></td>
		<% elseif gubun="witakoffshop" then %>
		<td width="100"><a href="?designer=<%= designer %>&gubun=witak&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>">1.위탁입고내역</a></td>
		<td width="100"><a href="?designer=<%= designer %>&gubun=witakchulgo&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>">2.위탁출고내역</a></td>
		<td width="100"><a href="?designer=<%= designer %>&gubun=witaksell&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>">3.위탁판매내역</b></td>
		<td width="150"><a href="?designer=<%= designer %>&gubun=witakoffshop&page=<%= page %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>"><b>4.위탁오프라인판매내역</b></a></td>
		<% end if %>
		<td align="right"><input type="button" value="정산대상내역으로저장" onclick="SaveArr()"></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>


	<% if gubun="witak" then %>
	<!-- 위탁 -->
	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
	    <tr align="center" bgcolor="#DDDDFF">
	      <td width="20" ><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
	      <td width="120" colspan="3">입고코드</td>
	      <td width="80">총매입가</td>
	      <td width="40">수량</td>
	      <td width="80">입고일</td>
	      <td width="120">등록일</td>
	    </tr>
	    <% for i=0 to ojungsan.FResultCount-1 %>
	    <% if precode<>ojungsan.FItemList(i).FCode then %>
	    <%
	    	ischeckd = false
	    	if Not IsNull(ojungsan.FItemList(i).FExecuteDate) then
	    		ischeckd = ((Cdate(ojungsan.FItemList(i).FExecuteDate)<checkdate1) and (Cdate(ojungsan.FItemList(i).FExecuteDate)>=checkdate2))
	    	end if
	    %>
	    <tr align="center" bgcolor="#FFFFFF">
	      <td></td>
	      <td align="left" colspan="3"><b><%= ojungsan.FItemList(i).FCode %></b></td>
	      <td align="right"><b><%= FormatNumber(ojungsan.FItemList(i).FTotalsuplycash,0) %></b></td>
	      <td></td>
	      <td><%= ojungsan.FItemList(i).FExecuteDate %></td>
	      <td><%= ojungsan.FItemList(i).FRegDate %></td>
	    </tr>
	    <% end if %>
	    <% precode = ojungsan.FItemList(i).FCode %>
	    <form name="frmBuyPrc_<%= i %>" >
	    <input type="hidden" name="idx" value="<%= ojungsan.FItemList(i).FID %>">
	    <tr align="center" bgcolor="#FFFFFF" <% if ischeckd then response.write "class='H'" %> >
	      <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <% if ischeckd then response.write "checked" %> ></td>
	      <td width="60"></td>
	      <td align="left"><%= ojungsan.FItemList(i).FItemName %></td>
	      <td><%= ojungsan.FItemList(i).FItemOptionName %></td>
	      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).FSuplycash,0) %></td>
	      <td align="center"><%= ojungsan.FItemList(i).FItemNo %></td>
	      <td></td>
	      <td></td>
	    </tr>
	    </form>
	    <% next %>
	</table>
	<% elseif gubun="witakchulgo" then %>
	<!-- 위탁 출고 -->
	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
	    <tr align="center" bgcolor="#DDDDFF">
	      <td width="20" ><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
	      <td width="120" colspan="3">입고코드</td>
	      <td width="80">총매입가</td>
	      <td width="40">수량</td>
	      <td width="80">출고예정일</td>
	      <td width="120">등록일</td>
	      <td></td>
	    </tr>
	    <% for i=0 to ojungsan.FResultCount-1 %>
	    <% if precode<>ojungsan.FItemList(i).FCode then %>
	    <%
	    	ischeckd = false
	    	if Not IsNull(ojungsan.FItemList(i).FScheduleDate) then
	    		ischeckd = ((Cdate(ojungsan.FItemList(i).FScheduleDate)<checkdate1) and (Cdate(ojungsan.FItemList(i).FScheduleDate)>=checkdate2))
	    	end if
	    %>
	    <tr align="center" bgcolor="#FFFFFF">
	      <td></td>
	      <td align="left"colspan="3"><b><%= ojungsan.FItemList(i).FCode %></b></td>
	      <td align="right"><b><%= FormatNumber(ojungsan.FItemList(i).FTotalsuplycash,0) %></b></td>
	      <td></td>
	      <td><%= ojungsan.FItemList(i).FScheduleDate %></td>
	      <td><%= ojungsan.FItemList(i).FRegDate %></td>
	      <td ><%= ojungsan.FItemList(i).FDesignerID %></td>
	    </tr>
	    <% end if %>
	    <% precode = ojungsan.FItemList(i).FCode %>
	    <form name="frmBuyPrc_<%= i %>" >
	    <input type="hidden" name="idx" value="<%= ojungsan.FItemList(i).FID %>">
	    <tr align="center" bgcolor="#FFFFFF" <% if ischeckd then response.write "class='H'" %> >
	      <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <% if ischeckd then response.write "checked" %> ></td>
	      <td width="60"></td>
	      <td><%= ojungsan.FItemList(i).FItemName %></td>
	      <td><%= ojungsan.FItemList(i).FItemOptionName %></td>
	      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).FSuplycash,0) %></td>
	      <td align="center"><%= ojungsan.FItemList(i).FItemNo %></td>
	      <td></td>
	      <td></td>
	      <td></td>
	    </tr>
	    </form>
	    <% next %>
	</table>
	<% elseif gubun="witaksell" then %>
	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
	    <tr align="center" bgcolor="#DDDDFF">
	      <td width="20" ><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
	      <td width="80">주문번호</td>
	      <td width="60">구매자</td>
	      <td>아이템명</td>
	      <td>옵션</td>
	      <td width="30">갯수</td>
	      <td width="60">판매가</td>
	      <td width="60">공급가</td>
	      <td width="60">주문일</td>
	      <td width="60">입금일</td>
	      <td width="60">배송일</td>
	      <td width="60">배송상태</td>
	      <td width="60">매입위탁</td>
	      <td width="60">과세면세</td>
	    </tr>
	    <% for i=0 to ojungsan.FResultCount-1 %>
	    <%
	    	ischeckd = false
	    	if Not IsNull(ojungsan.FItemList(i).FBeasongdate) then
	    		ischeckd = ((ojungsan.FItemList(i).FCurrState="6") or (ojungsan.FItemList(i).FCurrState="7")) and (Cdate(ojungsan.FItemList(i).FBeasongdate)<checkdate1) and (Cdate(ojungsan.FItemList(i).FBeasongdate)>=checkdate2)
	    	end if
			ischeckd = ischeckd and (ojungsan.FItemList(i).FMWDiv="W")

	    	'ischeckd = ischeckd or ((ojungsan.FItemList(i).FJumundiv="9") and (Cdate(ojungsan.FItemList(i).FRegDate)<checkdate1) and (Cdate(ojungsan.FItemList(i).FRegDate)>=checkdate2))
	    %>
	    <form name="frmBuyPrc_<%= i %>" >
	    <input type="hidden" name="idx" value="<%= ojungsan.FItemList(i).FIDX %>">
	    <tr align="center" bgcolor="#FFFFFF" <% if ischeckd then response.write "class='H'" %> >
	      <td width="20"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <% if ischeckd then response.write "checked" %> ></td>
	      <td ><a href="javascript:popOrderDetailEdit(<%= ojungsan.FItemList(i).FIDX %>);"><%= ojungsan.FItemList(i).Forderserial %></a></td>
	      <td ><%= ojungsan.FItemList(i).FBuyname %></td>
	      <td ><%= ojungsan.FItemList(i).FItemName %></td>
	      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
	      <td ><%= ojungsan.FItemList(i).FItemNo %></td>
	      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).FSellCash,0) %></td>
	      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).FBuyCash,0) %></td>
	      <td ><acronym title="<%= ojungsan.FItemList(i).FRegDate %>"><%= Left(ojungsan.FItemList(i).FRegDate,10) %></acronym></td>
	      <td ><acronym title="<%= ojungsan.FItemList(i).FIpkumDate %>"><%= Left(ojungsan.FItemList(i).FIpkumDate,10) %></acronym></td>
	      <td ><acronym title="<%= ojungsan.FItemList(i).FBeasongdate %>"><%= Left(ojungsan.FItemList(i).FBeasongdate,10) %></acronym></td>
	      <td >
	      <% if ojungsan.FItemList(i).FJumundiv="9" then %>
	      <font color="#FF33FF"><b>마이너스</b></font>
	      <% else %>
	      <font color="<%= IpkumDivColor(ojungsan.FItemList(i).FIpkumdiv) %>"><%= IpkumDivName(ojungsan.FItemList(i).FIpkumdiv) %></font>
	      <% end if %>
	      </td>
	      <td ><%= ojungsan.FItemList(i).FMWDiv %></td>
	      <td align="center"><%= ChkIIF(ojungsan.FItemList(i).FVatInclude="N","면세","") %></td>
	    </tr>
	    </form>
	    <% next %>
	</table>
	<% end if %>
<%
ojungsanselected.FRectdesigner = designer
ojungsanselected.FRectYYYYMM = yyyy1 + "-" + mm1
ojungsanselected.FRectgubun = "witak"
ojungsanselected.FRectdifferencekey = differencekey

if (designer<>"") then
    ojungsanselected.JungsanDetailListByYYYYMM
end if

dim sumtotal
%>
<br>
<table width="100%" cellspacing="1" class="a" >
<tr><td><hr></td></tr>
</table>
<div class="a"><b>등록된 위탁입고내역</b></div>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
	<% for i=0 to ojungsanselected.FResultCount-1 %>
    <% if precode<>ojungsanselected.FItemList(i).Fmastercode then %>
    <tr bgcolor="#FFFFFF">
      <td colspan="3"><%= ojungsanselected.FItemList(i).Fmastercode %></td>
      <td align="right"></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <% end if %>
    <% precode = ojungsanselected.FItemList(i).Fmastercode %>
    	<tr bgcolor="#FFFFFF"  >
	      <td width="60" align="right"><%= ojungsanselected.FItemList(i).FItemid %></td>
	      <td ><%= ojungsanselected.FItemList(i).FItemName %></td>
	      <td width="60"><%= ojungsanselected.FItemList(i).FItemOptionName %></td>
	      <td align="right"><%= FormatNumber(ojungsanselected.FItemList(i).Fsuplycash,0) %></td>
	      <td align="center"><%= ojungsanselected.FItemList(i).FItemNo %></td>
	      <td align="right"><%= FormatNumber(ojungsanselected.FItemList(i).Fsuplycash * ojungsanselected.FItemList(i).FItemNo,0) %></td>
	      <td></td>
	    </tr>
	<%
		sumtotal = sumtotal + ojungsanselected.FItemList(i).Fsuplycash * ojungsanselected.FItemList(i).FItemNo
	%>
    <% next %>
    <tr bgcolor="#DDDDFF"  >
      <td width="60">총계</td>
      <td></td>
      <td width="60"></td>
      <td align="right"></td>
      <td align="center"></td>
      <td align="right"><%= FormatNumber(sumtotal,0) %></td>
      <td></td>
    </tr>
</table>
<%
sumtotal =0

ojungsanselected.FRectgubun = "witakchulgo"
ojungsanselected.FRectdifferencekey = differencekey

if (designer<>"") then
    ojungsanselected.JungsanDetailListByYYYYMM
end if
%>
<br>
<table width="100%" cellspacing="1" class="a" >
<tr><td><hr></td></tr>
</table>
<div class="a"><b>등록된 위탁출고내역</b></div>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
	<% for i=0 to ojungsanselected.FResultCount-1 %>
    <% if precode<>ojungsanselected.FItemList(i).Fmastercode then %>
    <tr bgcolor="#FFFFFF">
      <td colspan="3"><%= ojungsanselected.FItemList(i).Fmastercode %></td>
      <td align="right"></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <% end if %>
    <% precode = ojungsanselected.FItemList(i).Fmastercode %>
    	<tr bgcolor="#FFFFFF"  >
	      <td width="60" align="right"><%= ojungsanselected.FItemList(i).FItemid %></td>
	      <td ><%= ojungsanselected.FItemList(i).FItemName %></td>
	      <td width="60"><%= ojungsanselected.FItemList(i).FItemOptionName %></td>
	      <td align="right"><%= FormatNumber(ojungsanselected.FItemList(i).Fsuplycash,0) %></td>
	      <td align="center"><%= ojungsanselected.FItemList(i).FItemNo %></td>
	      <td align="right"><%= FormatNumber(ojungsanselected.FItemList(i).Fsuplycash * ojungsanselected.FItemList(i).FItemNo,0) %></td>
	      <td></td>
	    </tr>
	<%
		sumtotal = sumtotal + ojungsanselected.FItemList(i).Fsuplycash * ojungsanselected.FItemList(i).FItemNo
	%>
    <% next %>
    <tr bgcolor="#DDDDFF"  >
      <td width="60">총계</td>
      <td></td>
      <td width="60"></td>
      <td align="right"></td>
      <td align="center"></td>
      <td align="right">
      	<%= FormatNumber(sumtotal,0) %>
      </td>
      <td></td>
    </tr>
</table>

<%
ojungsanselected.FRectgubun = "witaksell"
ojungsanselected.FRectdifferencekey = differencekey

if (designer<>"") then
    ojungsanselected.JungsanDetailListByYYYYMM
end if
%>
<br>
<table width="100%" cellspacing="1" class="a" >
<tr><td><hr></td></tr>
</table>
<div class="a"><b>등록된 위탁배송 정산 대상 내역</b></div>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
	<tr bgcolor="#DDDDFF">
      <td width="80">주문번호</td>
      <td width="50">구매자</td>
      <td width="50">수령인</td>
      <td width="120">아이템명</td>
      <td width="80">옵션명</td>
      <td width="40">(현재)<br>과세</td>
      <td width="40">갯수</td>
      <td width="50">판매가</td>
      <td width="50">공급가</td>
      <td width="50">삭제</td>
    </tr>
    <% for i=0 to ojungsanselected.FResultCount-1 %>
    <tr bgcolor="#FFFFFF">
      <td ><%= ojungsanselected.FItemList(i).Fmastercode %></td>
      <td ><%= ojungsanselected.FItemList(i).FBuyname %></td>
      <td ><%= ojungsanselected.FItemList(i).FReqname %></td>
      <td ><%= ojungsanselected.FItemList(i).FItemName %></td>
      <td ><%= ojungsanselected.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= ChkIIF(ojungsanselected.FItemList(i).Fvatinclude="N","면세","") %></td>
      <td ><%= ojungsanselected.FItemList(i).FItemNo %></td>
      <td ><%= ojungsanselected.FItemList(i).Fsellcash %></td>
      <td ><%= ojungsanselected.FItemList(i).Fsuplycash %></td>
      <td >삭제</td>
    </tr>
    <% next %>
</table>


<% elseif gubun="lecture" then %>
<!-- 강좌 정산대상 리스트 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="right">
        	<input type="button" value="정산대상내역으로 저장" onclick="SaveArr()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
    <tr align="center" bgcolor="#DDDDFF">
      <td width="20" ><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
      <td width="80">주문번호</td>
      <td width="50">구매자</td>
      <td>강좌명</td>
      <td>옵션</td>
      <td width="30">갯수</td>
      <td width="60">판매가</td>
      <td width="60">매입가</td>
      <td width="70">주문일</td>
      <td width="70">입금일</td>
      <td width="70">배송일</td>
      <td width="60">배송상태</td>
      <td width="60">강좌월</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    	ischeckd = false
    	if Not IsNull(ojungsan.FItemList(i).Flec_date) then
    		ischeckd = (ojungsan.FItemList(i).Flec_date=(YYYY1 + "-" + MM1))
    	end if

    	ischeckd = ischeckd or ((ojungsan.FItemList(i).FJumundiv="9"))
    %>
    <form name="frmBuyPrc_<%= i %>" >
    <input type="hidden" name="idx" value="<%= ojungsan.FItemList(i).FIDX %>">
    <tr align="center" bgcolor="#FFFFFF" <% if ischeckd then response.write "class='H'" %> >
      <td width="20"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <% if ischeckd then response.write "checked" %> ></td>
      <td ><%= ojungsan.FItemList(i).Forderserial %></td>
      <td ><%= ojungsan.FItemList(i).FBuyname %></td>
      <td align="left"><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td ><%= ojungsan.FItemList(i).FItemNo %></td>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).FSellCash,0) %></td>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).FBuyCash,0) %></td>
      <td ><acronym title="<%= ojungsan.FItemList(i).FRegDate %>"><%= Left(ojungsan.FItemList(i).FRegDate,10) %></acronym></td>
      <td ><acronym title="<%= ojungsan.FItemList(i).FIpkumDate %>"><%= Left(ojungsan.FItemList(i).FIpkumDate,10) %></acronym></td>
      <td ><acronym title="<%= ojungsan.FItemList(i).FBeasongdate %>"><%= Left(ojungsan.FItemList(i).FBeasongdate,10) %></acronym></td>
      <td >
      <% if ojungsan.FItemList(i).FJumundiv="9" then %>
      <font color="#FF33FF"><b>마이너스</b></font>
      <% else %>
      <font color="<%= UpCheBeasongStateColor(ojungsan.FItemList(i).FCurrState) %>"><%= UpCheBeasongState2Name(ojungsan.FItemList(i).FCurrState) %></font>
      <% end if %>
      </td>
      <td ><%= ojungsan.FItemList(i).Flec_date %></td>
    </tr>
    </form>
    <% next %>
</table>
<%
ojungsanselected.FRectdesigner = designer
ojungsanselected.FRectYYYYMM = yyyy1 + "-" + mm1
ojungsanselected.FRectgubun = "upche"
ojungsanselected.FRectdifferencekey = differencekey

if (designer<>"") then
    ojungsanselected.JungsanDetailListByYYYYMM
end if
%>
<br>
<table width="100%" cellspacing="1" class="a" >
<tr><td><hr></td></tr>
</table>
<div class="a"><b>등록된 강좌 정산 대상 내역</b></div>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="BABABA">
	<tr bgcolor="#DDDDFF">
      <td width="80">주문번호</td>
      <td width="50">구매자</td>
      <td width="50">수령인</td>
      <td width="120">아이템명</td>
      <td width="80">옵션명</td>
      <td width="40">갯수</td>
      <td width="50">판매가</td>
      <td width="50">공급가</td>
      <td width="50">삭제</td>
    </tr>
    <% for i=0 to ojungsanselected.FResultCount-1 %>
    <tr bgcolor="#FFFFFF">
      <td ><%= ojungsanselected.FItemList(i).Fmastercode %></td>
      <td ><%= ojungsanselected.FItemList(i).FBuyname %></td>
      <td ><%= ojungsanselected.FItemList(i).FReqname %></td>
      <td ><%= ojungsanselected.FItemList(i).FItemName %></td>
      <td ><%= ojungsanselected.FItemList(i).FItemOptionName %></td>
      <td ><%= ojungsanselected.FItemList(i).FItemNo %></td>
      <td ><%= ojungsanselected.FItemList(i).Fsellcash %></td>
      <td ><%= ojungsanselected.FItemList(i).Fsuplycash %></td>
      <td >삭제</td>
    </tr>
    <% next %>
</table>
<% end if %>
<%
set ojungsan = Nothing
set ojungsanselected = Nothing
%>
<form name="frmArrupdate" method="post" action="dodesignerjungsan.asp">
<input type="hidden" name="idx" value="">
<input type="hidden" name="gubun" value="<%= gubun %>">
<input type="hidden" name="designer" value="<%= designer %>">
<input type="hidden" name="yyyy1" value="<%= yyyy1 %>">
<input type="hidden" name="mm1" value="<%= mm1 %>">
<input type="hidden" name="differencekey" value="">
<input type="hidden" name="taxtype" value="">
<input type="hidden" name="mode" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->