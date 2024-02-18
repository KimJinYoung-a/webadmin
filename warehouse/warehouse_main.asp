<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  물류센터 메인
' History : 이상구 생성
'			2022.01.19 한용민 수정(한진택배,출고보고 추가)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/checknoticls.asp"-->
<%
dim BasicMonth
BasicMonth = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)
%>

<script type='text/javascript'>

function publicbarreg(barcode){
	var popwin = window.open('/common/popbarcode_input.asp?itembarcode=' + barcode,'popbarcode_input','width=500,height=400,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function popBrandRackCodeEdit(imakerid){
    var popwin = window.open('pop_BrandRackCodeEdit.asp?makerid=' + imakerid,'popBrandRackCodeEdit','width=500,height=200,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function itemrackcodereg(itemrackcode){
	var popwin = window.open('popitemrackcode_input.asp?itemrackcode=' + itemrackcode,'popitemrackcode_input','width=300,height=300,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popItemRackCodeEdit(itemrackcode){
    var popwin = window.open('pop_ItemRackCodeEdit.asp?itemrackcode=' + itemrackcode,'pop_ItemRackCodeEdit','width=400,height=300,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popitemsearch(barcode){
	var popwin = window.open('pop_item_search.asp?barcode=' + barcode,'popitemsearch','width=500,height=400,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function popRealErrInput(itemgubun,itemid,itemoption){
	var popwin = window.open('/common/poprealerrinput.asp?itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&BasicMonth=<%= BasicMonth %>','poprealerrinput','width=900,height=460,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popRealStockTaking(iitemid){
    var popwin = window.open('/admin/stock/jaegoadd.asp?itemid='+ iitemid,'poprealstockTaking','width=900,height=460,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function PopBadItemInput(){
	var popwin = window.open('/common/pop_baditem_input.asp','pop_baditem_input','width=900,height=400,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'simpleitemedit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PopItemWeightEdit(iitemid){
	var popwin = window.open('pop_ItemWeightEdit.asp?itemid=' + iitemid,'itemWeightEdit','width=500,height=300,scrollbars=yes,resizable=yes')
}

function PopSongjangList(songjangdiv){
	var popwin = window.open('/warehouse/pop_SongjangList.asp?songjangdiv='+songjangdiv,'PopSongjangList','width=500,height=300,scrollbars=yes,resizable=yes')
}

function reSearchExcelDown(select_type){
    alert('기다려 주세요. 작성중입니다.');
	frm.target = "exceldown";
	frm.action = "/admin/ordermaster/chulgobogo_ExcelDown.asp"
    frm.select_type.value=select_type;
    frm.submit();
	frm.target = "";
	frm.action = "";
    frm.select_type.value='';
}

</script>


<a href="javascript:publicbarreg('');">범용바코드등록</a>
<br>
<a href="javascript:popBrandRackCodeEdit('');">브랜드랙코드등록</a>
<br>
<a href="javascript:itemrackcodereg('');">랙코드별상품입력</a>
<br>
<a href="javascript:popItemRackCodeEdit('');">상품별랙코드입력</a>
<br>
<a href="javascript:popRealStockTaking('');">재고조사</a>
<br>
<a href="javascript:popRealErrInput('','','');">재고(오차)입력</a>

<br>
<a href="javascript:PopBadItemInput();">불량등록</a>
<br>
<a href="javascript:popitemsearch('');">상품검색</a>
<br>
<a href="javascript:PopItemSellEdit('');">상품속성수정</a>
<br>
<a href="javascript:PopItemWeightEdit('');">상품무게입력</a>
<br>
송장목록 엑셀받기 : 
<a href="#" onclick="PopSongjangList('2'); return false;">롯데택배</a>
/ <a href="#" onclick="PopSongjangList('1'); return false;">한진택배</a>
/ <a href="#" onclick="PopSongjangList('4'); return false;">CJ대한통운</a>

<hr>
<a href="undeliveredOrderList.asp">미배송 주문 목록</a>
<hr>
<form name="frm" method="get" action="" style="margin:0px;" >
<input type="hidden" name="research" value="on">
<input type="hidden" name="select_type" value="">
<br>
* 출고보고(날짜기준:전일)
<%
'<input type="button" class="button_s" value="엑셀다운로드(당일미출고주문)" onclick="reSearchExcelDown('samedaymichulgo');">
'<input type="button" class="button_s" value="엑셀다운로드(지연출고주문)" onclick="reSearchExcelDown('delaychulgo');">
'<input type="button" class="button_s" value="엑셀다운로드(지연출고_결제일빠른날짜)" onclick="reSearchExcelDown('delaychulgodate');">
'<input type="button" class="button_s" value="엑셀다운로드(지연출고_결제일빠른주문)" onclick="reSearchExcelDown('delaychulgocnt');">
%>
</form>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe src="about:blank" name="exceldown" border="0" width="100%" height="300"></iframe>
<% else %>
	<iframe src="about:blank" name="exceldown" border="0" width="100%" height="0"></iframe>
<% end if %>



















<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->