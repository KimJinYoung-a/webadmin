<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : agv
' History : 이상구 생성
'           2020.05.12 정태훈 수정
'           2020.05.20 한용민 수정
'####################################################
%>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_logisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/barcodefunction.asp"-->
<!-- #include virtual="/lib/classes/logistics/logistics_agvCls.asp"-->
<%
dim page, i
dim itembarcode, found
dim itemgubun, itemid, itemoption, isUsing, genBarcode
itembarcode 	= requestCheckVar(request("itembarcode"),32)
itemgubun 	= requestCheckVar(request("itemgubun"),2)
itemid 	= requestCheckVar(request("itemid"),10)
itemoption = requestCheckVar(request("itemoption"),4)
isUsing = requestCheckVar(request("isusing"),1)
page = chkIIF(page="",1,page)
isUsing = chkIIF(isUsing="","Y",isUsing)

if (itembarcode <> "") then
	if Len(itembarcode) > 8 or Not IsNumeric(itembarcode) then
		'// 길이가 8보다 크거나 숫자가 아닌경우 바코드인지 먼저 확인
		found = fnGetItemCodeByPublicBarcode(itembarcode, itemgubun, itemid, itemoption)
	end if

	if Not found and BF_IsMaybeTenBarcode(itembarcode) = True then
		'// 텐바이텐 : 물류코드 검색의 경우(10 111111 0000 또는 10 01000000 0000)
		itemgubun 	= BF_GetItemGubun(itembarcode)
		itemid 		= BF_GetItemId(itembarcode)
		itemoption 	= BF_GetItemOption(itembarcode)
		found = True
	end if

	if Not found and Len(itembarcode) <= 8 and IsNumeric(itembarcode) then
		'상품코드로 검색(111111 또는 1000000)
		itemgubun = "10"
		itemid = itembarcode
		itemoption  = "0000"
		itembarcode = BF_MakeTenBarcode(itemgubun, itemid, itemoption)
	end if

end if

dim oAGV
Set oAGV = new CAGVItems
    oAGV.FPageSize = 10000
    oAGV.FCurrPage = page
	oAGV.FRectItemGubun = itemgubun
	oAGV.FRectItemID  =itemid
	oAGV.FRectItemoption = itemoption
    oAGV.FRectIsUsing = isUsing
    oAGV.GetShelfItemList
%>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>

<link rel="stylesheet" type="text/css" href="https://cdn3.devexpress.com/jslib/19.1.4/css/dx.common.css" />
<link rel="stylesheet" type="text/css" href="https://cdn3.devexpress.com/jslib/19.1.4/css/dx.light.compact.css" />
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.1.2/jszip.min.js"></script>
<script src="https://cdn3.devexpress.com/jslib/19.1.4/js/dx.all.js"></script>
<style type="text/css">
.dx-widget {font-size:12px;}
</style>
<script type="text/javascript">
function goSearch(frm) {
    frm.submit();
}

function resetSearchForm(frm) {
    frm.itembarcode.value="";
    frm.itemgubun.value="";
    frm.itemid.value="";
    frm.itemoption.value="";
    frm.isusing.value="Y";
    goSearch(frm);
}
</script>
<!-- 검색 영역 시작 -->
<form name="frm" method="get" >
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#F4F4F4">
    <td width="50" bgcolor="#EEEEEE">검색<br>조건</td>
	<td>
        <table width="100%" class="a">
        <tr>
            <td align="left">
                <label>· 바코드 : <input type="text" name="itembarcode" size="16" value="<%=itembarcode%>" class="text" /> </label> &nbsp;
                <label>· 구분 : <input type="text" name="itemgubun" size="2" value="<%=itemgubun%>" class="text" /> </label> &nbsp;
                <label>· 상품코드 : <input type="text" name="itemid" size="10" value="<%=itemid%>" class="text" /> </label> &nbsp;
                <label>· 옵션코드 : <input type="text" name="itemoption" size="4" value="<%=itemoption%>" class="text" /> </label> &nbsp;
                <label>· 사용여부 :
                    <select name="isusing" class="select">
                    <option value="A" <%=chkIIF(isUsing="A","selected","")%>>모두</option>
                    <option value="Y" <%=chkIIF(isUsing="Y","selected","")%>>사용</option>
                    <option value="N" <%=chkIIF(isUsing="N","selected","")%>>삭제</option>
                    </select>
                </label>
            </td>
            <td align="right">
                <a href="#" onClick="resetSearchForm(document.frm);" title="검색 조건을 초기화합니다.">Reset</a>
            </td>
        </tr>
        </table>
    </td>
    <td width="80" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="검색" onClick="goSearch(document.frm);">
	</td>
</tr>
</table>
</form>
<!-- 검색 영역 끝 -->
<!-- 데이터 그리드 시작 -->
<div class="dx-viewport" style="margin-top:5px;padding:5px;">
    <div class="demo-container">
        <div id="gridContainer"><center>No Data...</center></div>
    </div>
</div>
<!-- 데이터 그리드 끗 -->
<script type="text/javascript">
$(function(){
    $("#gridContainer").dxDataGrid({
        showColumnLines: true, // 컬럼 라인
        showRowLines: true, // 로우 라인
        rowAlternationEnabled: true, // 로우별 회색 색상
        showBorders: true, // 전체 보더
        columnChooser: { // 화면에 보여주는 컬럼 선택
            enabled: true,
            mode: "select" // or "dragAndDrop"
        },
        allowColumnReordering: true, // 컬럼 순서 변경
        "export": { // 엑셀 다운로드 관련
            enabled: true,
            fileName: "EmailCustomerList",
            allowExportSelectedData: true
        },
        headerFilter: { // 컬럼명 깔대기 검색 
            visible: true
        },
        columnAutoWidth: true,
        columns: [
            {dataField : "일련번호",alignment : "center",dataType: "number",fixed: true},
            {dataField : "바코드",alignment : "center",dataType: "string"},
            {dataField : "예정수량",alignment : "center",dataType: "number",format: "fixedPoint"},
            {dataField : "입고수량",alignment : "center",dataType: "number",format: "fixedPoint"},
            {dataField : "등록일",alignment : "center",dataType: "date"},
            {dataField : "수정일",alignment : "center",dataType: "date"},
            {dataField : "랙코드",alignment : "center",dataType: "string"},
            {dataField : "쉘프코드",alignment : "center",dataType: "string"},
            {dataField : "사용여부",alignment : "center",dataType: "string"},
            {dataField : "상태",alignment : "center",dataType: "string"},
        ],
        dataSource: [
        <% For i = 0 To oAGV.FResultCount-1 %>
            {"일련번호":<%=oAGV.FItemList(i).FIdx%>,
            "바코드":"<%=oAGV.FItemList(i).FItemGubun &"-"&Num2Str(oAGV.FItemList(i).FItemid,6,"0","R")&"-"&oAGV.FItemList(i).FItemOption%>",
            "예정수량":<%=oAGV.FItemList(i).FRealStock%>,
            "입고수량":<%=oAGV.FItemList(i).FfixedStock%>,
            "등록일":"<%=oAGV.FItemList(i).FRegdate%>",
            "수정일":"<%=oAGV.FItemList(i).Flastupdate%>",
            "랙코드":"<%=oAGV.FItemList(i).FRackCode%>",
            "쉘프코드":"<%=oAGV.FItemList(i).FShelfCode%>",
            "사용여부":"<%=oAGV.FItemList(i).getIsUsing%>",
            "상태":"<%=oAGV.FItemList(i).getStatus%>"},
        <% Next %>
        ]
    });
});
</script>
<% Set oAGV = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->