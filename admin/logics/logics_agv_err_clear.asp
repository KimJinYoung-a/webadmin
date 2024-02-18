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
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<style type="text/css">
.dx-widget {font-size:12px;}
</style>
<script type="text/javascript">
function goSearch(frm) {
    frm.submit();
}

function jsReset(frm) {
    frm.skuCd.value = '';
}

function jsSendErrClear(frm) {
    var url;
    var skuCd, skuCdArr;

    skuCd = frm.skuCd.value;
    // skuCdArr = skuCd.replace('\n', ',');
    skuCdArr = skuCd.replace(/\n/gi, ',');

    if (skuCdArr == '') {
        alert('전송할 SKU코드가 없습니다.');
        return;
    }

    // alert(skuCdArr);

    url = 'https://wapi.10x10.co.kr';
    url = url + '/agv/api.asp?mode=agvSendErrClear&skuCdArray=' + skuCdArr;

    $.ajax({
        url: url,
        type: 'get',
        crossDomain: true,
        data: {},
        contentType: "application/x-www-form-urlencoded; charset=UTF-8",
        dataType: 'json',
        success: function(data) {
            if (data.resultCode == '200') {
                alert('전송되었습니다.');
            } else {
                alert(data.resultMessage);
            }
        },
        error: function(jqXHR, textStatus, ex) {
            alert(textStatus + "," + ex + "," + jqXHR.responseText);
        }
    });
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
                * SKU코드
                <textarea cols="16" rows="12" name="skuCd"></textarea>
            </td>
            <td align="right">
                <a href="#" onClick="jsReset(document.frm);" title="검색 조건을 초기화합니다.">초기화</a>
            </td>
        </tr>
        </table>
    </td>
    <td width="80" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="전송" onClick="jsSendErrClear(document.frm);">
	</td>
</tr>
</table>
</form>
<!-- 검색 영역 끝 -->
<br />
* 오차내역을 초기화할 SKU코드를 입력하세요.<br />
* 초기화하는 경우, AGV 실사재고가 AGV 시스템재고가 됩니다.
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
