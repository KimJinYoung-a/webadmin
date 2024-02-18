<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/naverEp/epShopCls.asp"-->
<%
dim exceptTp : exceptTp = requestCheckvar(request("exceptTp"),10)
dim makerid  : makerid = requestCheckvar(request("makerid"),32)
dim itemid  : itemid = requestCheckvar(request("itemid"),10)
dim mode : mode = requestCheckvar(request("mode"),10)
if (mode="") then mode="I"

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript" src="/js/jsCal/js/jscal2.js"></script>
<script language="javascript" src="/js/jsCal/js/lang/ko.js"></script>
<script>
$(function() {
	var CAL_iAasignMxDt = new Calendar({
		inputField : "iAasignMxDt", trigger    : "iAasignMxDt_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
});

function addNvCpnExceptBrand(){
    var frm = document.frmAct;

    if (frm.makerid.value.length<1){
        alert("제외할 브랜드를 선택하세요.");
        frm.makerid.focus();
        return;
    }

    if (confirm("네이버 쿠폰 제외 브랜드를 등록하시겠습니까?")){
        frm.submit();
    }

}

function addNvCpnExceptItem(){
    var frm = document.frmAct;

    if (frm.itemid.value.length<1){
        alert("제외할 상품을 선택하세요.");
        frm.itemid.focus();
        return;
    }

    if (confirm("네이버 쿠폰 제외 상품을 등록하시겠습니까?")){
        frm.submit();
    }

}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmAct" method="post" action="exceptNvCpn_process.asp">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="excepttp" value="<%=excepttp%>">
<% if (exceptTp="B") then %>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td >브랜드ID</td>
        <td bgcolor="#FFFFFF"><input type="text" name="makerid" value="<%=makerid%>" <%=CHKIIF(makerid<>"","readonly","")%>  size="16" maxlength="32">
        <input type="button" class="button" value="ID검색" onclick="jsSearchBrandID(this.form.name,'makerid');" >
        </td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td>제외 기간</td>
        <td bgcolor="#FFFFFF">
            ~ <input type="text" id="iAasignMxDt" name="iAasignMxDt" value="" size="10" maxlength="10">
            <img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='iAasignMxDt_trigger' border='0' style='cursor:pointer' align='absmiddle' />
            (빈값일 경우 무기한)
        </td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
        <td colspan="2"><input type="button" value="네이버 쿠폰 불가 브랜드 등록" onClick="addNvCpnExceptBrand()"></td>
    </tr>
<% else %>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td >상품코드</td>
        <td bgcolor="#FFFFFF"><input type="text" name="itemid" value="<%=itemid%>" <%=CHKIIF(itemid<>"","readonly","")%> size="8" maxlength="10" ></td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td>제외 기간</td>
        <td bgcolor="#FFFFFF">
            ~ <input type="text" id="iAasignMxDt" name="iAasignMxDt" value="" size="10" maxlength="10">
            <img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='iAasignMxDt_trigger' border='0' style='cursor:pointer' align='absmiddle' />
            (빈값일 경우 무기한)
        </td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
        <td colspan="2"><input type="button" value="네이버 쿠폰 불가 상품 등록" onClick="addNvCpnExceptItem()"></td>
    </tr>
<% end if %>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

