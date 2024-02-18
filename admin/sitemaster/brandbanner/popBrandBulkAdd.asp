<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim vidx
vidx = requestCheckVar(Request("idx"),10)

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<script language='javascript'>
function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function searchFrm(p){
	frm1.page.value = p;
	frm1.submit();
}

//브랜드 ID 검색 팝업창
function jsBrandBannerSearchBrandID(){
    var popwin2 = window.open("popBrandSearch.asp?idx=<%=vidx%>","popBrandSearch2","width=800 height=400 scrollbars=yes resizable=yes");
	popwin2.focus();
}

//브랜드 ID 등록 팝업창
function fnBrandBulkADD(){
    var popwin3 = window.open("popBrandBulkAdd.asp?idx=<%=vidx%>","popBrandSearch3","width=800 height=400 scrollbars=yes resizable=yes");
	popwin3.focus();
}

function chkAllItem() {
	if($("input[name='chkIdx']:first").attr("checked")=="checked") {
		$("input[name='chkIdx']").attr("checked",false);
	} else {
		$("input[name='chkIdx']").attr("checked","checked");
	}
}

function fnBrandBulkADD(){
	if($("#idarr").val()=="") {
		alert("브랜드 ID를 입력해 주세요.");
		return;
	}
	if(confirm("브랜드를 등록 하시겠습니까?")) {
		document.frmList.mode.value="idarr";
		document.frmList.action="addbrandproc.asp";
		document.frmList.submit();
	}
}

function fnreload(){
	window.location.reload();
}
</script>
</head>
<body>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<div class="contSectFix scrl">
	<div class="cont">
		<div class="pad20">
			<div>
				<form name="frmList" method="post">
				<input type="hidden" name="mode">
				<input type="hidden" name="idx" value="<%=vidx%>">
				<table class="tbType1 listTb">
					<thead>
					<tr>
						<th colspan="2"><div>브랜드 일괄 등록</div></th>
					</tr>
					</thead>
					<tbody>
							<tr>
								<td>브랜드 아이디</td>
								<td class="lt">
                                    <textarea name="idarr" id="idarr" cols="50" rows="10"></textarea><br>
                                    <span class="cRd1">※ 브랜드 ID를 콤마(,)로 구분하여 공백없이 입력해주세요. (예:aaa,bbb,ccc)</span>
                                </td>
							</tr>
					</tfoot>
				</table>
				</form>
                <div align="right">
                <input type="button" class="btn" value="등록" onClick="fnBrandBulkADD();" />
                </div>
			</div>
		</div>
	</div>
</div>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->