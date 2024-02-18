<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 핑거스 계약 관리
' Hieditor : 2016.08.10 서동석 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/classes/partners/fingersUpcheAgreeCls.asp"-->
<%
dim sqlStr
dim agreeIdx

agreeIdx  = requestCheckvar(request("agreeIdx"),10)


dim onecontract
set onecontract = new CFingersUpcheAgree
onecontract.FRectDelInclude = "on"
onecontract.FRectagreeIdx = agreeIdx

if agreeIdx<>"" then
    onecontract.getOneFingersUpcheAgree
end if

if onecontract.FResultCount<1 then
    response.write "권한이 없거나, 유효한 계약번호가 아닙니다."
    dbget.close()	:	response.End
end if


dim itypeName : itypeName = onecontract.FoneItem.getContractTypeAgreeName
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<title></title>
<style type="text/css">
html, body, div, span, object, iframe, h1, h2, h3, h4, h5, h6, p, blockquote, pre, abbr, address, cite, code, del, dfn, em, img, ins, kbd, q, samp, small, strong, sub, sup, var, b, i, dl, dt, dd, ol, ul, li, fieldset, form, label, legend, table, caption, tbody, tfoot, thead, tr, th, td, article, aside, canvas, details, figcaption, figure, footer, header, hgroup, menu, nav, section, summary, time, mark, audio, video {margin:0; padding:0; border:0; font-size:100%; vertical-align:baseline;}
article, aside, details, figcaption, figure, footer, header, hgroup, menu, nav, section {display:block;}
body {line-height:1;}
ol, ul {list-style:none;}
fieldset, img {border:0;}
blockquote, q {quotes: none;}
blockquote:before, blockquote:after, q:before, q:after {content:''; content: none;}
table {border-collapse:collapse; border-spacing:0; empty-cells:show;}
del {text-decoration: line-through;}
input, select {vertical-align:middle;}
i, em, address {font-style:normal; font-weight:normal;}
html {font-size:12px; font-family:'돋움', dotum, helvetica, sans-serif; line-height:100%; overflow-x:auto; overflow-y:hidden;}
a {text-decoration:none;}

.docuAgree  {position:absolute; top:0; right:0; bottom:0; left:0; width:100%; height:100%;}
.floatBar {display:table; position:absolute; left:0; bottom:0; width:100%; height:55px; background-color:rgba(0,0,0,.5); color:#fff; text-align:center; font-size:12px;}
.floatBar p {display:table-cell; vertical-align:middle;}
.floatBar input[type=checkbox] {width:16px; height:16px; vertical-align:middle; margin-top:-2px;}
.btnOk {width:20%; height:35px; background-color: #d60000; border:none; color:#fff; font-weight:bold;}
</style>
<script language='javascript'>
function delThisContract(comp){
    var frm = comp.form;
    if (confirm('삭제하시겠습니까?')){
        frm.mode.value="del";
        frm.submit();
    }
}
</script>
</head>
<body>
<div class="docuAgree">
	<iframe width="100%" height="100%" frameborder="0" src="fingersAgreeView_ifr.asp?agreeIdx=<%=agreeIdx%>"></iframe>
</div>
<form name="frmAgree" method="post" action="doFingersAgree_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="agreeIdx" value="<%=agreeIdx%>">
<div class="floatBar">
	<p>
	    <% if (onecontract.FOneItem.IsAgreeFinished) then %>
	    <%= onecontract.FOneItem.getAgreeText %>
	    <% else %>
		현재 미동의 상태 입니다.
		<% if (onecontract.FOneItem.isDeletedContract) then %>
		(<strong>삭제된 내역입니다.</strong>)
	    <% else %>
		<input type="button" value=" 삭 제 " onclick="delThisContract(this)">
		<% end if %>
		<% end if %>
	</p>
</div>
</form>
</body>
</html>

<%


set onecontract = Nothing

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
