<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/fingersUpcheAgreeCls.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<%
dim sqlStr
dim agreeIdx
dim makerid : makerid = session("ssBctID")
dim groupid : groupid = getPartnerId2GroupID(makerid)

agreeIdx  = requestCheckvar(request("agreeIdx"),10)


dim onecontract
set onecontract = new CFingersUpcheAgree
onecontract.FRectagreeIdx = agreeIdx

if agreeIdx<>"" then
    onecontract.FRectMakerid = makerid
    onecontract.getOneFingersUpcheAgree
end if

if onecontract.FResultCount<1 then
    response.write "권한이 없거나, 유효한 계약번호가 아닙니다."
    dbget.close()	:	response.End
end if


dim itypeName : itypeName = onecontract.FoneItem.getContractTypeAgreeName

dim isInvalidAgree : isInvalidAgree=false
isInvalidAgree = (LEN(onecontract.FOneItem.FContractContents)<10) ''계약서가 잘못생성되었을경우 대비

dim iUrlParam
iUrlParam = "agreeIdx="&agreeIdx&"&gkey="&onecontract.FoneItem.Fgroupid&"&ekey="&onecontract.FoneItem.getEkey&"&chkcf="&CHKIIF(onecontract.FoneItem.IsPrivContractAddItem,"1","")
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
function doAgree(comp){
    var frm = comp.form;
    if (!frm.agreechk.checked){
        alert('위 <%=itypeName%>에 동의합니다 에 체크해주세요.');
        //frm.agreechk.focus();
        return;   
    }
    
    if (confirm('위 <%=itypeName%>에 동의합니까?')){
        <% if (isInvalidAgree) then %>
        alert('죄송합니다. 현재 동의 불가 - 관리자 문의 요망.');
        <% else %>
        frm.submit(); 
        <% end if %>
    }
    
}
</script>
</head>
<body>
<div class="docuAgree">
	<iframe width="100%" height="100%" frameborder="0" src="ifrconfirmContract.asp?<%=iUrlParam%>"></iframe>
</div>
<form name="frmAgree" method="post" action="doAgree.asp">
<input type="hidden" name="mode" value="iagree">
<input type="hidden" name="agreeIdx" value="<%=agreeIdx%>">
<div class="floatBar">
	<p>
	    <% if (onecontract.FOneItem.IsAgreeFinished) then %>
	    <%= onecontract.FOneItem.getAgreeText %>
	    <% else %>
		<label><input type="checkbox" name="agreechk" /> 위 <%=itypeName%>에 동의합니다.</label> <button type="button" class="btnOk" onclick="doAgree(this)">확인</button>
		<% end if %>
	</p>
</div>
</form>
</body>
</html>

<%

if NOT onecontract.FOneItem.IsAgreeFinished then
    sqlStr="update db_partner.dbo.tbl_partner_fingers_agreeHist"&vbCRLF
    sqlStr=sqlStr&" set viewdate=isNULL(viewdate,getdate())"&vbCRLF
    sqlStr=sqlStr&" where agreeIdx="&agreeIdx
    dbget.Execute sqlStr
end if

set onecontract = Nothing

''업체가 다운로드 할 경우 확인일 플래그 업데이트
'if (chkcf="1") then
'    sqlStr=" update db_partner.dbo.tbl_partner_ctr_master"
'    sqlStr=sqlStr&" set confirmdate=IsNULL(confirmdate,getdate())"
'    sqlStr=sqlStr&" ,ctrState=(CASE WHEN ctrState in (1,2) then 3 else ctrState end )"
'    sqlStr=sqlStr&" where ctrKey="&ctrKey
'
'    ''dbget.Execute sqlStr
'end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
