<%@ language=vbscript %>
<% option explicit %>

<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>

<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim idxArr,iSall
idxArr = Replace(request.Form("idxArr"), " ", "")
idxArr = Trim(idxArr)
iSall   =  requestCheckVar(request("isall"), 32)

if (Right(idxArr,1)=",") then idxArr=Left(idxArr,Len(idxArr)-1)

if (Len(idxArr)<1) and (iSall="") then
    response.write "<script>alert('선택된 주문건이 없습니다.');</script>"
    dbget.close()	:	response.End
end if
%>
<script language='javascript'>
function popDeliverCode(){
    var popwin = window.open('popDeliverCode.asp','popDeliverCode','width=400,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function downloadDeliveXL(){
    var xlfrm = document.xlfrm;
	xlfrm.target="iiframeXL";
	xlfrm.action="upchecs_songjanglistexcel.asp";
	xlfrm.submit();
}

function NextStep(frm){
    if (frm.songjangfile.value.length<1){
        alert('업로드할 CSV파일을 선택하세요.')
        return;
    }

    if (confirm('다음 단계로 진행 하시겠습니까?')){
        frm.submit();
    }
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
    <td>
        1. 엑셀 양식을 다운받아 <strong>택배사 코드</strong>와 <strong>송장 번호</strong>를 입력 하신 후 업로드 하시면 일괄 발송 처리 됩니다.<br>
        2. 송장번호 입력 후 저장은 <strong>CSV 형식으로 저장</strong> 하시기 바랍니다.<br>
        3. 업체 정보 수정에서 <strong>기본 택배사를 지정</strong> 해 놓으시면 택배사 코드가 기본으로 지정되어 다운로드 할 수 있습니다.<br>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="right">
    <a href="javascript:popDeliverCode();"><font color="blue">[택배사 코드 보기]</font></a>
    </td>
</tr>
</table>
<p>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="50">
    <td width="100" align="center">Step 1</td>
    <td>엑셀 양식을 다운 받으세요. (미출고 내역만 작성됩니다.)<a href="javascript:downloadDeliveXL();"><font color="blue">[다운로드]</font></a>
        <br>엑셀양식내의 택배사코드는 기본택배사로 지정되오니, 유의하시기 바랍니다.
        <br>기본택배사는 업체정보 수정에서 수정가능합니다.
    </td>
</tr>
</table>

<p>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="50">
    <td width="100" align="center">Step 2</td>
    <td>파일을 열어 송장 번호를 입력 하신 후 개인 PC에 CSV 파일로 저장하세요.</td>
</tr>
</table>
<p>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmNext" method="post" action="upchecs_pop_BatchSongjangInputStep2.asp" onsubmit="return false;" enctype="multipart/form-data">
<tr bgcolor="#FFFFFF" height="50">
    <td width="100" align="center">Step 3</td>
    <td>저장한 CSV 파일을 지정하신 후 다음단계로 이동하세요.
        <input type="file" name="songjangfile" size="30" value="">
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
    <td colspan="2" align="center">
    <input type="button" value="다음단계로 진행" onClick="NextStep(frmNext)">
    </td>
</tr>
</form>
</table>
<iframe name="iiframeXL" name="iiframeXL" width="110" height="110" frameborder=0 scrolling=no marginheight=0 marginwidth=0 align=center></iframe>
<form name=xlfrm method=post action="">
<input type="hidden" name="idxArr" value="<%= idxArr %>">
<input type="hidden" name="iSall" value="<%= iSall %>">
</form>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
