<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  매장고객방문카운트
' History : 2012.05.10 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/guest/shop_guestcount_cls.asp"-->
<%
dim iMaxLength , memupos
	memupos = requestCheckVar(request("memupos"),10)
	iMaxLength = 5
%>

<script language="javascript">

function fnChkFile(sFile, arrExt){
    //파일 업로드 유무확인
     if (!sFile){
    	 return true;
    	}

    var blnResult = false;

    //파일 확장자 확인
	var pPoint = sFile.lastIndexOf('.');
	var fPoint = sFile.substring(pPoint+1,sFile.length);
	var fExet = fPoint.toLowerCase();

	for (var i = 0; i < arrExt.length; i++)
	   	{
	    	if (arrExt[i].toLowerCase() == fExet)
	    	{
	   			blnResult =  true;
	   		}
		}

	return blnResult;
}

function XLSumbit(){
	var frm = document.frmFile;
    
	arrFileExt = new Array();			
	arrFileExt[arrFileExt.length]  = "xls";
	
	if (frm.sFile.value==''){
		alert('파일을 입력해 주세요');
		frm.sFile.focus();
		return;
	}
	
	//파일유효성 체크
	if (!fnChkFile(frm.sFile.value, arrFileExt)){
		alert("파일은 xls파일만 업로드 가능합니다.");
		return;
	}
	
	frm.submit();
}

</script>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmFile" method="post" action="<%=uploadUrl%>/linkweb/offshop/guest/shopguestcount/shop_guestcount_upload.asp"  enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="iML" value="<%=iMaxLength%>">
<input type="hidden" name="mode" value="excelupload">
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2">
		<b>매장 고객방문 엑셀등록</b>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>필독!!</td>
	<td align="left">
		매장 고객방문 클라이언트(remote manager)를 실행,
		<Br><Br>로그인(마스타 관리자 로그인 체크후 비밀번호 1)완료후, 왼쪽 하단에 피플카운트를 클릭,
		<br><Br>상단 오른쪽에 데이터리포트 클릭후, 그래프를 시간대별 리포트로 선택,
		<br><Br>하단에 사업장(매장)과 출력할 월을 선택후 적용 클릭,
		<Br><Br>그래프에 파일저장을 클릭해서 다운로드후, 다운로드 받은 내역을,
		<br><Br><font color="red">새로운 엑셀시트에 복사&붙여넣기후 Excel 97 -2003 통합문서로 저장</font>후 등록하시면 됩니다
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>샘플</td>
	<td align="left"><a href="/common/offshop/guest/sample.xls" onfocus="this.blur()">sample.xls</a></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>파일명:</td>
	<td align="left"><input type="file" name="sFile" class="button"></td>
</tr>				
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2">
	    <input type="button" class="button" value="등록" onClick="XLSumbit();">
	    <input type="button" class="button" value="취소" onClick="self.close();">
	</td>
</tr>
</form>	
</table>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->