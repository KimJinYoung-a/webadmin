<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  오프샵 아이템 엑셀 업로드 일괄 수정
' History : 2018.09.28 정태훈 생성
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<script type="text/javascript">

document.domain = "10x10.co.kr";

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

	frm.target='view';
	frm.submit();
}

//이미지 서버에서 타고 들어옴
function openerreload(){
	opener.location.reload();
	self.close();
}

</script>

<form name="frmFile" method="post" action="<%= uploadUrl %>/linkweb/offshop/upload_shopitemlist_excel.asp"  enctype="MULTIPART/FORM-DATA" style="margin:0px;">

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#999999">
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2">
		<b>오프라인 상품 엑셀 일괄 업로드 수정</b>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td width="60">샘플</td>
	<td align="left"><a href="<%= uploadUrl %>/offshop/sample/item/shop_item_list_sample_v1.xls" target="_blank">shop_item_list_sample.xls</a></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td><font color="red"><b>주의사항</b></font></td>
	<td align="left">
		※ 엑셀에서 저장형식이 <font color="red"><b>Save As Excel 97 -2003 통합문서</b></font> 형태만 인식 합니다.
		<br><br>* 위에 샘플을 다운 받으신후, 입력 하시거나, 리스트에 있는 <font color="red"><b>엑셀다운로드</b></font>를 받으셔서 등록하시면 됩니다.
		<br><br>* 맨 상단의 브랜드ID, 상품코드, 상품명 등이 있는 <font color="red"><b>1줄은 그대로 두세요.</b></font>
		<br><br>* 상품코드를 기준으로 업데이트 되기 때문에 <font color="red"><b>상품코드는 절대 공란이거나, 틀리면 안됩니다.</b></font>
		<br><br><font color="red"><b>* 소비자가, 판매가, 매입가, 매장공급가, 센터매입구분, 범용바코드, 사용여부, ON/OFF 가격연동</b></font>필드를 입력하시면 되며, 그대로 저장 됩니다.
	</td>
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
</table>

</form>

<iframe id="view" name="view" src="" width=1280 height=30 frameborder="0" scrolling="no"></iframe>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->