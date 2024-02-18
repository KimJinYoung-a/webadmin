<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description : AGV 상품 엑셀 일괄 업로드
' History : 2020.05.22 한용민 생성
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

<style>
html, body { margin: 0; padding: 0; }
</style>

<p />

<form name="frmFile" method="post" action="<%= uploadUrl %>/linkweb/item/agv/upload_AGV_stock_invest_item_excel.asp"  enctype="MULTIPART/FORM-DATA" style="margin:0px;">

<table width="98%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#999999">
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2">
		<b>AGV 상품 엑셀 일괄 업로드</b>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td width="60">샘플</td>
	<td align="left"><a href="<%= uploadUrl %>/offshop/sample/item/logics_agv_stock_invest_sample_v1.xls" target="_blank">logics_agv_stock_invest_sample_v1.xls</a></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td><font color="red"><b>주의사항</b></font></td>
	<td align="left">
		※ 엑셀에서 저장형식이 <font color="red"><b>Save As Excel 97 -2003 통합문서</b></font> 형태만 인식 합니다.
		<br><br>* 위에 샘플을 다운 받으신후 등록하시면 됩니다.
		<br><br>* 맨 상단의 물류코드가 있는 <font color="red"><b>1줄은 그대로 두세요.</b></font>
		<br><br>* 물류코드를 기준으로 업데이트 되기 때문에 <font color="red"><b>물류코드는 절대 공란이거나, 틀리면 안됩니다.</b></font>
		<br>물류코드 중간에 다시기호(-) 는 생략 하셔도 등록이 됩니다. 예)10-01239286-0000 -> 10012392860000
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

<p />

<iframe id="view" name="view" src="" width="98%" height=30 frameborder="0" scrolling="no" align="center" style="display:block;"></iframe>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
