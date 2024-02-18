<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<script language="javascript">
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
	
	frm.submit();
}

</script>

<form name="frmFile" method="post" action="<%=uploadUrl%>/linkweb/offshop/workschedule/upload_work_schedule_proc.asp"  enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="reguserid" value="<%=session("ssBctId")%>">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#999999">
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2">
		<b>오프라인직원 스케줄 엑셀등록</b>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td width="60">샘플</td>
	<td align="left"><a href="/common/offshop/staff/schedule.xls" onfocus="this.blur()">schedule.xls</a></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>주의사항</td>
	<td align="left">
	* 맨 상단의 년, 월 등이 있는 <b>1줄은 그대로 두세요.</b><br>
	* 이름은 잘못 입력해도 무관하나 <b>사번은 절대 틀리면 안됩니다.</b> 이름은 따로 저장을 하지않지만 <b>사번은 저장하여 모든 데이터를 사번으로 조회되기에 주의해서 입력</b>하기 바랍니다.<br>
	* 1 ~ 31 칸에는 해당일에 해당된 업무코드를 넣으시면 됩니다. <b>지정된 업무코드 외에 입력시 에러가 납니다.</b><br>
	* 년, 월에 해당되는 날짜는 반드시 <b>달력에 있는 그대로의 날 수</b> 만큼 입력하세요.
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

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->