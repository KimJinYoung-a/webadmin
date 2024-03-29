<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : 푸시 직접 타게팅
' Hieditor : 2018.08.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/push/apppush_msg_cls.asp" -->
<%
Dim idx, iMaxLength, mode
	idx = requestCheckVar(request("idx"),32)
    mode = requestCheckVar(request("mode"),32)

IF iMaxLength = "" THEN iMaxLength = 1
%>

<script type="text/javascript">

function fnChkFile(sFile, sMaxSize, arrExt){
    //파일 업로드 유무확인
     if (!sFile){
    	 return true;
    	}

    var blnResult = false;

    //파일 용량 확인
    var maxsize = sMaxSize * 1024 * 1024;

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

function frmSumbit(){
	arrFileExt = new Array();
	arrFileExt[arrFileExt.length]  = "csv";

	//파일유효성 체크
	if (!fnChkFile(frmFile.sFile.value, <%=iMaxLength%>, arrFileExt)){
		alert("파일은 <%=iMaxLength%>MB이하의 csv파일만 업로드 가능합니다.");
		return;
	}
	if (frmFile.sFile.value==''){
		alert("파일을 선택해 주세요.");
		return;
	}

	frmFile.submit();
}

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td align="center" width=100>
		CSV 업로드샘플
	</td>
	<td align="left">
        <font color="red">* 3만명씩 짤라서 입력해 주세요.</font>
        <br><br>
        텐바이텐고객번호 = 고객번호(앰플리튜트,앱보이) / 3
        <br>
        <a href="https://imgstatic.10x10.co.kr/offshop/sample/push/푸시_텐바이텐고객번호_샘플.csv" onfocus="this.blur();">
		<font color="red">푸시_텐바이텐고객번호_샘플.csv</font></a>
        <br><br>
        텐바이텐고객아이디
        <br>
        <a href="https://imgstatic.10x10.co.kr/offshop/sample/push/푸시_텐바이텐고객아이디_샘플.csv" onfocus="this.blur();">
		<font color="red">푸시_텐바이텐고객아이디_샘플.csv</font></a>
	</td>
</tr>
</table>
<!-- 액션 끝 -->
<Br>
<form name="frmFile" method="post" action="/admin/appmanage/push/msg/doPushmsgfile_proc.asp" enctype="MULTIPART/FORM-DATA" style="margin:0px;">
<input type="hidden" name="iML" value="<%=iMaxLength%>">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="idx" value="<%= idx %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">파일명:</td>
	<td align="left">
		<input type="file" name="sFile" class="file">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" colspan="2">
	    <input type="button" class="button" value="등록" onClick="frmSumbit();">
	    <input type="button" class="button" value="취소" onClick="self.close();">
	</td>
</tr>
</form>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
session.codePage = 949
%>