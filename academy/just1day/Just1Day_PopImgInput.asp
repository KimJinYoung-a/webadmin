<%@ language=vbscript %>
<% option explicit %>
<% Response.ChaRset = "EUC-KR" %>
<%
'###########################################################
' Description : �ΰŽ� ����Ʈ ������ �̹��� ���ε�
' Hieditor : 2016.08.01 ���¿� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->

<%
dim orgImgName ,newImgName ,ImagePath ,maxFileSize ,maxFileWidth ,makeThumbYn,divName,inputname

orgImgName 		= RequestCheckvar(request("orgImgName"),32)
ImagePath 		= request("ImagePath")
maxFileSize 	= RequestCheckvar(request("maxFileSize"),10)
maxFileWidth	= RequestCheckvar(request("maxFileWidth"),10)
makeThumbYn 	= RequestCheckvar(request("makeThumbYn"),10)
divName = RequestCheckvar(request("divName"),32)
inputname = RequestCheckvar(request("inputname"),32)

newImgName 		= year(now())&month(now())&day(now())&hour(now())&minute(now())&second(now())

if ImagePath <> "" then
	if checkNotValidHTML(ImagePath) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if

%>
<script language="javascript">

function subchk(){
	if (document.regfrm.sfImg.value.length<1){
		alert('�̹����� ������ �ּ���');
		return false;
	}
	document.regfrm.submit();
}
window.resizeTo(400,250);
</script>
<!-- ��� �޴� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top" style="padding : 0 0 10 0">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td align="center">
        	<b>�̹�������</b></td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#9d9d9d">
	<form name="regfrm" method="post" target="prcframe" action="<%=uploadImgUrl%>/linkweb/academy/just1day_image.asp" enctype="multipart/form-data">	
	<input type="hidden" name="divName" value="<%= divName %>">
	<input type="hidden" name="orgImgName" value="<%= orgImgName %>">
	<input type="hidden" name="newImgName" value="<%= newImgName %>">
	<input type="hidden" name="inputname" value="<%= inputname %>">

	<input type="hidden" name="ImagePath" value="<%= ImagePath %>">
	<input type="hidden" name="maxFileSize" value="<%= maxFileSize %>">
	<input type="hidden" name="maxFileWidth" value="<%= maxFileWidth %>">
	<input type="hidden" name="makeThumbYn" value="<%= makeThumbYn %>">

	<tr bgcolor="#FFFFFF">
		<td>
			<input type="file" name="sfImg" size="35" value="">
			<br>
			�̹����� �����Ͻ� ��� ���� �̹����� ���� �˴ϴ�.
		</td>
	</tr>

	</form>
</table>
<!-- �ϴ� ����¡ ���� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
			<input type="button" class="button" value="����" onclick="subchk();"/>&nbsp;&nbsp;
			<input type="button" class="button" value="���" onclick="self.close();"/>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
</body>
</html>
<iframe name="prcframe" src="" frameborder="0" width="600" height="600"></iframe>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->