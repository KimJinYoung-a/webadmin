<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary_collection/diary_collection_cls.asp" -->

<%
dim idx
idx = request("idx")
dim objDiary
set objDiary = new ClsDiary
objDiary.getDiaryItem idx


%>

<script language="javascript">

function subchk(){

	if(document.regfrm.itemid.value.length<1){
		document.regfrm.itemid.focus();
		alert('��ǰ ��ȣ�� �Է��ϼž� �մϴ�.');
		return false;
	}
	if(document.regfrm.basicimgName.value.length<1){
		alert('�̹����� �Է��� �ּ���');
		return false;
	}
	document.regfrm.submit();
}


function jsImgInput(divnm,iptNm,vPath,Fsize,Fwidth,thumb){


	window.open('','imginput','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imginput';
	document.imginputfrm.action='diary_img_input.asp';
	document.imginputfrm.submit();
}

function showimage(img){
	var pop = window.open('viewImage.asp?imageUrl='+img,'imgview','width=600,height=600,scrollbars=yes,resizable=yes');
}
document.domain = "10x10.co.kr";
window.resizeTo(500,500)
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
        	<b>���̾ �������� ��� </b></td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#9d9d9d">
	<form name="regfrm" method="post" action="proc_diary_reg.asp">
	<input type="hidden" name="mode" value="edit">
	<input type="hidden" name="idx" value="<%= objDiary.DiaryPrd.FIdx %>">
	<tr bgcolor="#FFFFFF">
		<td align="center" width="100" bgcolor="<%= adminColor("topbar") %>"><b>����</b></td>
		<td>
			<select name="diaryType">
				<option value="illust" <% if objDiary.DiaryPrd.FDiaryType="illust" then response.write "selected" %>>�Ϸ���Ʈ</option>
				<option value="photo"  <% if objDiary.DiaryPrd.FDiaryType="photo" then response.write "selected" %>>����/��ȭ</option>
				<option value="system" <% if objDiary.DiaryPrd.FDiaryType="system" then response.write "selected" %>>�ý���</option>
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>��ǰ��ȣ</b></td>
		<td><input type="text" name="itemid" maxlength="10" value="<%= objDiary.DiaryPrd.FItemid %>" /></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>�⺻ �̹���</b><br></td>
		<td>
			<label><input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('imgdiv','basicimgName','basic','200','400','true');"/></label>
			(<b><font color="red">400x400</font></b>,<b><font color="red">JPG</font></b>������)
			<input type="hidden" name="basicimgName" value="<%= objDiary.DiaryPrd.FBasicImg %>">
			<div align="right" id="imgdiv"><img src="<%= objDiary.DiaryPrd.getBasicImgUrl %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= objDiary.DiaryPrd.getBasicImgUrl %>');"></div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>��Ÿ</b></td>
		<td>
			<label><input type="checkbox" name="hitYn" <% if objDiary.DiaryPrd.FhitYn="Y" then response.write "checked" %> />HIT</label><font color="orange" size="3"> | </font>
			<label><input type="checkbox" name="giftYn" <% if objDiary.DiaryPrd.FgiftYn="Y" then response.write "checked" %> />Gift</label><font color="orange" size="3"> | </font>
			<label><input type="checkbox" name="onlyyearYn" <% if objDiary.DiaryPrd.FonlyYearYn="Y" then response.write "checked" %> />2007�� ���� </label><font color="orange" size="3"> | </font>
			<label><input type="checkbox" name="freeBaeSongYn" disabled <% if objDiary.DiaryPrd.Fitemcoupontype="3" then response.write "checked" %> onclick="return false;"/>���� ��� </label>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>��� ����</b></td>
		<td>
		<% if objDiary.DiaryPrd.Fisusing="Y" then %>
			<label><input type="radio" name="isusing" value="Y" checked /> ��� </label><label><input type="radio" name="isusing" value="N" /> ������ </label>
		<% else %>
			<label><input type="radio" name="isusing" value="Y" /> ��� </label><label><input type="radio" name="isusing" value="N" checked /> ������ </label>
		<% end if %>
		</td>
	</tr>
	</form>
</table>
<!-- �ϴ�  ���� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
			<input type="button" class="button" value="Ȯ��" onclick="subchk();"/>&nbsp;&nbsp;
			<input type="button" class="button" value="���" onclick="window.close();"/>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<form name="imginputfrm" method="post" action="">
<input type="hidden" name="YearUse" value="<%= objDiary.DiaryPrd.Fyear %>">
<input type="hidden" name="divName" value="">
<input type="hidden" name="orgImgName" value="">
<input type="hidden" name="inputname" value="">
<input type="hidden" name="ImagePath" value="">
<input type="hidden" name="maxFileSize" value="">
<input type="hidden" name="maxFileWidth" value="">
<input type="hidden" name="makeThumbYn" value="">
</form>
</body>
</html>
<% set objDiary = nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->