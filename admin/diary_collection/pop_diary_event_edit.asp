<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary_collection/diary_collection_cls.asp" -->

<%
dim YearUse
YearUse = request("YearUse")

dim bannerid
bannerid =request("bannerid")
dim objevt,intLoop
set objevt = new ClsDiary
objevt.getBannerOne bannerid



%>
<script language="javascript">

function subchk(){

	if(document.regfrm.bannerUrl.value.length<1){
		document.regfrm.bannerUrl.focus();
		alert('��ũ�� �Է��ϼž� �մϴ�.');
		return false;
	}
	if(document.regfrm.imagename.value.length<1){
		alert('�̹����� �Է��� �ּ���');
		return false;
	}
	document.regfrm.submit();
}

function showimage(img){
	var pop = window.open('viewImage.asp?imageUrl='+img,'imgview','width=600,height=600,resizable=yes');
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

function jsImgDel(divnm,iptNm,vPath){

	window.open('','imgdel','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imgdel';
	document.imginputfrm.action='diary_img_input.asp';
	document.imginputfrm.submit();
}

function jsSelevtType(v){

	var aa = document.getElementsByTagName("div");
	
	for (var i=0;i<aa.length;i++){
		if (aa[i].id==v||aa[i].id=='imgdiv'){
			aa[i].style.display='';
		}else {
			aa[i].style.display='none';
		}
	}
	if (v=='dzone'||v=='evtmain'){
		document.regfrm.mapusing.value='Y';
		fnmapusing('Y');
	} else {
		document.regfrm.mapusing.value='N';
		fnmapusing('N');
	}
}

function fnmapusing(v){
	var ipturl = document.getElementById('bannerUrl');
	var udiv = document.getElementById('urldiv');
	var tTxt = ipturl.value;
	
	if(v=='Y'){
		//udiv.innerHTML='<input type="text" name="bannerUrl" size="60" maxlength="100" value="">';
		udiv.innerHTML='<textarea name="bannerUrl" cols="60" rows="10" style="overflow=auto" wrap="hard">'+ tTxt + '</textarea>';
	}else{
		udiv.innerHTML='<input type="text" name="bannerUrl" size="60" maxlength="100" value="' + tTxt + '">';
	}
}

document.domain = "10x10.co.kr";
window.resizeTo(700,700);
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
        	<b>���̾ �̺�Ʈ ��ʵ��</b></td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- �ߴ� ���� -->
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#9d9d9d">
	<form name="regfrm" method="post" action="proc_diary_event.asp">
	<input type="hidden" name="mode" value="edit">
	<input type="hidden" name="YearUse" value="<%= YearUse %>">
	<input type="hidden" name="bannerid" value="<%= objevt.DiaryPrd.FBanneridx %>">
	<input type="hidden" name="mapusing" value="<%= objevt.DiaryPrd.FBannerMapUsing %>">
	<tr bgcolor="#FFFFFF">
		<td align="center" width="120" bgcolor="<%= adminColor("topbar") %>"><b>��ȣ</b></td>
		<td><%= objevt.DiaryPrd.FBanneridx %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>�����ġ</b></td>
		<td>
			<select name="bannerType" onchange="jsSelevtType(this.value)">
				<option value="multi" <% if objevt.DiaryPrd.FBannerType="multi" then response.write "selected" %> >���� ��Ƽ</option>
				<option value="left" <% if objevt.DiaryPrd.FBannerType="left" then response.write "selected" %>>���� �޴�</option>
				<option value="power" <% if objevt.DiaryPrd.FBannerType="power" then response.write "selected" %>>�Ŀ��̺�Ʈ</option>
				<option value="today" <% if objevt.DiaryPrd.FBannerType="today" then response.write "selected" %>>Today`s Diary</option>
				<option value="quiz" <% if objevt.DiaryPrd.FBannerType="quiz" then response.write "selected" %>>Diary Quiz</option>
				<option value="dzone" <% if objevt.DiaryPrd.FBannerType="dzone" then response.write "selected" %>>��������</option>
				<option value="tdayitem" <% if objevt.DiaryPrd.FBannerType="tdayitem" then response.write "selected" %>>Today`s Item</option>
				<option value="evtmain" <% if objevt.DiaryPrd.FBannerType="evtmain" then response.write "selected" %>>�̺�Ʈ���� ���</option>
				<option value="othermall_left" <% if objevt.DiaryPrd.FBannerType="othermall_left" then response.write "selected" %>>[�ܺθ�]�����޴�</option>
				<option value="othermall_multi" <% if objevt.DiaryPrd.FBannerType="othermall_multi" then response.write "selected" %>>[�ܺθ�]���θ�Ƽ</option>
				<option value="othermall_right" <% if objevt.DiaryPrd.FBannerType="othermall_right" then response.write "selected" %>>[�ܺθ�]�����޴�</option>
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>�̺�Ʈ �ڵ�</b></td>
		<td><input type="text" name="evtcode" size="10" maxlength="10" value="<%= objevt.DiaryPrd.FEvtCode %>" />(������ ����)</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>URL</b></td>
		<td><!--<input type="checkbox" name="mapusing" onclick="fnmapusing(this.checked);" <% if objevt.DiaryPrd.FbannerMapUsing="Y" then response.write "checked" %>>�̹����� ���<br>-->
			(http://www.10x10.co.kr)���� / �̹����� &lt;map name="design_zone"&gt; <br>
			<span id="urldiv">
			<textarea name="bannerUrl" cols="60" rows="10" style="overflow=auto" wrap="hard"><%= objevt.DiaryPrd.FBannerUrl %></textarea>
			</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>��� �̹���</b><br></td>
		<td>
			<!-- ���θ�Ƽ -->
			<div id="multi" style="display:none;">
				<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('imgdiv','imagename','eventbanner','400','600','false');"/>
				(<b><font color="red">Width 600</font></b>,<b><font color="red">JPG,GIF</font></b>������)
			</div>
			<!-- ������� -->
			<div id="left" style="display:none">
				<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('imgdiv','imagename','eventbanner','100','194','false');"/>
				(<b><font color="red">Width 194</font></b>,<b><font color="red">JPG,GIF</font></b>������)
			</div>
			<!-- �Ŀ��̺�Ʈ -->
			<div id="power" style="display:none">
				<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('imgdiv','imagename','eventbanner','100','230','false');"/>
				(<b><font color="red">Width 230</font></b>,<b><font color="red">JPG,GIF</font></b>������)
			</div>
			<!-- ������ ���̾ -->
			<div id="today" style="display:none">
				<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('imgdiv','imagename','eventbanner','100','192','false');"/>
				(<b><font color="red">192 X 124</font></b>,<b><font color="red">JPG,GIF</font></b>������)
			</div>
			<!-- ���� �̺�Ʈ -->
			<div id="quiz" style="display:none">
				<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('imgdiv','imagename','eventbanner','100','207','false');"/>
				(<b><font color="red">207 X 133</font></b>,<b><font color="red">JPG,GIF</font></b>������)
			</div>
			<!-- ���� ������ �� -->
			<div id="dzone" style="display:none">
				<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('imgdiv','imagename','eventbanner','100','285','false');"/>
				(<b><font color="red">285 X 133</font></b>,<b><font color="red">JPG,GIF</font></b>������)
			</div>
			<!-- ������ ������ -->
			<div id="tdayitem" style="display:none">
				<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('imgdiv','imagename','eventbanner','100','153','false');"/>
				(<b><font color="red">153 X 145</font></b>,<b><font color="red">JPG,GIF</font></b>������)
			</div>
			<!-- �̺�Ʈ ���� ��� -->
			<div id="evtmain" style="display:none">
				<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('imgdiv','imagename','eventbanner','400','750','false');"/>
				(<b><font color="red">width 750</font></b>,<b><font color="red">JPG,GIF</font></b>������)
			</div>
			<!-- ���޸� ���� -->
			<div id="othermall_left" style="display:none">
				<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('imgdiv','imagename','eventbanner','100','194','false');"/>
				(<b><font color="red">Width 194</font></b>,<b><font color="red">JPG,GIF</font></b>������)
			</div>
			<!-- ���޸� ��Ƽ -->
			<div id="othermall_multi" style="display:none">
				<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('imgdiv','imagename','eventbanner','400','600','false');"/>
				(<b><font color="red">Width 600</font></b>,<b><font color="red">JPG,GIF</font></b>������)
			</div>
			<!-- ���޸� ���� -->
			<div id="othermall_right" style="display:none">
				<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('imgdiv','imagename','eventbanner','121','153','false');"/>
				(<b><font color="red">Width 153 height 121</font></b>,<b><font color="red">JPG,GIF</font></b>������)
			</div>
			<input type="hidden" name="imagename" value="<%= objevt.DiaryPrd.FBannerImg%>">
			<div align="right" id="imgdiv"><img src="<%= objevt.DiaryPrd.getBannerImgUrl%>" width="50" height="50" border="0" style="cursor:pointer" onclick="showimage('<%= objevt.DiaryPrd.getBannerImgUrl %>');"></div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("topbar") %>"><b>��� ����</b></td>
		<td>
			<label><input type="radio" name="isusing" value="Y" <% if objevt.DiaryPrd.FBannerUsing="Y" then response.write "checked"%> /> ��� </label>
			<label><input type="radio" name="isusing" value="N" <% if objevt.DiaryPrd.FBannerUsing="N" then response.write "checked"%>/> ������ </label>
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
			<input type="button" class="button" value="���" onclick="history.go(-1);"/>
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
<input type="hidden" name="YearUse" value="<%= YearUse %>">
<input type="hidden" name="divName" value="">
<input type="hidden" name="orgImgName" value="">
<input type="hidden" name="inputname" value="">
<input type="hidden" name="ImagePath" value="">
<input type="hidden" name="maxFileSize" value="">
<input type="hidden" name="maxFileWidth" value="">
<input type="hidden" name="makeThumbYn" value="">
</form>
<script>jsSelevtType(document.regfrm.bannerType.value);fnmapusing(document.regfrm.mapusing.value);

</script>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->
