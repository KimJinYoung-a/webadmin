<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->

<%
	Dim vIdx, cDiary, vMakerId, vCateCode, vListtitleimgName, vListmainimgName, vListspareimgName, vIsUsing, vListText, vContentHtml, vSorting, vRegdate, vContentTitleName
	vIdx = request("idx")
	If vIdx <> "" Then
		set cDiary = new DiaryCls
		cDiary.Fidx = vIdx
		cDiary.getBrandInterviewDetail
		
		vMakerId			= cDiary.FItem.fmakerid
		vCateCode			= cDiary.FItem.FCateCode
		vListtitleimgName	= cDiary.FItem.FImage1
		vListmainimgName	= cDiary.FItem.FImage2
		vListspareimgName	= cDiary.FItem.FImg
		vIsUsing			= cDiary.FItem.FisUsing
		vListText			= cDiary.FItem.Fexplain
		vContentTitleName	= cDiary.FItem.ConfImg
		vContentHtml		= cDiary.FItem.ConTTxt
		vSorting			= cDiary.FItem.Fsorting
		vRegdate			= cDiary.FItem.FRegdate
		
		set cDiary = Nothing
	Else
		vSorting = "1"
		vIsUsing = "N"
	End If
%>

<script type="text/javascript" src="http://www.10x10.co.kr/lib/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
<!--
function CtgBestRefresh() {
	$("#branddiv").show();
	
	var str = $.ajax({
		type: "GET",
		url: "/admin/Diary2009/Lib/brand_interview_brandlist.asp",
		data: "",
		dataType: "text",
		async: false
	}).responseText;
	$("#branddiv").html(str);
}

function selectMakerid(m)
{
	document.frmreg.makerid.value = m;
	document.getElementById("branddiv").style.display = "none";
}
	
function showimage(img){
	var pop = window.open('/lib/showimage.asp?img='+img,'imgview','width=600,height=600,resizable=yes');
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
	document.imginputfrm.action='PopImgInput.asp';
	document.imginputfrm.submit();
}

function delimage(gubun)
{
	var aa = eval("document.frmreg."+gubun+"");
	aa.value = "";
	frmreg.submit();
}

function goSubmit()
{
	if(document.frmreg.makerid.value == "")
	{
		alert("�귣�带 �����ϼ���.");
		return;
	}
	
	document.frmreg.submit();
}

document.domain = "10x10.co.kr";
-->
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
<form name="frmreg" method="post" action="/admin/Diary2009/brand_interview_proc.asp">
<input type="hidden" name="idx" value="<%= vIdx %>">
<tr>
	<td>
		<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF" height="25">
			<td colspan="2" align="center"><%=CHKIIF(vIdx<>"","No."&vIdx&" ","")%> <b>�귣�� ���ͺ� ����</b><%=CHKIIF(vIdx<>"","&nbsp;&nbsp;&nbsp;�����:"&vRegdate,"")%></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap> �귣��(�귣����̵�)</td>
			<td bgcolor="#FFFFFF" align="left"><input type="text" class="text" name="makerid" value="<%=vMakerId%>" size="20">
			<div id="branddiv" style="background-color:white; border-width:1px; border-style:solid; padding-right:20px; position:absolute; z-index:1; height:300px; display:none; overflow-y:scroll;"></div>
			<input type="button" class="button" value="ã��" onClick="CtgBestRefresh();">
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap> ����</td>
			<td bgcolor="#FFFFFF" align="left"><% SelectList "cate", vCateCode %></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150"> ����Ʈ Ÿ��Ʋ �̹��� </td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('imgdiv22','listtitleimgName','listtitleimg','2000','750','false');"/>
				<input type="hidden" name="listtitleimgName" value="<%= vListtitleimgName %>">
				<div align="right" id="imgdiv22"><% IF vListtitleimgName<>"" THEN %><img src="<%= vListtitleimgName %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= vListtitleimgName %>');"><a href="javascript:delimage('listtitleimgName');">[����]</a><% End IF %></div>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150"> ����Ʈ ū �̹���</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('imgdiv','listmainimgName','listmainimg','2000','750','false');"/>
				(<b><font color="red">440x420</font></b>,<b><font color="red">JPG,GIF</font></b>������)
					<input type="hidden" name="listmainimgName" value="<%= vListmainimgName %>">
					<div align="right" id="imgdiv"><% IF vListmainimgName<>"" THEN %><img src="<%= vListmainimgName %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= vListmainimgName %>');"><a href="javascript:delimage('listmainimgName');">[����]</a><% End IF %></div>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap> ����Ʈ ���� �����</td>
			<td bgcolor="#FFFFFF" align="left"><textarea name="list_text" cols="70" rows="7"><%=vListText%></textarea></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150"> ����Ʈ ���� �̹���</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('imgdiv3','listspareimgName','listspareimg','2000','750','false');"/>
				(��ǰ�� 1���� ��츸 ������ �̹���)
				<input type="hidden" name="listspareimgName" value="<%= vListspareimgName %>">
				<div align="right" id="imgdiv3"><% IF vListspareimgName<>"" THEN %><img src="<%= vListspareimgName %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= vListspareimgName %>');"><a href="javascript:delimage('listspareimgName');">[����]</a><% End IF %></div>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap> ����Ʈ ���� ����</td>
			<td bgcolor="#FFFFFF" align="left"><input type="text" name="sorting" value="<%=vSorting%>" size="5">(���� ū���ڰ� ����, �⺻ 1)</td>
		</tr>
		
		
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap width="150"> ����� Ÿ��Ʋ �̹���</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="button" class="button" size="30" value="�̹��� �ֱ�" onclick="jsImgInput('imgdiv33','contenttitleName','contenttitle','2000','750','false');"/>
				(��ǰ�� 1���� ��츸 ������ �̹���)
				<input type="hidden" name="contenttitleName" value="<%= vContentTitleName %>">
				<div align="right" id="imgdiv33"><% IF vContentTitleName<>"" THEN %><img src="<%= vContentTitleName %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= vContentTitleName %>');"><a href="javascript:delimage('contenttitleName');">[����]</a><% End IF %></div>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap> ����� Html</td>
			<td bgcolor="#FFFFFF" align="left"><textarea name="content_html" cols="80" rows="16"><%=vContentHtml%></textarea></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td nowrap> ��뿩��</td>
			<td bgcolor="#FFFFFF" align="left">
				<input type="radio" name="isusing" value="Y" <% IF vIsUsing="Y" THEN %>checked<% END IF %>>���
				<input type="radio" name="isusing" value="N" <% IF vIsUsing="N" THEN %>checked<% END IF %> >������
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td colspan="2" align="center"><br>
		<img src="http://webadmin.10x10.co.kr/images/icon_save.gif" border="0" onClick="goSubmit();" style="cursor:pointer">
		<img src="http://webadmin.10x10.co.kr/images/icon_cancel.gif" border="0" onClick="window.close();" style="cursor:pointer">
		<img src="http://testwebadmin.10x10.co.kr/images/icon_new_registration.gif" border="0" onClick="location.href='/admin/diary2009/brand_interview_detail.asp';" style="cursor:pointer">
	</td>
</tr>
</form>
</table>

<form name="imginputfrm" method="post" action="">
<input type="hidden" name="divName" value="">
<input type="hidden" name="orgImgName" value="">
<input type="hidden" name="inputname" value="">
<input type="hidden" name="ImagePath" value="">
<input type="hidden" name="maxFileSize" value="">
<input type="hidden" name="maxFileWidth" value="">
<input type="hidden" name="makeThumbYn" value="">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->