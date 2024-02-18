<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/videoInfoCls.asp"-->
<%
'###############################################
' PageName : videoList.asp
' Discription : ������ ���� ���
' History : 2009.09.29 ������ : ����
'           2013.08.23 ������; jwplayer6.6 ���׷��̵�
'           2022.02.08 ������; copy script ����
'###############################################

dim page, div, i, lp

page = request("page")
if page = "" then page=1
div = request("div")

dim oVideo
set oVideo = New CVideo
oVideo.FCurrPage = page
oVideo.FPageSize=20
oVideo.FRectDiv = div
oVideo.FRectUsing = "Y"
oVideo.GetVideoList

%>
<script type="text/javascript">
<!--
// ������ �̵�
function goPage(pg)
{
	document.refreshFrm.page.value=pg;
	document.refreshFrm.action="videoList.asp";
	document.refreshFrm.submit();
}

// ������ �ҽ� ����(html5)
function copySrcNew(vNo,vFn,vTh,vWd,vHt) {
	var doc = "<video preload=\"auto\" autoplay=\"true\" loop=\"loop\" muted=\"muted\" volume=\"0\" style=\"";
		if(vWd=="0"){
			doc += "width:100%;"
		} else {
			doc += "width:"+vWd+"px;"
		}
		if(vWd!="0"){
			doc += "height:"+vHt+"px;"
		}
		doc += "\" playsinline>\n"
		doc += "    <source src=\""+vFn+"\" type=\"video/mp4\" />\n";
		doc += "    <img src=\""+vTh+"\" alt=\"\" />\n";
		doc += "</video>";
	const t = document.createElement("textarea");
	document.body.appendChild(t);
	t.value = doc;
	t.select();
	document.execCommand('copy');
	document.body.removeChild(t);

	alert('�����Ͻ� �������� �ҽ��� ����Ǿ����ϴ�. ����Ͻ� ���� Ctrl+V �Ͻø�˴ϴ�.');
}

// ������ �ҽ� ����
function copySrc(vNo,vFn,vTh,vWd,vHt) {
	//var doc = String.fromCharCode(60) + "script language='javascript'" + String.fromCharCode(62);
	//	doc += "if ((navigator.userAgent.indexOf('iPhone') != -1)||(navigator.userAgent.indexOf('iPod') != -1)||(navigator.userAgent.indexOf('iPad') != -1)) {";
	//	doc += "	document.write(\"<video width=\'"+vWd+"\' height=\'"+vHt+"\' poster=\'"+vTh+"\' src=\'"+vFn+"\' controls=\'true\' type=\'video/mp4\'></video>\");";
	//	doc += "} else{";
	//	doc += "	document.write(\"<object classid=\'clsid:d27cdb6e-ae6d-11cf-96b8-444553540000\' codebase=\'http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0\' width=\'"+vWd+"\' height=\'"+(vHt+20)+"\' align=\'middle\'>\");";
	//	doc += "	document.write(\"<param name=\'allowScriptAccess\' value=\'always\'>\");";
	//	doc += "	document.write(\"<param name=\'movie\' value=\'http://fiximage.10x10.co.kr/flash/flvplayer.swf?file="+vFn+"&image="+vTh+"\'>\");";
	//	doc += "	document.write(\"<param name=\'menu\' value=\'false\'>\");";
	//	doc += "	document.write(\"<param name=\'quality\' value=\'high\'>\");";
	//	doc += "	document.write(\"<param name=\'wmode\' value=\'transparent\'>\");";
	//	doc += "	document.write(\"<embed src=\'http://fiximage.10x10.co.kr/flash/flvplayer.swf?file="+vFn+"&image="+vTh+"\' menu=\'false\' quality=\'high\' wmode=\'transparent\' width=\'"+vWd+"\' height=\'"+(vHt+20)+"\' align=\'middle\' allowScriptAccess=\'always\' allowfullscreen=\'true\' allownetworking=\'all\' type=\'application/x-shockwave-flash\' pluginspage=\'http://www.macromedia.com/go/getflashplayer\' />\");";
	//	doc += "	document.write(\"</object>\");";
	//	doc += "}";
	//	doc += "<\/script>";

	var doc = "<div id='player"+vNo+"'>�ε� ��</div>";
		doc += String.fromCharCode(60) + "script type='text/javascript'>";
		doc += "jwplayer('player"+vNo+"').setup({";
		doc += "	width:"+vWd+", height:"+vHt+",";
		doc += "	file: '"+vFn+"',";
		doc += "	image: '"+vTh+"',";
		doc += "	abouttext: '�ٹ����� 10X10',";
		doc += "	aboutlink: 'http://www.10x10.co.kr'";
		doc += "});";
		doc += "<\/script>";
	copyStringToClipboard(doc);
	alert('�����Ͻ� �������� �ҽ��� ����Ǿ����ϴ�. ����Ͻ� ���� Ctrl+V �Ͻø�˴ϴ�.');
}

// ����Ͽ� ������ �ҽ� ����
function copySrcM(vFn,vTh,vWd,vHt) {
	var doc = "<video poster='"+vTh+"' src='"+vFn+"' controls='true'></video>"
	copyStringToClipboard(doc.replace(/\'/gi,String.fromCharCode(34)));
	alert('�����Ͻ� �������� �ҽ��� ����Ǿ����ϴ�. ����Ͻ� ���� Ctrl+V �Ͻø�˴ϴ�.');
}

// ������ �˾� ����
function copyPopup(vSn) {
	var doc = "<%=www2009url%>/common/popFLVPlayer.asp?vSn=" + vSn;
	copyStringToClipboard(doc);
	alert('�����Ͻ� �������� �˾� �������� ����Ǿ����ϴ�. ����Ͻ� ���� Ctrl+V �Ͻø�˴ϴ�.');
}
//-->
</script>
<!-- ��� �˻��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get" onSubmit="frm_search()" action="videoList.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
		������ ����
		<%=drawVDivSelect("div",div)%>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="�˻�">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td align="right"><input type="button" value="������ �߰�" onclick="self.location='videoWrite.asp?mode=add&menupos=<%= menupos %>'" class="button"></td>
</tr>
</table>
<!-- �׼� �� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		�˻���� : <b><%=oVideo.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oVideo.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��ȣ</td>
	<td>����</td>
	<td>�����</td>
	<td>����</td>
	<td>�ʺ�</td>
	<td>����</td>
	<td>�����</td>
	<td>&nbsp;</td>
</tr>
<%	if oVideo.FResultCount < 1 then %>
<tr>
	<td colspan="8" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� �������� �����ϴ�.</td>
</tr>
<%
	else
		for i=0 to oVideo.FResultCount-1
%>
<tr bgcolor="#FFFFFF">
	<td align="center"><a href="videoWrite.asp?mode=edit&menupos=<%= menupos %>&videoSn=<%= oVideo.FItemList(i).FvideoSn %>"><%= oVideo.FItemList(i).FvideoSn %></a></td>
	<td align="center"><%= getVDivName(oVideo.FItemList(i).FvideoDiv) %></td>
	<td align="center"><a href="videoWrite.asp?mode=edit&menupos=<%= menupos %>&videoSn=<%= oVideo.FItemList(i).FvideoSn %>"><img src="<%= webImgUrl & "/video/" & oVideo.FItemList(i).FvideoThumb %>" width="100" border="0"></a></td>
	<td align="center"><a href="videoWrite.asp?mode=edit&menupos=<%= menupos %>&videoSn=<%= oVideo.FItemList(i).FvideoSn %>"><%= oVideo.FItemList(i).FvideoTitle %></a></td>
	<td align="center"><%= oVideo.FItemList(i).FvideoWidth %>px</td>
	<td align="center"><%= oVideo.FItemList(i).FvideoHeight %>px</td>
	<td align="center"><%= left(oVideo.FItemList(i).Fregdate,10) %></td>
	<td align="center">
		<input type="button" class="button" value="�ҽ�����" onClick="copySrcNew('<%= oVideo.FItemList(i).FvideoSn %>','<%= webImgUrl&"/video/"&oVideo.FItemList(i).FvideoFile %>','<%= webImgUrl&"/video/"&oVideo.FItemList(i).FvideoThumb %>',<%= oVideo.FItemList(i).FvideoWidth %>,<%= oVideo.FItemList(i).FvideoHeight+20 %>)">
	</td>
</tr>
<%
		next
	end if
%>
<!-- ���� ��� �� -->
<tr bgcolor="#FFFFFF">
	<td colspan="8" align="center">
	<!-- ������ ���� -->
	<%
		if oVideo.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oVideo.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oVideo.StartScrollPage to oVideo.FScrollCount + oVideo.StartScrollPage - 1

			if lp>oVideo.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>" & lp & "</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>" & lp & "</a> "
			end if

		next

		if oVideo.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- ������ �� -->
	</td>
</tr>
</table>
<%
set oVideo = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->