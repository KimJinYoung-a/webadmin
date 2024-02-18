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
' Discription : 동영상 관리 목록
' History : 2009.09.29 허진원 : 생성
'           2013.08.23 허진원; jwplayer6.6 업그레이드
'           2022.02.08 허진원; copy script 변경
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
// 페이지 이동
function goPage(pg)
{
	document.refreshFrm.page.value=pg;
	document.refreshFrm.action="videoList.asp";
	document.refreshFrm.submit();
}

// 동영상 소스 생성(html5)
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

	alert('선택하신 동영상의 소스가 복사되었습니다. 사용하실 곳에 Ctrl+V 하시면됩니다.');
}

// 동영상 소스 생성
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

	var doc = "<div id='player"+vNo+"'>로딩 중</div>";
		doc += String.fromCharCode(60) + "script type='text/javascript'>";
		doc += "jwplayer('player"+vNo+"').setup({";
		doc += "	width:"+vWd+", height:"+vHt+",";
		doc += "	file: '"+vFn+"',";
		doc += "	image: '"+vTh+"',";
		doc += "	abouttext: '텐바이텐 10X10',";
		doc += "	aboutlink: 'http://www.10x10.co.kr'";
		doc += "});";
		doc += "<\/script>";
	copyStringToClipboard(doc);
	alert('선택하신 동영상의 소스가 복사되었습니다. 사용하실 곳에 Ctrl+V 하시면됩니다.');
}

// 모바일용 동영상 소스 생성
function copySrcM(vFn,vTh,vWd,vHt) {
	var doc = "<video poster='"+vTh+"' src='"+vFn+"' controls='true'></video>"
	copyStringToClipboard(doc.replace(/\'/gi,String.fromCharCode(34)));
	alert('선택하신 동영상의 소스가 복사되었습니다. 사용하실 곳에 Ctrl+V 하시면됩니다.');
}

// 동영상 팝업 생성
function copyPopup(vSn) {
	var doc = "<%=www2009url%>/common/popFLVPlayer.asp?vSn=" + vSn;
	copyStringToClipboard(doc);
	alert('선택하신 동영상의 팝업 페이지가 복사되었습니다. 사용하실 곳에 Ctrl+V 하시면됩니다.');
}
//-->
</script>
<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get" onSubmit="frm_search()" action="videoList.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		동영상 구분
		<%=drawVDivSelect("div",div)%>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td align="right"><input type="button" value="동영상 추가" onclick="self.location='videoWrite.asp?mode=add&menupos=<%= menupos %>'" class="button"></td>
</tr>
</table>
<!-- 액션 끝 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		검색결과 : <b><%=oVideo.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oVideo.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>번호</td>
	<td>구분</td>
	<td>썸네일</td>
	<td>제목</td>
	<td>너비</td>
	<td>높이</td>
	<td>등록일</td>
	<td>&nbsp;</td>
</tr>
<%	if oVideo.FResultCount < 1 then %>
<tr>
	<td colspan="8" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 동영상이 없습니다.</td>
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
		<input type="button" class="button" value="소스복사" onClick="copySrcNew('<%= oVideo.FItemList(i).FvideoSn %>','<%= webImgUrl&"/video/"&oVideo.FItemList(i).FvideoFile %>','<%= webImgUrl&"/video/"&oVideo.FItemList(i).FvideoThumb %>',<%= oVideo.FItemList(i).FvideoWidth %>,<%= oVideo.FItemList(i).FvideoHeight+20 %>)">
	</td>
</tr>
<%
		next
	end if
%>
<!-- 메인 목록 끝 -->
<tr bgcolor="#FFFFFF">
	<td colspan="8" align="center">
	<!-- 페이지 시작 -->
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
	<!-- 페이지 끝 -->
	</td>
</tr>
</table>
<%
set oVideo = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->