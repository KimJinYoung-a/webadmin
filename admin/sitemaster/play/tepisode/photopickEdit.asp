<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : T-Episode
' Hieditor : 이종화 생성
'			 2022.07.07 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/play/playCls.asp" -->
<%
Dim idx, viewtitle, subtitle, isusing, PPimg, regdate, playcate, menupos, style_html_m
Dim oPick , oground
	idx = requestCheckVar(getNumeric(request("idx")),10)
	menupos = requestCheckVar(getNumeric(request("menupos")),10)
playcate = 7 't-episode

Set oPick = new CPlayContents
	oPick.FRectIdx = idx

If idx <> "" Then
	oPick.GetPhotoPickOne()
	If oPick.FResultCount > 0 Then
		idx			= oPick.FOneItem.Fidx
		viewtitle	= oPick.FOneItem.FViewtitle
		subtitle	= oPick.FOneItem.FSubtitle
		isusing		= oPick.FOneItem.FIsusing
		PPimg		= oPick.FOneItem.FPPimg
		regdate		= oPick.FOneItem.FRegdate
		style_html_m =  oPick.FOneItem.fstyle_html_m
	End If
End If
set oPick = Nothing

If isusing = "" Then isusing = "Y"

%>
<script type="text/javascript">
<!--
	//이미지 확대화면 새창으로 보여주기
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('/admin/sitemaster/play/lib/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}

	function jsDelImg(sName, sSpan){
		if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
		   eval("document.all."+sName).value = "";
		   eval("document.all."+sSpan).style.display = "none";
		}
	}

	function jsSetImg(sImg, sName, sSpan){
		document.domain ="10x10.co.kr";

		var winImg;
		winImg = window.open('/admin/sitemaster/play/lib/pop_theme_uploadimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	function jsTagview(gidx , idx){	
		var poptag;
		poptag = window.open('/admin/sitemaster/play/lib/pop_tagReg.asp?idx='+gidx+'&playcate='+<%=playcate%>,'poptag','width=500,height=400,scrollbars=yes,resizable=yes');
		poptag.focus();
	}

	function subcheck(){
		var frm=document.inputfrm;

		if (!frm.viewtitle.value){
			alert('viewtitle를 등록해주세요');
			frm.viewtitle.focus();
			return;
		}

		if (!frm.subtitle.value){
			alert('subtitle을 등록해주세요');
			frm.subtitle.focus();
			return;
		}

		frm.submit();
	}

	function jsManagePlayImage(){
		var playManageDir = window.open('http://<%=CHKIIF(application("Svr_Info")="Dev","test","")%>upload.10x10.co.kr/linkweb/play/playManageDir.asp?folder=tepisode&idx=<%=idx%>','playManageDir','width=1000,height=600,scrollbars=yes,resizable=yes');
		playManageDir.focus();
	}
//-->
</script>

<form name="inputfrm" method="post" action="photopickProc.asp">
<input type="hidden" name="idx" value="<%= idx %>"/>
<input type="hidden" name="photopickimg" value="<%=PPimg%>"/>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>T-episode등록/수정</b></font><br/><br/>
	</td>
</tr>
<% If idx <> "0" Then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">idx</td>
	<td bgcolor="#FFFFFF">
		<b><%=idx%></b>
	</td>
</tr>
<% End If %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">제목</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="viewtitle" value="<%= ReplaceBracket(viewtitle) %>" size="50"/>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">서브카피</td>
	<td bgcolor="#FFFFFF">
		<textarea name="subtitle" rows="8" cols="50"><%= ReplaceBracket(subtitle) %></textarea>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">사용</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" <%= chkiif(isusing = "Y", "checked", "")%> >Y
		<input type="radio" name="isusing" value="N" <%= chkiif(isusing = "N", "checked", "")%> >N
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnorgmg" value="이미지등록" onClick="jsSetImg('<%=PPimg%>','photopickimg','photopickimgdiv')" class="button"/> **이미지 세로길이 566px로 맞춰주세요
		<div id="photopickimgdiv" style="padding: 5 5 5 5">
			<%If PPimg <> "" THEN %>
				<img src="<%=PPimg%>" border="0" height=100 onclick="jsImgView('<%=PPimg%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('photopickimg','photopickimgdiv');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				이미지 주소 : <%=PPimg%>
			<%END IF%>
		</div>
	</td>
</tr>
<% If idx <> "0" Then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">Tag</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnviewImg" value="태그 관리" onClick="jsTagview('<%=idx%>', '')" class="button"/><br/><br/>
		※태그관리는 팝업으로 관리 합니다 개별 등록 해주세요.※
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">
		모바일수작업영역
		<% If idx <> "" Then %>
			<br /><br /><br /><br /><br /><input type="button" value="이미지관리" class="button" onClick="jsManagePlayImage('<%=idx%>');">
		<% End If %>		
	</td>
	<td bgcolor="#FFFFFF">
		<textarea name="style_html_m" style="width:100%; height:240px;"><%= ReplaceBracket(style_html_m) %></textarea>
	</td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" 저장 " class="button" onclick="subcheck();"/> &nbsp;&nbsp;
		<input type="button" value=" 취소 " class="button" onclick="location.href='/admin/sitemaster/play/tepisode/?menupos=<%=menupos%>';"/>
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
