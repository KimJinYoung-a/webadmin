<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/play/playCls.asp" -->
<%
	Dim idx , listimg , state , reservationdate , viewtitle , viewtext , playcate , videourl , partWDid, videourlM
	Dim viewno ,  worktext
	Dim oPlay
	idx = request("idx")
    playcate = 6 '비디오클립
	'//db 1row
	set oPlay = new CPlayContents
		 oPlay.FRectIdx = idx

		if idx <> "" Then
			oPlay.GetOneRowVideoClipContent()

			if oPlay.FResultCount > 0 then
				listimg = oPlay.FOneItem.Flistimg
				viewtitle = oPlay.FOneItem.Fviewtitle
				viewtext = oPlay.FOneItem.Fviewtext
				viewno = oPlay.FOneItem.Fviewno
				worktext = oPlay.FOneItem.Fworktext
				reservationdate = oPlay.FOneItem.Freservationdate
				state = oPlay.FOneItem.Fstate
				videourl  = oPlay.FOneItem.Fvideourl
				partWDid	= oPlay.FOneItem.FpartWDid
				videourlM = oPlay.FOneItem.FvideourlM
			end if
		end if
	set oPlay = Nothing
%>

<script type="text/javascript">
<!--
//-- jsPopCal : 달력 팝업 --//
	function jsPopCal(sName){
		var winCal;

		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

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

	function jsTagview(idx){
		var poptag;
		poptag = window.open('/admin/sitemaster/play/lib/pop_tagReg.asp?idx='+idx+'&playcate='+<%=playcate%>,'poptag','width=500,height=400,scrollbars=yes,resizable=yes');
		poptag.focus();
	}


	function subcheck(){
		var frm=document.inputfrm;

		if (!frm.viewno.value){
			alert('No.을 등록해주세요');
			frm.viewno.focus();
			return;
		}

		if (!frm.viewtitle.value){
			alert('상세제목을 등록해주세요');
			frm.viewtitle.focus();
			return;
		}

		if (!frm.viewtext.value){
			alert('상세내용을 등록해주세요');
			frm.viewtext.focus();
			return;
		}

		if (!frm.worktext.value){
			alert('작업내용을 등록해주세요');
			frm.worktext.focus();
			return;
		}

		if (!frm.reservationdate.value){
			alert('오픈예정일을 등록해주세요');
			frm.reservationdate.focus();
			return;
		}

		if(!frm.state.value){
			alert("상태를 선택해주세요");
			frm.state.focus();
			return;
		}

		if(!frm.partwdid.value){
			alert("담당 WD를 선택하세요.");
			frm.partwdid.focus();
			return;
		}

		frm.submit();
	}

	function jsManagePlayImage(){
		var playManageDir = window.open('http://<%=CHKIIF(application("Svr_Info")="Dev","test","")%>upload.10x10.co.kr/linkweb/play/playManageDir.asp?folder=video&idx=<%=idx%>','playManageDir','width=1000,height=600,scrollbars=yes,resizable=yes');
		playManageDir.focus();
	}
//-->
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="videoclipProc.asp">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="pdlistimg" value="<%=listimg%>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>Play&gt;&gt;비디오클립 등록/수정</b></font>
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
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">No.</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="viewno" value="<%=viewno%>" size="10"/>※ 숫자가 클수록 우선 표시 됩니다. ※
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">제목</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="viewtitle" value="<%=viewtitle%>" size="50"/>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">상태</td>
	<td bgcolor="#FFFFFF">
		<% Draweventstate2 "state" , state ,"" %> ※ 오픈을 해서 저장하여도 시작일 =< 오늘 이어야만 노출이 됩니다.
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">담당WD</td>
	<td bgcolor="#FFFFFF">
		<% sbGetpartid "partwdid",partwdid,"","12" %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">시작일</td>
	<td bgcolor="#FFFFFF">
   		<%IF state = "9" THEN%>
   			<%=reservationdate%><input type="hidden" name="reservationdate" size=20 maxlength=10 value="<%=reservationdate%>"/>
   		<%ELSE%>
   			<input type="text" name="reservationdate" size=20 maxlength=10 value="<%=reservationdate%>" onClick="jsPopCal('reservationdate');"  style="cursor:pointer;"/>
   		<%END IF%>
		예) (<%=Left(Now(),10)%>)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">썸네일이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnlistImg" value="이미지등록" onClick="jsSetImg('<%=listimg%>','pdlistimg','listimgdiv')" class="button"/>
		<div id="listimgdiv" style="padding: 5 5 5 5">
			<%IF listimg <> "" THEN %>
				<img src="<%=listimg%>" border="0" height=100 onclick="jsImgView('<%=listimg%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('pdlistimg','listimgdiv');"><img src="/images/icon_delete2.gif" border="0"/></a>
			<%END IF%>
		</div>
	</td>
</tr>

<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">상세내용</td>
	<td bgcolor="#FFFFFF">
		<textarea name="viewtext" rows="8" cols="50"><%=viewtext%></textarea>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">작업 전달 사항</td>
	<td bgcolor="#FFFFFF">
		<textarea name="worktext" rows="8" cols="50"><%=worktext%></textarea>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">영상링크</td>
	<td bgcolor="#FFFFFF">
		<textarea name="videourl" rows="5" cols="50"><%=videourl%></textarea><br/><br/>
		※ &lt;iframe width="960" height="540" src="주소" frameborder="0" allowfullscreen &gt;  &lt;/iframe&gt;<br/>※ width , height 확인요망 <br/>위와 같은 형식으로 입력 해주세요
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">
		영상링크(모바일용)
		<% If idx <> "" Then %>
			<br /><br /><br /><br /><br /><input type="button" value="이미지관리" class="button" onClick="jsManagePlayImage('<%=idx%>');">
		<% End If %>		
	</td>
	<td bgcolor="#FFFFFF">
		<textarea name="videourlM" rows="5" cols="50"><%=videourlM%></textarea><br/><br/>
		※ &lt;iframe src="주소" frameborder="0" allowfullscreen &gt;  &lt;/iframe&gt;<br/>※ width , height 값은 입력하지 않습니다. <br/>위와 같은 형식으로 입력 해주세요
	</td>
</tr>

<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">상세페이지 태그</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnviewImg" value="태그 관리" onClick="jsTagview('<%=idx%>')" class="button"/><br/><br/>
		※태그관리는 팝업으로 관리 합니다 개별 등록 해주세요.※
	</td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" 저장 " class="button" onclick="subcheck();"/> &nbsp;&nbsp;
		<input type="button" value=" 취소 " class="button" onclick="history.back();"/>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
