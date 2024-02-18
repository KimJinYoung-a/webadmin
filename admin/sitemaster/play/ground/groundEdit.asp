<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/play/playCls.asp" -->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
	Dim idx , listimg , state , reservationdate , viewtitle , viewtext , playcate , mainimg
	Dim viewno , worktext , partMKid,partWDid , i
	Dim oPlay , oground
	idx = request("idx")
    playcate = 1 'ground

	'//db 1row
	set oPlay = new CPlayContents
		 oPlay.FRectIdx = idx

		if idx <> "" Then
			oPlay.GetRowGroundMain()

			if oPlay.FResultCount > 0 then
				listimg					= oPlay.FOneItem.Flistimg
				viewtitle				= oPlay.FOneItem.Fviewtitle
				mainimg					= oPlay.FOneItem.Fmainimg
				reservationdate			= oPlay.FOneItem.Freservationdate
				state					= oPlay.FOneItem.Fstate
				viewno					= oPlay.FOneItem.Fviewno
				worktext				= oPlay.FOneItem.Fworktext
				partMKid				= oPlay.FOneItem.FpartMKid
				partWDid				= oPlay.FOneItem.FpartWDid
			end if
		end If
		
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

	function AddNewContents(idx,gidx){
		var popwin = window.open('/admin/sitemaster/play/ground/groundweekEdit.asp?idx=' + idx+'&gidx='+gidx,'cateHotPosCodeEdit','width=800,height=500,scrollbars=yes,resizable=yes');
		popwin.focus();
	}

	function jsSetImg(sImg, sName, sSpan){
		document.domain ="10x10.co.kr";

		var winImg;
		winImg = window.open('/admin/sitemaster/play/lib/pop_theme_uploadimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	function jsTagview(gidx , idx){	
		var poptag;
		poptag = window.open('/admin/sitemaster/play/lib/pop_tagReg.asp?idx='+gidx+'&subidx='+idx+'&playcate='+<%=playcate%>,'poptag','width=500,height=400,scrollbars=yes,resizable=yes');
		poptag.focus();
	}

	function subcheck(){
		var frm=document.inputfrm;

		if (!frm.viewno.value){
			alert('No.를 등록해주세요');
			frm.viewno.focus();
			return;
		}

		if (!frm.viewtitle.value){
			alert('상세제목을 등록해주세요');
			frm.viewtitle.focus();
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

		frm.submit();
	}

	function workerlist() //담당자
	{
		var openWorker = null;
		var worker = inputfrm.selMId.value;
		openWorker = window.open('/admin/sitemaster/play/lib/PopWorkerList.asp?worker='+worker+'&team=22','openWorker','width=570,height=570,scrollbars=yes');
		openWorker.focus();
	}

	function jsSetItem(idx){//상품등록 확인
		var popitem;
		popitem = window.open('pop_itemReg.asp?idx='+idx,'popitem','width=500,height=400,scrollbars=yes,resizable=yes');
		popitem.focus();
	}
//-->
</script>
<script type="text/javascript">
<!--
	function copy_url(url) {
		var IE=(document.all)?true:false;
		if (IE) {
			if(confirm("이 글의 URL 주소를 클립보드에 복사하시겠습니까?"))
				window.clipboardData.setData("Text", url);
		} else {
			temp = prompt("이 글의 트랙백 주소입니다. Ctrl+C를 눌러 클립보드로 복사하세요", url);
		}
	}
//-->
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="groundProc.asp">
<input type="hidden" name="idx" value="<%= idx %>"/>
<input type="hidden" name="groundtitleimg" value="<%=listimg%>"/>
<input type="hidden" name="playmainimg" value="<%=mainimg%>"/>
<input type="hidden" name="position" value="main"/>
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>Play&gt;&gt;ground 메인 등록/수정</b></font><br/><br/>
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
		<input type="text" class="text" name="viewno" value="<%=viewno%>" size="10"/>※ 숫자만 적어주세요 ※
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">상세제목</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="viewtitle" value="<%=viewtitle%>" size="50"/>
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
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">상태</td>
	<td bgcolor="#FFFFFF">
		<% Draweventstate2 "state" , state ,"" %> ※ 오픈을 해서 저장하여도 시작일 =< 오늘 이어야만 노출이 됩니다.
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">담당자</td>
	<td bgcolor="#FFFFFF">
		<% sbGetwork "selMId",partMKid,"" %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">담당WD</td>
	<td bgcolor="#FFFFFF">
		<% sbGetpartid "partwdid",partwdid,"","12" %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">타이틀이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnorgmg" value="이미지등록" onClick="jsSetImg('<%=listimg%>','groundtitleimg','groundtitleimgdiv')" class="button"/>
		<div id="groundtitleimgdiv" style="padding: 5 5 5 5">
			<%If listimg <> "" THEN %>
				<img src="<%=listimg%>" border="0" height=100 onclick="jsImgView('<%=listimg%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('groundtitleimg','groundtitleimgdiv');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				이미지 주소 : <%=listimg%>
			<%END IF%>
		</div>
	</td>
</tr>
<!-- <tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">play메인용썸네일</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnorgmg" value="이미지등록" onClick="jsSetImg('<%=mainimg%>','playmainimg','playmainimgdiv')" class="button"/>
		<div id="playmainimgdiv" style="padding: 5 5 5 5">
			<%If mainimg <> "" THEN %>
				<img src="<%=mainimg%>" border="0" height=100 onclick="jsImgView('<%=mainimg%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('playmainimg','playmainimgdiv');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				이미지 주소 : <%=mainimg%>
			<%END IF%>
		</div>
	</td>
</tr> -->
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">작업 전달 사항</td>
	<td bgcolor="#FFFFFF">
		<textarea name="worktext" rows="8" cols="50"><%=worktext%></textarea>
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

<% If idx > "0" Then %>
<%
	set oground = new CPlayContents
		oground.FPageSize = 50
		oground.FCurrPage = 1
		oground.FRPlaycate = playcate
		oground.FRectIdx = idx
		oground.fnGetGroundSubList()

	
%>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 10px 0;">
<tr>
	<td align="left">
		<font color="red"> ※ 리스트 노출 : 상태가 오픈인 것과 시작일 =< 오늘 인것만 노출이 됩니다. 순서는 No. 번호(높은순서) 순서로 노출됩니다.<br/>※ 하단 리스트중 상태 부분을 누르시면 리뷰 페이지가 열립니다.</font>
	</td>
	<td align="right">
		<input type="button" class="button" value="신규등록" onclick="AddNewContents('0','<%=idx%>');">
	</td>
</tr>
</table>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="5%">회차</td>
	<td width="10%">상태</td>
	<td width="10%">타이틀이미지</td>
	<td width="10%">제목</td>
	<td width="5%">시작일</td>
	<td width="10%">태그</td>
	<td width="5%">담당자</td>
	<td width="5%">기획WD</td>
	<td width="7%">관리</td>
</tr>
<% if oground.FresultCount > 0 then %>
<% for i=0 to oground.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td align="center"><%= oground.FItemList(i).Fviewno %></td>
	<td align="center"><%= geteventstate(oground.FItemList(i).Fstate) %> (<%=oground.FItemList(i).Fstate %>)<br/><br/>
	<a href="http://www.10x10.co.kr/play/playGround_review.asp?gidx=<%=idx%>&gcidx=<%=oground.FItemList(i).Fidxsub%>" target="_blank">PC미리보기</a>&nbsp;&nbsp;<input type="button" onclick="copy_url('http://www.10x10.co.kr/play/playGround.asp?gidx=<%=idx%>&gcidx=<%=oground.FItemList(i).Fidxsub%>')" value="PC-URL복사"/><br/>
	<a href="http://m.10x10.co.kr/play/playGround_review.asp?idx=<%=oground.FItemList(i).Fmo_idx%>&contentsidx=<%=oground.FItemList(i).Fidxsub%>" target="_blank">M 미리보기</a>&nbsp;&nbsp;<input type="button" onclick="copy_url('http://m.10x10.co.kr/play/playGround.asp?idx=<%=oground.FItemList(i).Fmo_idx%>&contentsidx=<%=oground.FItemList(i).Fidxsub%>')" value="M-URL복사"/></td>
	<td align="center"><img src="<%= oground.FItemList(i).Fviewthumbimg1 %>" width="80"/>&nbsp;<img src="<%= oground.FItemList(i).Fviewthumbimg2 %>" width="80"/></td>
	<td align="center"><%= oground.FItemList(i).Fviewtitle %></td>
	<td align="center"><%= left(oground.FItemList(i).Freservationdate,10) %></td>
	<td align="center"><a href="#" onclick="jsTagview('<%=idx%>','<%= oground.FItemList(i).Fidxsub %>');" style="cursor:pointer;"><%=chkiif(oground.FItemList(i).Ftagcnt>0,"등록","미등록")%>(<%=oground.FItemList(i).Ftagcnt%>) </a></td>
	<td align="center"><%= oground.FItemList(i).FpartMKname %></td>
	<td align="center"><%= oground.FItemList(i).FpartWDname %></td>
	<td align="center">
		<input type="button" class="button" value="수정" onclick="AddNewContents('<%= oground.FItemList(i).Fidxsub %>','<%=idx%>');"/>
		<input type="button" value="아이템등록[<%=oground.FItemList(i).Fitemcnt%>]" onClick="jsSetItem('<%= oground.FItemList(i).Fidxsub %>')" class="button"/>
	</td>
</tr>
<% Next %>
<% end if %>
</table>
<%
	set oground = nothing
%>
<% End If %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
