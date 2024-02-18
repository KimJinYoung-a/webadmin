<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  play
' History : 2013.09.03 이종화 생성
'			2014.10.21 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/play/playCls.asp" -->
<%
Dim idx , listimg , state , reservationdate , viewtitle , viewtext , playcate, viewno , textimg , worktext
Dim viewimg1, viewimg2 ,viewimg3 ,viewimg4 ,viewimg5, partMDid,partWDid, style_html_m
Dim itemcnt1,itemcnt2,itemcnt3,itemcnt4,itemcnt5
	idx = request("idx")

playcate = 2 'Style+

dim oPlay
set oPlay = new CPlayContents
	 oPlay.FRectIdx = idx
	
	if idx <> "" Then
		oPlay.GetOneRowStyleContent()

		if oPlay.FResultCount > 0 then
			style_html_m			= oPlay.FOneItem.fstyle_html_m
			listimg					= oPlay.FOneItem.Flistimg
			viewimg1				= oPlay.FOneItem.Fviewimg1
			viewimg2				= oPlay.FOneItem.Fviewimg2
			viewimg3				= oPlay.FOneItem.Fviewimg3
			viewimg4				= oPlay.FOneItem.Fviewimg4
			viewimg5				= oPlay.FOneItem.Fviewimg5
			textimg					= oPlay.FOneItem.Ftextimg
			viewtitle				= oPlay.FOneItem.Fviewtitle
			reservationdate			= oPlay.FOneItem.Freservationdate
			state					= oPlay.FOneItem.Fstate
			viewno					= oPlay.FOneItem.Fviewno
			worktext				= oPlay.FOneItem.Fworktext
			partMDid				= oPlay.FOneItem.FpartMDid
			partWDid				= oPlay.FOneItem.FpartWDid
			itemcnt1				= oPlay.FOneItem.Fitemcnt1
			itemcnt2				= oPlay.FOneItem.Fitemcnt2
			itemcnt3				= oPlay.FOneItem.Fitemcnt3
			itemcnt4				= oPlay.FOneItem.Fitemcnt4
			itemcnt5				= oPlay.FOneItem.Fitemcnt5
		end if	
	end if
set oPlay = Nothing
%>

<script type="text/javascript">

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
	poptag = window.open('/admin/sitemaster/play/lib/pop_tagReg.asp?idx='+idx+'&playcate='+<%=playcate%>,'poptag','width=1024,height=768,scrollbars=yes,resizable=yes');
	poptag.focus();
}

function jsSetItem(idx , styleno){
	var popitem;
	popitem = window.open('/admin/sitemaster/play/lib/pop_itemReg.asp?idx='+idx+'&number='+styleno,'popitem','width=500,height=400,scrollbars=yes,resizable=yes');
	popitem.focus();
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

	if(!frm.partmdid.value){
		alert("담당 MD를 선택하세요.");
		frm.partmdid.focus();
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
    var playManageDir = window.open('http://<%=CHKIIF(application("Svr_Info")="Dev","test","")%>upload.10x10.co.kr/linkweb/play/playManageDir.asp?folder=style&idx=<%=idx%>','playManageDir','width=1000,height=600,scrollbars=yes,resizable=yes');
    playManageDir.focus();
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="/admin/sitemaster/play/styleplus/styleplusProc.asp">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="styleviewimg1" value="<%=viewimg1%>">
<input type="hidden" name="styleviewimg2" value="<%=viewimg2%>">
<input type="hidden" name="styleviewimg3" value="<%=viewimg3%>">
<input type="hidden" name="styleviewimg4" value="<%=viewimg4%>">
<input type="hidden" name="styleviewimg5" value="<%=viewimg5%>">
<input type="hidden" name="stylelistimg" value="<%=listimg%>">
<input type="hidden" name="styletitleimg" value="<%=textimg%>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>Play&gt;&gt;Style+ 등록/수정</b></font><br/><br/>
		※ 상세이미지 저장은 최대 5장까지 가능하며 이미지 저장후 꼭!!!! 아이템을 등록 해주셔야 합니다.※<br/>
		※ 아이템 등록 버튼은 이미지 저장후 생성됩니다 ※
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
		<input type="text" class="text" name="viewtitle" value="<%=viewtitle%>" size="100"/>
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
		<% sbGetpartid "partmdid",partmdid,"","23" %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">담당WD</td>
	<td bgcolor="#FFFFFF">
		<% sbGetpartid "partwdid",partwdid,"","12" %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">작업 전달 사항</td>
	<td bgcolor="#FFFFFF">
		<textarea name="worktext" rows="8" cols="50"><%=worktext%></textarea>
	</td>
</tr>
<% If idx <> "0" Then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">리스트이미지(썸네일)</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnlistImg" value="이미지등록" onClick="jsSetImg('<%=listimg%>','stylelistimg','stylelistimgdiv')" class="button"/>
		<div id="stylelistimgdiv" style="padding: 5 5 5 5">
			<%IF listimg <> "" THEN %>			
				<img src="<%=listimg%>" border="0" height=100 onclick="jsImgView('<%=listimg%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('stylelistimg','stylelistimgdiv');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				이미지 주소 : <%=listimg%>
			<%END IF%>
		</div>
		(이미지 Size는 240 x 자유 입니다.)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">타이틀이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnorgmg" value="이미지등록" onClick="jsSetImg('<%=textimg%>','styletitleimg','styletitleimgdiv')" class="button"/>
		<div id="styletitleimgdiv" style="padding: 5 5 5 5">
			<%If textimg <> "" THEN %>			
				<img src="<%=textimg%>" border="0" height=100 onclick="jsImgView('<%=textimg%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('styletitleimg','styletitleimgdiv');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				이미지 주소 : <%=textimg%>
			<%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">상세이미지 1</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnviewImg1" value="이미지등록" onClick="jsSetImg('<%=viewimg1%>','styleviewimg1','styleviewimgdiv1')" class="button"/>
		<% IF viewimg1 <> "" THEN %>
		<div style="position:absolute;width:100%;margin-left:-100px;">
			<div style="position:relative;float:right;">※이미지 저장후 아이템을 꼭 등록 해주세요※<marquee style="font:11 돋움체;" behavior="alternate" width="36" scrollamount="50" scrolldelay="200"> --&gt;</marquee><input type="button" value="아이템등록[<%=itemcnt1%>]" onClick="jsSetItem('<%=idx%>','1')" class="button"/></div>
		</div>
		<% End If %>
		<div id="styleviewimgdiv1" style="padding: 5 5 5 5">
			<%IF viewimg1 <> "" THEN %>			
				<img src="<%=viewimg1%>" border="0" height=100 onclick="jsImgView('<%=viewimg1%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('styleviewimg1','styleviewimgdiv1');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				이미지 주소 : <%=viewimg1%>
			<%END IF%>
		</div>
		(이미지 Size는 1140 x 560 입니다.)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">상세이미지 2</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnviewImg2" value="이미지등록" onClick="jsSetImg('<%=viewimg2%>','styleviewimg2','styleviewimgdiv2')" class="button"/>
		<% IF viewimg2 <> "" THEN %>
		<div style="position:absolute;width:100%;margin-left:-100px;">
			<div style="position:relative;float:right;">※ 이미지 저장후 아이템을 꼭 등록 해주세요 ※<marquee style="font:11 돋움체;" behavior="alternate" width="36" scrollamount="50" scrolldelay="200"> --&gt;</marquee><input type="button" value="아이템등록[<%=itemcnt2%>]" onClick="jsSetItem('<%=idx%>','2')" class="button"/></div>
		</div>
		<% End If %>
		<div id="styleviewimgdiv2" style="padding: 5 5 5 5">
			<%IF viewimg2 <> "" THEN %>			
				<img src="<%=viewimg2%>" border="0" height=100 onclick="jsImgView('<%=viewimg2%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('styleviewimg2','styleviewimgdiv2');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				이미지 주소 : <%=viewimg2%>
			<%END IF%>
		</div>
		(이미지 Size는 1140 x 560 입니다.)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">상세이미지 3</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnviewImg3" value="이미지등록" onClick="jsSetImg('<%=viewimg3%>','styleviewimg3','styleviewimgdiv3')" class="button"/>
		<% IF viewimg3 <> "" THEN %>
		<div style="position:absolute;width:100%;margin-left:-100px;">
			<div style="position:relative;float:right;">※ 이미지 저장후 아이템을 꼭 등록 해주세요 ※<marquee style="font:11 돋움체;" behavior="alternate" width="36" scrollamount="50" scrolldelay="200"> --&gt;</marquee><input type="button" value="아이템등록[<%=itemcnt3%>]" onClick="jsSetItem('<%=idx%>','3')" class="button"/></div>
		</div>
		<% End If %>
		<div id="styleviewimgdiv3" style="padding: 5 5 5 5">
			<%IF viewimg3 <> "" THEN %>			
				<img src="<%=viewimg3%>" border="0" height=100 onclick="jsImgView('<%=viewimg3%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('styleviewimg3','styleviewimgdiv3');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				이미지 주소 : <%=viewimg3%>
			<%END IF%>
		</div>
		(이미지 Size는 1140 x 560 입니다.)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">상세이미지 4</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnviewImg4" value="이미지등록" onClick="jsSetImg('<%=viewimg4%>','styleviewimg4','styleviewimgdiv4')" class="button"/>
		<% IF viewimg4 <> "" THEN %>
		<div style="position:absolute;width:100%;margin-left:-100px;">
			<div style="position:relative;float:right;">※ 이미지 저장후 아이템을 꼭 등록 해주세요 ※<marquee style="font:11 돋움체;" behavior="alternate" width="36" scrollamount="50" scrolldelay="200"> --&gt;</marquee><input type="button" value="아이템등록[<%=itemcnt4%>]" onClick="jsSetItem('<%=idx%>','4')" class="button"/></div>
		</div>
		<% End If %>
		<div id="styleviewimgdiv4" style="padding: 5 5 5 5">
			<%IF viewimg4 <> "" THEN %>			
				<img src="<%=viewimg4%>" border="0" height=100 onclick="jsImgView('<%=viewimg4%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('styleviewimg4','styleviewimgdiv4');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				이미지 주소 : <%=viewimg4%>
			<%END IF%>
		</div>
		(이미지 Size는 1140 x 560 입니다.)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">상세이미지 5</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnviewImg5" value="이미지등록" onClick="jsSetImg('<%=viewimg5%>','styleviewimg5','styleviewimgdiv5')" class="button"/>
		<% IF viewimg5 <> "" THEN %>
		<div style="position:absolute;width:100%;margin-left:-100px;">
			<div style="position:relative;float:right;">※ 이미지 저장후 아이템을 꼭 등록 해주세요 ※<marquee style="font:11 돋움체;" behavior="alternate" width="36" scrollamount="50" scrolldelay="200"> --&gt;</marquee><input type="button" value="아이템등록[<%=itemcnt5%>]" onClick="jsSetItem('<%=idx%>','5')" class="button"/></div>
		</div>
		<% End If %>
		<div id="styleviewimgdiv5" style="padding: 5 5 5 5">
			<%IF viewimg5 <> "" THEN %>			
				<img src="<%=viewimg5%>" border="0" height=100 onclick="jsImgView('<%=viewimg5%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('styleviewimg5','styleviewimgdiv5');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				이미지 주소 : <%=viewimg5%>
			<%END IF%>
		</div>
		(이미지 Size는 1140 x 560 입니다.)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">상세페이지 태그</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnviewImg" value="태그 관리" onClick="jsTagview('<%=idx%>')" class="button"/><br/><br/>
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
		<textarea name="style_html_m" style="width:100%; height:240px;"><%=style_html_m%></textarea>
	</td>
</tr>
<% End If %>
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
