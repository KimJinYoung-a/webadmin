<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 스타일픽 관리
' Hieditor : 2011.04.06 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylelifeCls.asp"-->

<%
dim menupos ,oTheme ,catename
dim idx,title,subcopy,state,banner_img,title_img,startdate,enddate,isusing,regdate,comment, sortno
dim lastadminid,cd1,opendate,closedate,partMDid,partWDid
	idx = request("idx")
	menupos = request("menupos")

'//이벤트정보
set oTheme = new ClsStyleLife
	oTheme.frectidx = idx
	
	if idx <> "" then
		oTheme.fnGetTheme_item()
		
		if oTheme.ftotalcount > 0 then			
			title = oTheme.foneitem.ftitle
			subcopy = oTheme.foneitem.fsubcopy
			state = oTheme.foneitem.fstate
			banner_img = oTheme.foneitem.fbanner_img
			title_img = oTheme.foneitem.ftitle_img
			startdate = left(oTheme.foneitem.fstartdate,10)
			enddate = left(oTheme.foneitem.fenddate,10)
			regdate = oTheme.foneitem.fregdate
			comment = oTheme.foneitem.fcomment
			lastadminid = oTheme.foneitem.flastadminid
			cd1 = oTheme.foneitem.fcd1
			opendate = oTheme.foneitem.fopendate
			closedate = oTheme.foneitem.fclosedate
			partMDid = oTheme.foneitem.fpartMDid
			partWDid = oTheme.foneitem.fpartWDid
			catename = oTheme.foneitem.fcatename
			sortno = oTheme.foneitem.fsortno
		end if	
	end if
set oTheme = nothing
	
if isusing = "" then isusing = "Y"
%>

<script language="javascript">

	//-- jsPopCal : 달력 팝업 --//
	function jsPopCal(sName){
		var winCal;

		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	//이미지 확대화면 새창으로 보여주기
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
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
		winImg = window.open('pop_theme_uploadimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	//저장
	function jsEvtSubmit(){

		if(!frm.cd1.value){
			alert("스타일을 선택해주세요");
			frm.cd1.focus();
			return;
		}

		if(!frm.title.value){
			alert("제목을 입력해주세요");
			frm.title.focus();
			return;
		}
		
		if(!frm.state.value){
			alert("상태를 선택해주세요");
			frm.state.focus();
			return;
		}
	
		if(!frm.startdate.value){
			alert("시작일을 입력해주세요");
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
	
	function TextCD1(g)
	{
		if(g == "0P0")
		{
			document.getElementById("txtcd1").innerHTML = "스타일픽<br>센터";
		}
		else
		{
			document.getElementById("txtcd1").innerHTML = "타이틀";
		}
	}
</script>

<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" action="/admin/stylepick/stylelife_theme_process.asp" method="post">
<input type="hidden" name="mode" value="eventedit">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="banner_img" value="<%=banner_img%>">
<input type="hidden" name="title_img" value="<%=title_img%>">
<input type="hidden" name="opendate" value="<%=opendate%>">
<input type="hidden" name="closedate" value="<%=closedate%>">
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">기획전번호</td>
	<td bgcolor="#FFFFFF"><%= idx %><input type="hidden" name="idx" value="<%=idx%>"></td>
</tr>	
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">스타일</td>
	<td bgcolor="#FFFFFF">
		<select name="cd1" onChange="TextCD1(this.value);">
			<option value="">-스타일-</option>
			<option value="010" <%=CHKIIF(cd1="010","selected","")%>>클래식</option>
			<option value="020" <%=CHKIIF(cd1="020","selected","")%>>큐트</option>
			<option value="040" <%=CHKIIF(cd1="040","selected","")%>>모던</option>
			<option value="050" <%=CHKIIF(cd1="050","selected","")%>>네추럴</option>
			<option value="060" <%=CHKIIF(cd1="060","selected","")%>>오리엔탈</option>
			<option value="070" <%=CHKIIF(cd1="070","selected","")%>>팝</option>
			<option value="080" <%=CHKIIF(cd1="080","selected","")%>>로맨틱</option>
			<option value="090" <%=CHKIIF(cd1="090","selected","")%>>빈티지</option>
			<option value="0P0" <%=CHKIIF(cd1="0P0","selected","")%>>스타일픽</option>
		</select>
	</td>
</tr>
	
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">제목</td>
	<td bgcolor="#FFFFFF"><input type="text" size=64 maxlength=64 name="title" value="<%=title%>"></td>
</tr>
	
<!--
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">서브카피</td>
	<td bgcolor="#FFFFFF"><input type="text" size=64 maxlength=64 name="subcopy" value="<%=subcopy%>"></td>
</tr>
//-->
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">상태</td>
	<td bgcolor="#FFFFFF">
		<% Draweventstate2 "state" , state ,"" %>
		&nbsp;&nbsp;&nbsp;※ 실제 오픈되는 조건은 상태가 <font color="red"><b>오픈</b></font>, 기간이 <font color="red"><b>시작일 <= 현재일</b></font> 로써 두가지 모두 성립 된 것만 보이게 됩니다.
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">기간</td>
	<td bgcolor="#FFFFFF">
   		<%IF state = "9" THEN%>
   			시작일 : <%=startdate%><input type="hidden" name="startdate" size=10 maxlength=10 value="<%=startdate%>">
   		<%ELSE%>
   			시작일 : <input type="text" name="startdate" size=10 maxlength=10 value="<%=startdate%>" onClick="jsPopCal('startdate');"  style="cursor:hand;">
   		<%END IF%>
   		<%
		if opendate <> "1900-01-01" and opendate <> "" then response.write " 오픈처리일 : " & opendate
		if closedate <> "1900-01-01" and closedate <> "" then response.write " 종료처리일 : " & closedate
		%>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">담당MD</td>
	<td bgcolor="#FFFFFF">
		<% sbGetpartid "partmdid",partmdid,"","11,21" %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">담당WD</td>
	<td bgcolor="#FFFFFF">
		<% sbGetpartid "partwdid",partwdid,"","12" %>
	</td>
</tr>
<% If idx <> "" Then %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">순서</td>
	<td bgcolor="#FFFFFF"><input type="text" size=7 maxlength=5 name="sortno" value="<%=sortno%>"></td>
</tr>
<% End If %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">작업전달사항</td>
	<td bgcolor="#FFFFFF">
		<textarea rows=10 cols=100 name="comment"><%=comment%></textarea>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">기본배너이미지</td>
	<td bgcolor="#FFFFFF">		
		<input type="button" name="btnBanImg" value="이미지등록" onClick="jsSetImg('<%=banner_img%>','banner_img','banner_imgdiv')" class="button">
		<div id="banner_imgdiv" style="padding: 5 5 5 5">
			<%IF banner_img <> "" THEN %>			
				<img src="<%=banner_img%>" border="0" width=100 height=100 onclick="jsImgView('<%=banner_img%>');" alt="누르시면 확대 됩니다">
				<a href="javascript:jsDelImg('banner_img','banner_imgdiv');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center"><span id="txtcd1">타이틀</span>이미지</td>
	<td bgcolor="#FFFFFF">		
		<input type="button" name="btnTitImg" value="이미지등록" onClick="jsSetImg('<%=title_img%>','title_img','title_imgdiv')" class="button">
		<div id="title_imgdiv" style="padding: 5 5 5 5">
			<%IF title_img <> "" THEN %>			
				<img src="<%=title_img%>" border="0" width=100 height=100 onclick="jsImgView('<%=title_img%>');" alt="누르시면 확대 됩니다">
				<a href="javascript:jsDelImg('title_img','title_imgdiv');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="2" align="center"><input type="button" onclick="jsEvtSubmit();" class="button" value="저장"></td>
</tr>	
</form>
</table>

<script>
<%
If cd1 = "0P0" Then
Response.Write "TextCD1('0P0');"
End If
%>
</script>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
