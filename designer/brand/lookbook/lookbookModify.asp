<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.19 김진영 생성
'			2013.08.29 한용민 수정
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/lookbookCls.asp"-->
<%
Dim mode, olookbook
Dim idx, makerid, title, state, mainimg, isusing, sortNo, regdate, lastupdate, regadminid,lastadminid
dim comment
	mode	= requestCheckVar(request("mode"),30)
	idx		= requestCheckVar(Request("idx"),10)
	makerid	= requestCheckVar(request("makerid"),50)
	menupos	= requestCheckVar(request("menupos"),10)
	
If idx = "" Then
	mode = "I"
Else
	mode = "U"
End If

makerid = session("ssBctID")

SET olookbook = new clookbook
	olookbook.FrectIdx = idx
	olookbook.frectmakerid = makerid
	
	if idx <> "" then
		olookbook.sblookbookmodify
	end if
	
	if olookbook.ftotalcount > 0 then
		idx = olookbook.FOneItem.Fidx
		makerid = olookbook.FOneItem.Fmakerid
		title = olookbook.FOneItem.Ftitle
		state = olookbook.FOneItem.Fstate
		mainimg = olookbook.FOneItem.Fmainimg
		isusing = olookbook.FOneItem.Fisusing
		sortNo = olookbook.FOneItem.FsortNo
		regdate = olookbook.FOneItem.Fregdate
		lastupdate = olookbook.FOneItem.Flastupdate
		regadminid = olookbook.FOneItem.Fregadminid
		lastadminid = olookbook.FOneItem.Flastadminid
		comment = olookbook.FOneItem.fcomment
	end if
%>

<script language="javascript">

function form_check(mode){
	var frm = document.frm;
	
	if(frm.makerid.value==''){
		alert('브랜드를 선택하세요.');
		frm.makerid.focus();
		return;
	}
	
	if(frm.title.value==''){
		alert('제목을 입력하세요.');
		frm.title.focus();
		return;
	}
	
	if(frm.isusing.value==''){
		alert('사용여부를 선택하세요.');
		frm.isusing.focus();
		return;
	}
	
	if(frm.mainimg.value==""){
		alert('메인 이미지를 등록하세요');
		frm.mainimg.focus();
		return;
	}

	var state = '<%= state %>';
	var message;
	if (state=='7'){
		message = '오픈상태에서 수정을 하실경우, 상태가 등록중 상태가 되며,\n텐바이텐에 승인요청을 하셔야 합니다.\n\n저장하시겠습니까?';
	}else{
		message = '저장하시겠습니까?';
	}

	if(confirm(message)){
		frm.mode.value=mode;
		frm.submit();
	}
}

function jsSetImg(sFolder, sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/designer/brand/lookbook/pop_lookbook_uploadimg.asp?sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}

//상태변경
function chstate(state){
	if(confirm("상태를 변경 하시겠습니까?")){
		frmchstate.mode.value='chstate';
		frmchstate.state.value=state;
		frmchstate.submit();
	}
}

</script>

<!-- #include virtual="/designer/brand/inc_streetHead.asp"-->

<img src="/images/icon_arrow_link.gif"> <b><b>LOOKBOOK 등록</b></b>

<table border="0" cellpadding="0" cellspacing="0" class="a" width="100%">
<form name="frmchstate" method="post" action="/designer/brand/lookbook/lookbook_process.asp" style="margin:0px;">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="state">
	<input type="hidden" name="idx" value="<%=idx%>">
	<input type="hidden" name="mode">
</form>
<form name="frm" method="post" action="/designer/brand/lookbook/lookbook_process.asp" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="mainimg" value="<%=mainimg%>">
<input type="hidden" name="statcd" value="">
<tr>
	<td style="padding-bottom:10">
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" width="100%">
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>" width=200>번호</td>
			<td bgcolor="#FFFFFF">
				<%=idx%>
				<input type="hidden" name="idx" value="<%=idx%>">
			</td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">브랜드</td>
			<td bgcolor="#FFFFFF">
				<%= makerid %>
				<input type="hidden" name="makerid" value="<%= makerid %>">	
			</td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">제목</td>
			<td bgcolor="#FFFFFF"><input type="text" size="70" maxlength=50 name="title" value="<%= title %>"></td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">상태</td>
			<td bgcolor="#FFFFFF">
				<% if mode="U" then %>
					<%' drawlookbookstats "state" , state , " onchange='gosubmit("""");'" %>
					<%= lookbookstatsname(state) %>
					<input type="hidden" name="state" value="<%=state%>">
				<% else %>
					등록중
				<% end if %>
			</td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">사용</td>
			<td bgcolor="#FFFFFF" >
				<% drawSelectBoxUsingYN "isusing", isusing %>
			</td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">썸네일이미지<br>(필수입력 : 259 x 360)</td>
			<td bgcolor="#FFFFFF">
				<input type="button" name="btnBan" value="이미지등록" onClick="jsSetImg('lookbook','<%= mainimg %>','mainimg','spanban')" class="button">
	   			<div id="spanban" style="padding: 5 5 5 5">
	   				<% IF mainimg <> "" THEN %>
	   					<img src="<%=mainimg%>" border="0" width="259" height="360">
	   					<a href="javascript:jsDelImg('mainimg','spanban');"><img src="/images/icon_delete2.gif" border="0"></a>
	   				<%END IF%>
	   			</div>
			</td>
		</tr>
		
		<% If mode = "U" Then %>
			<tr>
				<td align="center"  bgcolor="<%= adminColor("tabletop") %>">
					작업코맨트
					<Br>(반려사유나 본사코맨트)					
				</td>
				<td bgcolor="#FFFFFF" >
					<%= nl2br(comment) %>
					<input type="hidden" name="comment" value="<%=comment%>">
				</td>
			</tr>
			<tr>
				<td align="center"  bgcolor="<%= adminColor("tabletop") %>">LOOKBOOK이미지</td>
				<td bgcolor="#FFFFFF">
					<iframe id="iframG" frameborder="0" width="100%" src="/designer/brand/lookbook/iframe_lookbook_detail.asp?idx=<%=idx%>" height=300></iframe>
				</td>
			</tr>
		<% else %>
			<tr>
				<td align="center"  bgcolor="<%= adminColor("tabletop") %>">LOOKBOOK이미지</td>
				<td bgcolor="#FFFFFF">
					신규등록 완료후 LOOKBOOK이미지를 입력 하실수 있습니다.
				</td>
			</tr>			
		<% End If %>
		
		<tr align="center">
			<td bgcolor="#FFFFFF" colspan=2>
				<% If mode = "U" Then %>
					<% If state = "1" or state = "2" or state = "7" Then %>
						<input type="button" value="수정" class="button" onclick="form_check('U');">
					<% end if %>
				<% elseif mode = "I" Then %>
					<input type="button" value="신규등록" class="button" onclick="form_check('I');">
				<% End If %>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<% If mode = "U" Then %>
					<%
					'/반려(수정요청)일경우
					If state = "1" Then
					%>
						<input type="button" value="승인요청" class="button" onclick="chstate('3');">
					<% end if %>
					<%
					'/등록중일경우
					If state = "2" Then
					%>
						<input type="button" value="승인요청" class="button" onclick="chstate('3');">
					<% end if %>
				<% End If %>
			</td>
		</tr>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->