<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  기프트
' History : 2014.03.19 한용민 생성
' History : 2014.10.31 유태욱 mtitle 추가
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/giftday_cls.asp"-->

<%
Dim mode, cgiftday, menupos
Dim masteridx, title, mtitle, startdate, enddate, listtopimg_w, listtopimg_m, regtopimg_W, regtopimg_m, mainimg_W, regdate, isusing
	mode = requestCheckVar(request("mode"),32)
	masteridx = requestCheckVar(getNumeric(request("masteridx")),10)
	menupos = requestCheckVar(getNumeric(request("menupos")),10)

SET cgiftday = new Cgiftday_list
	cgiftday.Frectmasteridx = masteridx

	if masteridx <> "" then
		cgiftday.getgiftday_master_one
	end if
	
	if cgiftday.ftotalcount > 0 then
		masteridx = cgiftday.FOneItem.Fmasteridx
		title = ReplaceBracket(cgiftday.FOneItem.Ftitle)
		mtitle = ReplaceBracket(cgiftday.FOneItem.Fmtitle)
		startdate = cgiftday.FOneItem.Fstartdate
		enddate = cgiftday.FOneItem.Fenddate
		listtopimg_w = cgiftday.FOneItem.flisttopimg_w
		listtopimg_m = cgiftday.FOneItem.flisttopimg_m
		regtopimg_w = cgiftday.FOneItem.fregtopimg_w
		regtopimg_m = cgiftday.FOneItem.fregtopimg_m
		mainimg_W = cgiftday.FOneItem.fmainimg_W
		regdate = cgiftday.FOneItem.Fregdate
		isusing = cgiftday.FOneItem.Fisusing
	end if
	
if isusing="" then isusing="Y"
%>

<script language="javascript">

function form_check(){
	var frm = document.frm;
	
	if(frm.title.value==''){
		alert('제목을 입력하세요.');
		frm.title.focus();
		return;
	}
	if(frm.mtitle.value==''){
		alert('모바일 제목을 입력하세요.');
		frm.mtitle.focus();
		return;
	}
	if(frm.startdate.value==''){
		alert('시작일을 입력하세요.');
		frm.startdate.focus();
		return;
	}
	if(frm.enddate.value==''){
		alert('종료일을 입력하세요.');
		frm.enddate.focus();
		return;
	}
	if(frm.isusing.value==''){
		alert('사용여부를 선택하세요.');
		frm.isusing.focus();
		return;
	}
	
	if(confirm('저장하시겠습니까?')){
		frm.mode.value='dayedit';
		frm.submit();
	}
}

function jsSetImg(sFolder, sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/admin/sitemaster/gift/day/giftday_edit_uploadimg.asp?sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
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

<table border="0" cellpadding="0" cellspacing="0" class="a" width="100%">
<form name="frm" method="post" action="/admin/sitemaster/gift/day/giftday_process.asp" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="listtopimg_w" value="<%=listtopimg_w%>">
<input type="hidden" name="listtopimg_m" value="<%=listtopimg_m%>">
<input type="hidden" name="regtopimg_w" value="<%=regtopimg_w%>">
<input type="hidden" name="regtopimg_m" value="<%=regtopimg_m%>">
<input type="hidden" name="mainimg_W" value="<%=mainimg_W%>">
<tr>
	<td style="padding-bottom:10">
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" width="100%">
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>" width=200>번호</td>
			<td bgcolor="#FFFFFF">
				<%=masteridx%>
				<input type="hidden" name="masteridx" value="<%=masteridx%>">
			</td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">제목</td>
			<td bgcolor="#FFFFFF"><input type="text" size="50" maxlength=50 name="title" value="<%= title %>"></td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">모바일 제목</td>
			<td bgcolor="#FFFFFF"><input type="text" size="50" maxlength=50 name="mtitle" value="<%= mtitle %>"></td>
		</tr>	
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">기간</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="startdate" size=10 value="<%= startdate %>" class="text">			
				<a href="javascript:calendarOpen3(frm.startdate,'시작일',frm.startdate.value)">
				<img src="/images/calicon.gif" width="21" border="0" align="middle"></a> -
				<input type="text" name="enddate" size=10  value="<%= left(enddate,10) %>" class="text">
				<a href="javascript:calendarOpen3(frm.enddate,'마지막일',frm.enddate.value)">
				<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>	
			</td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">WWW<br>리스트탑이미지</td>
			<td bgcolor="#FFFFFF">
				<input type="button" name="btnBan" value="이미지등록" onClick="jsSetImg('listtopimg_w','<%= listtopimg_w %>','listtopimg_w','sblisttopimg_w')" class="button">
	   			<div id="sblisttopimg_w" style="padding: 5 5 5 5">
	   				<% IF listtopimg_w <> "" THEN %>
	   					<img src="<%=listtopimg_w%>" border="0" width="259" height="360">
	   					<a href="javascript:jsDelImg('listtopimg_w','sblisttopimg_w');"><img src="/images/icon_delete2.gif" border="0"></a>
	   				<%END IF%>
	   			</div>
			</td>
		</tr>
		<!--<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">WWW<br>메인배너이미지</td>
			<td bgcolor="#FFFFFF">
				<input type="button" name="btnBan" value="이미지등록" onClick="jsSetImg('mainimg_W','<%= mainimg_W %>','mainimg_W','sbmainimg_W')" class="button">
	   			<div id="sbmainimg_W" style="padding: 5 5 5 5">
	   				<% IF mainimg_W <> "" THEN %>
	   					<img src="<%=mainimg_W%>" border="0" width="259" height="360">
	   					<a href="javascript:jsDelImg('mainimg_W','sbmainimg_W');"><img src="/images/icon_delete2.gif" border="0"></a>
	   				<%END IF%>
	   			</div>
			</td>
		</tr>-->		
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">M<br>리스트탑이미지</td>
			<td bgcolor="#FFFFFF">
				<input type="button" name="btnBan" value="이미지등록" onClick="jsSetImg('listtopimg_m','<%= listtopimg_m %>','listtopimg_m','sblisttopimg_m')" class="button">
	   			<div id="sblisttopimg_m" style="padding: 5 5 5 5">
	   				<% IF listtopimg_m <> "" THEN %>
	   					<img src="<%=listtopimg_m%>" border="0" width="259" height="360">
	   					<a href="javascript:jsDelImg('listtopimg_m','sblisttopimg_m');"><img src="/images/icon_delete2.gif" border="0"></a>
	   				<%END IF%>
	   			</div>
			</td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">WWW<Br>등록탑이미지</td>
			<td bgcolor="#FFFFFF">
				<input type="button" name="btnBan" value="이미지등록" onClick="jsSetImg('regtopimg_w','<%= regtopimg_w %>','regtopimg_w','sbregtopimg_w')" class="button">
	   			<div id="sbregtopimg_w" style="padding: 5 5 5 5">
	   				<% IF regtopimg_w <> "" THEN %>
	   					<img src="<%=regtopimg_w%>" border="0" width="259" height="360">
	   					<a href="javascript:jsDelImg('regtopimg_w','sbregtopimg_w');"><img src="/images/icon_delete2.gif" border="0"></a>
	   				<%END IF%>
	   			</div>
			</td>
		</tr>			
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">M<Br>등록탑이미지</td>
			<td bgcolor="#FFFFFF">
				<input type="button" name="btnBan" value="이미지등록" onClick="jsSetImg('regtopimg_m','<%= regtopimg_m %>','regtopimg_m','sbregtopimg_m')" class="button">
	   			<div id="sbregtopimg_m" style="padding: 5 5 5 5">
	   				<% IF regtopimg_m <> "" THEN %>
	   					<img src="<%=regtopimg_m%>" border="0" width="259" height="360">
	   					<a href="javascript:jsDelImg('regtopimg_m','sbregtopimg_m');"><img src="/images/icon_delete2.gif" border="0"></a>
	   				<%END IF%>
	   			</div>
			</td>
		</tr>		
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">사용</td>
			<td bgcolor="#FFFFFF" >
				<% drawSelectBoxUsingYN "isusing", isusing %>
			</td>
		</tr>
		<tr align="center">
			<td bgcolor="#FFFFFF" colspan=2>
				<input type="button" value="저장" class="button" onclick="form_check();">
			</td>
		</tr>
	</td>
</tr>
</form>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->