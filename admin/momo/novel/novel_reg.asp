<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 한줄소설
' Hieditor : 2009.11.17 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim ocontents,i , wordimage , winner
dim novelid,startdate,enddate,regdate,prolog,title,genre,isusing
	novelid = requestcheckvar(request("novelid"),8)

'//상세
set ocontents = new cnovel_list
	ocontents.frectnovelid = novelid
	
	'//수정일경우에만 쿼리
	if novelid <> "" then
	ocontents.fnovel_oneitem()
	end if
		
	if ocontents.ftotalcount > 0 then
		novelid = ocontents.FOneItem.fnovelid
		startdate = ocontents.FOneItem.fstartdate
		enddate = ocontents.FOneItem.fenddate
		regdate = ocontents.FOneItem.fregdate
		prolog = ocontents.FOneItem.fprolog
		title = ocontents.FOneItem.ftitle
		genre = ocontents.FOneItem.fgenre
		isusing = ocontents.FOneItem.fisusing
		wordimage = ocontents.FOneItem.fwordimage
		winner = ocontents.FOneItem.fwinner
	end if
%>

<script language="javascript">

	document.domain = "10x10.co.kr";
	
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
	
	function jsImgDel(divnm,iptNm,vPath){
	
		window.open('','imgdel','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
		document.imginputfrm.divName.value=divnm;
		document.imginputfrm.inputname.value=iptNm;
		document.imginputfrm.ImagePath.value = vPath;
		document.imginputfrm.maxFileSize.value = Fsize;
		document.imginputfrm.maxFileWidth.value = Fwidth;
		document.imginputfrm.makeThumbYn.value = thumb;
		document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
		document.imginputfrm.target='imgdel';
		document.imginputfrm.action='PopImgInput.asp';
		document.imginputfrm.submit();
	}

	//저장
	function reg(){
		if (frm.title.value==''){
		alert('주제를 입력해주세요');
		frm.title.focus();
		return;
		}
		if (frm.genre.value==''){
		alert('장르를 입력해주세요');
		frm.genre.focus();
		return;
		}
		if (frm.startdate.value==''){
		alert('시작일을 입력해주세요');
		frm.startdate.focus();
		return;
		}		
		if (frm.enddate.value==''){
		alert('종료일을 입력해주세요');
		frm.enddate.focus();
		return;
		}	
		if (frm.prolog.value==''){
		alert('프롤로그를 입력해주세요');
		frm.prolog.focus();
		return;
		}						
		if (frm.isusing.value==''){
		alert('사용여부를 선택해주세요');
		return;
		}
		
		frm.action='/admin/momo/novel/novel_process.asp';
		frm.mode.value='edit';
		frm.submit();
	}
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post">
<input type="hidden" name="mode" >
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>한줄소설ID</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= novelid %><input type="hidden" name="novelid" value="<%= novelid %>">		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>초판발행일</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= regdate %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>주제</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="title" value="<%= title %>" size=64 maxlength=35>		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>장르</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="genre" value="<%= genre %>" size=20 maxlength=10>		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><b>기간</b><br></td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="startdate" size=10 value="<%= startdate %>">			
		<a href="javascript:calendarOpen3(frm.startdate,'시작일',frm.startdate.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a> -
		<input type="text" name="enddate" size=10  value="<%= left(enddate,10) %>">
		<a href="javascript:calendarOpen3(frm.enddate,'마지막일',frm.enddate.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>	
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>프롤로그</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="prolog" style="width:450px; height:100px;"><%=prolog%></textarea>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>소설시작이미지</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('wordimgdiv','wordimg','word','2000','800','true');"/>		
		<input type="hidden" name="wordimg" value="<%= wordimage %>">
		<div align="right" id="wordimgdiv"><% IF wordimage<>"" THEN %><img src="<%=webImgUrl%>/momo/novel/word/<%= wordimage %>" width=50 height=50 style="cursor:pointer" ><% End IF %></div>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>당첨자ID</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" name="winner" value="<%=winner%>"> ※노출이 필요한 경우만 입력하세요	
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>사용여부</td>
	<td bgcolor="#FFFFFF" align="left">
		<select name="isusing" value="<%=isusing%>">
			<option value="" <% if isusing = "" then response.write " selected" %>>사용여부</option>
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>			
	</td>
</tr>
<tr align="center" bgcolor="FFFFFF">
	<td colspan=2><input type="button" onclick="reg();" value="저장" class="button"></td>
</tr>
</form>
</table>
<%
	set ocontents = nothing
%>
<form name="imginputfrm" method="post" action="">
	<input type="hidden" name="divName" value="">
	<input type="hidden" name="orgImgName" value="">
	<input type="hidden" name="inputname" value="">
	<input type="hidden" name="ImagePath" value="">
	<input type="hidden" name="maxFileSize" value="">
	<input type="hidden" name="maxFileWidth" value="">
	<input type="hidden" name="makeThumbYn" value="">
</form>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->