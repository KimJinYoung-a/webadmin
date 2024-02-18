<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  핑거스 아카데미 PC메인 작가&강사 링크 입력,수정 팝업
' History : 2016-10-24 유태욱 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/academy/PcMainLectureLinkCls.asp" -->
<%
	Dim idx, oPcMainLectureLink, sqlStr
	Dim startdate, titletext, contentstext, lectureid, isusing

	idx = RequestCheckvar(request("idx"),10)
	set oPcMainLectureLink = new CPcMainLectureLinkContents
		 oPcMainLectureLink.FRectIdx = idx

		if idx <> "" Then
			oPcMainLectureLink.GetOneRowPcMainLectureLinkContent()
			if oPcMainLectureLink.FResultCount > 0 then
				titletext	= oPcMainLectureLink.FOneItem.Ftitletext
				contentstext	= oPcMainLectureLink.FOneItem.Fcontentstext
				lectureid	= oPcMainLectureLink.FOneItem.Flectureid
				startdate	= oPcMainLectureLink.FOneItem.Fstartdate
				isusing	= oPcMainLectureLink.FOneItem.Fisusing
			end if
		end if
	set oPcMainLectureLink = Nothing

	if isusing = "" then isusing = "Y"
%>
<script type="text/javascript">
	//''jsPopCal : 달력 팝업
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	//저장
	function subcheck(){
		var frm=document.inputfrm;
	    
		if (!frm.titletext.value){
			alert('제목을 등록해주세요');
			frm.titletext.focus();
			return;
		}

		if (!frm.titletext.value){
			alert('설명을 등록해주세요');
			frm.titletext.focus();
			return;
		}

		if (!frm.startdate.value){
			alert('시작일을 등록해주세요');
			frm.startdate.focus();
			return;
		}

//		if (!frm.viewtext1.value){
//			alert('상세내용을 등록해주세요');
//			frm.viewtext1.focus();
//			return;
//		}

		frm.submit();
	}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="PcMainLectureLinkProc.asp">
<input type="hidden" name="idx" value="<%= idx %>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>핑거스PC메인 작가/강사 링크 등록/수정</b></font>
	</td>
</tr>
<!---------------------------------------------------------------------------------------->
<% If idx <> "0" Then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">idx</td>
	<td bgcolor="#FFFFFF">
		<b><%=idx%></b>
	</td>
</tr>
<% End If %>
<!---------------------------------------------------------------------------------------->
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">작가/강사 ID</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="lectureid" value="<%=lectureid%>" size="50"/>
	</td>
</tr>

<!---------------------------------------------------------------------------------------->

<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">제목</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="titletext" value="<%=titletext%>" size="50"/>
	</td>
</tr>
<!---------------------------------------------------------------------------------------->
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">설  명</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="contentstext" value="<%=contentstext%>" size="50"/>
	</td>
</tr>
<!---------------------------------------------------------------------------------------->
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">시작일</td>
	<td bgcolor="#FFFFFF">
   		<input type="text" name="startdate" size=20 maxlength=10 value="<%= startdate %>" onClick="jsPopCal('startdate');"  style="cursor:pointer;"/>
		<font color="red">☜클릭후 달력에서 선택</font>
	</td>
</tr>
<!---------------------------------------------------------------------------------------->
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center"> 사용여부 </td>
	<td colspan="2">
		<input type="radio" name="isusing" value="Y" <%=chkIIF(isusing="Y","checked","")%>/>사용함 &nbsp;&nbsp;&nbsp; 
		<input type="radio" name="isusing" value="N" <%=chkIIF(isusing="N","checked","")%>/>사용안함
	</td>
</tr>
<!---------------------------------------------------------------------------------------->
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" 저장 " class="button" onclick="subcheck();"/> &nbsp;&nbsp;
		<input type="button" value=" 취소 " class="button" onclick="window.close();"/>
	</td>
</tr>
<!---------------------------------------------------------------------------------------->
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->