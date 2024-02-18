<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/appmanage/hitchhiker.asp" -->
<%
Dim mode, idx, hitch
mode = request("mode")
idx = request("idx")

Dim Vdevice, Vstartdate, Venddate, Vcontents, Visusing
Set hitch = new Hitchhiker
If mode = "U" Then
	hitch.Midx = idx
	hitch.FNoticeView

	Vdevice = hitch.FNdevice
	Vstartdate = hitch.FNstartdate
	Venddate = hitch.FNenddate
	Vcontents = hitch.FNcontents
	Visusing = hitch.FNisusing
ElseIf mode = "II" OR mode = "UU" Then
	Dim assigndevice, startdate, enddate, contents, isusing
	assigndevice= request("assigndevice")
	startdate	= request("startdate")
	enddate 	= request("enddate")
	contents	= request("contents")
	isusing		= request("isusing")

	If mode = "II" Then
		Dim sSQL
		sSQL = ""
		sSQL = sSQL & "INSERT INTO db_contents.dbo.tbl_hhiker_notice " & VBCRLF
		sSQL = sSQL & "(assigndevice, startdate, enddate, contents, regdate, isusing) VALUES " & VBCRLF
		sSQL = sSQL & "('"&assigndevice&"', '"&startdate&"', '"&enddate&"', '"&db2html(contents)&"', getdate(), '"&isusing&"') " & VBCRLF
	ElseIf mode = "UU" Then
		sSQL = ""
		sSQL = sSQL & "UPDATE db_contents.dbo.tbl_hhiker_notice SET " & VBCRLF
		sSQL = sSQL & "assigndevice = '"&assigndevice&"', startdate = '"&startdate&"', enddate = '"&enddate&"', contents = '"&db2html(contents)&"', isusing = '"&isusing&"' " & VBCRLF
		sSQL = sSQL & "WHERE idx = '"&idx&"' " & VBCRLF
	End If
	dbget.execute(sSQL)
	Response.Write "<script language='javascript'>" & vbCrLf &_
			"alert('저장되었습니다.');"& vbCrLf &_
			"opener.location.reload();" & vbCrLf &_
			"window.close();"& vbCrLf &_
			"</script>"
	response.End
End If
Set hitch = nothing
%>
<script language="javascript">
function NoticeWrite(){
	var frm = document.frmcontents;
	if(frm.startdate.value==""){
		alert("시작일을 입력하세요");
		frm.startdate.focus();
		return;
	}
	if(frm.enddate.value==""){
		alert("종료일을 입력하세요");
		frm.enddate.focus();
		return;
	}
	if(frm.contents.value==""){
		alert("내용을 입력하세요");
		frm.contents.focus();
		return;
	}
	var chk = 0;
	for(var j=0; j < frm.isusing.length; j++) {
		if(frm.isusing[j].checked) chk++;
	}
	if (chk < 1){
		alert("사용유무에 체크하세요");
		return false;
	}
	if(confirm("저장하시겠습니까?")){
		if(document.getElementById("mode").value == "I"){
			document.getElementById("mode").value = "II";
		}else if(document.getElementById("mode").value == "U"){
			document.getElementById("mode").value = "UU";
		}
		frm.submit();	
	}
}
</script>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmcontents" method="post" action="pop_hitch_NoticeWrite.asp">
<input type="hidden" id="mode" name="mode" value="<%=mode%>">
<input type="hidden" name="idx" value="<%=idx%>">
<tr bgcolor="#FFFFFF">
	<td>Device</td>
	<td>
		<select name="assigndevice" class="select">
			<option value="0" <%=chkiif(Vdevice="0","selected","")%> >전체
			<option value="1" <%=chkiif(Vdevice="1","selected","")%> >IOS
			<option value="2" <%=chkiif(Vdevice="2","selected","")%> >Android
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>시작일</td>
	<td>
		<input type="text" name="startdate" size="10" maxlength=10 readonly value="<%=Vstartdate%>"> 00:00:00
		<a href="javascript:calendarOpen(frmcontents.startdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>종료일</td>
	<td>
		<input type="text" name="enddate" size="10" maxlength=10 readonly value="<%=Venddate%>"> 00:00:00
		<a href="javascript:calendarOpen(frmcontents.enddate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>내용</td>
	<td>
		<textarea class="textarea" name="contents" cols="66" rows="10"><%=Vcontents%></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>사용여부</td>
	<td>
		<input type="radio" name="isusing" value="Y" <%=chkiif(Visusing="Y","checked","")%> >Y
		<input type="radio" name="isusing" value="N" <%=chkiif(Visusing="N","checked","")%> >N
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center" colspan=2>
    	<input type="button" value="저장" onClick="NoticeWrite();" class="button">
    </td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
