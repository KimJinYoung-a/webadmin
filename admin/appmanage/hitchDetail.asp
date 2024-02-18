<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/appmanage/hitchhiker.asp" -->
<%
Dim midx, vol, rev
midx = request("midx")
vol = request("vol")
rev = request("rev")

Dim page, hitch, i
If page = "" Then page = 1
Set hitch = new Hitchhiker
	hitch.FPageSize = 20
	hitch.Midx = midx
	hitch.FCurrPage = page
	hitch.HitchDetailList
%>
<script language="javascript">
function DetailUP(){
	var winImg;
	winImg = window.open('/admin/appmanage/pop_hitch_detail.asp?mode=I&midx=<%=midx%>&vol=<%=vol%>&rev=<%=rev%>','popImg','width=650,height=400, status=yes');
	winImg.focus();
}
function DetailUpdate(dseq, dgunm, ddevice){
	var winImg2;
	winImg2 = window.open('/admin/appmanage/pop_hitch_detail.asp?mode=U&detailSeq='+dseq+'&dgunm='+dgunm+'&midx=<%=midx%>&vol=<%=vol%>&rev=<%=rev%>&device='+ddevice+'','popImg2','width=650,height=400, status=yes');
	winImg2.focus();
}
</script>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="frm" method="get">
<tr>
	<td align="left">
		<input type= "button" value="상세업로드" class="button" onclick="javascript:DetailUP();">
	</td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30"><td><img src="/images/icon_arrow_link.gif"></td><td style="padding-top:3">&nbsp;<b>히치하이커 Detail</b></td></tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>seq</td>
	<td>구분</td>
	<td>orgfileName</td>
	<td>contURL</td>
	<td>음악제목</td>
	<td>음악가</td>
	<td>배너클릭시 링크</td>
	<td>사용유무</td>
	<td>정렬번호</td>
</tr>
<%
	For i = 0 to hitch.FResultCount -1
%>
<tr height="25" bgcolor="FFFFFF" onclick="javascript:DetailUpdate('<%=hitch.FhitchList(i).FctSeq%>','<%=hitch.FhitchList(i).Fctgbnname%>', '<%=hitch.FhitchList(i).Fdevice%>');" style="cursor:pointer;">
	<td align="center"><%=hitch.FhitchList(i).FctSeq%></td>
	<td align="center"><%=hitch.FhitchList(i).Fctgbnname%></td>
	<td align="center"><%=hitch.FhitchList(i).ForgfileName%></td>
	<td align="left">
	<% If hitch.FhitchList(i).Fctgbnname = "bgImage" Then %>
		<img src="<%=hitch.FhitchList(i).FcontURL%>" width="60" height="80">
	<% ElseIf hitch.FhitchList(i).Fctgbnname = "bgSound" Then %>
	    <%= hitch.FhitchList(i).FcontURL %> <font color='red'>(IOS)</font><br>
	    <%= replace(replace(hitch.FhitchList(i).FcontURL,"http://","rtsp://"),"/playlist.m3u8","") %> <font color='red'>(AND)</font><br>
	<%
'			response.write hitch.FhitchList(i).FcontURL
'			If trim(hitch.FhitchList(i).Fdevice) = "" or trim(hitch.FhitchList(i).Fdevice) = "ios" Then
'				response.write " <strong><font color='red'>(IOS)</font></strong>"
'			ElseIf trim(hitch.FhitchList(i).Fdevice) = "android" Then
'				response.write " <strong><font color='blue'>(ANDROID)</font></strong>"
'			End If
	   End If
	%>
	</td>
	<td align="center"><%=hitch.FhitchList(i).FmusicTitle%></td>
	<td align="center"><%=hitch.FhitchList(i).Fmusician%></td>
	<td align="center"><%=hitch.FhitchList(i).FlinkURL%></td>
	<td align="center"><%=hitch.FhitchList(i).Fisusing%></td>
	<td align="center"><%=hitch.FhitchList(i).ForderNo%></td>
</tr>
<%
	Next
%>
</form>
</table>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type= "button" value="이전리스트로" class="button" onclick="javascript:location.href='hitchList.asp';">
	</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->