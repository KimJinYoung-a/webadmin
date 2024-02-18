<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/appmanage/hitchhiker.asp" -->
<%
Dim idx, mode, vol, rev, hitch
Dim detailseq, dgunm
idx = request("idx")
mode = request("mode")

Dim bannerImg, clickURL, isusing, usetype, startdate, enddate
If mode = "U" Then
Set hitch = new Hitchhiker
	hitch.Midx = idx
	hitch.HitchBannereView

	bannerImg = hitch.SBbannerImg
	clickURL = hitch.SBclickURL
	usetype = hitch.SBusetype
	isusing = hitch.SBisusing
	startdate = hitch.SBstartdate
	enddate = hitch.SBenddate
Set hitch = nothing
End If
%>
<script language="javascript">
<!--
	document.domain ="10x10.co.kr";
	function jsUpload(){
		if(document.regfrm.clickURL.value == ""){
			alert("배너이미지 클릭시 이동할 URL을 입력하세요");
			document.regfrm.clickURL.focus();
			return false;
		}
		var chk = 0;
		for(var j=0; j<document.regfrm.isusing.length; j++) {
			if(document.regfrm.isusing[j].checked) chk++;
		}
		if (chk < 1){
			alert("사용유무에 체크하세요");
			return false;
		}
	}
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 파일 업로드 처리</div>
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="regfrm" method="post" action="<%=staticImgUrl%>/linkweb/appmanage/banner_upload.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="mode" value="<%=mode%>">
<tr bgcolor="#FFFFFF">
    <td width="100" align="center">Idx</td>
    <td><%=Idx%></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" align="center">파일</td>
	<td bgcolor="#FFFFFF">
		<% If bannerImg <> "" Then %>
			<img src="<%=bannerImg%>" width="100" height="100">
		<% End If %>
		<input type="file" name="Files" size= "35">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" align="center">타입</td>
    <td>
    	<select name="usetype" class="select">
    		<option value="bnImage" <%=chkiif(usetype="bnImage","selected","")%>>배너이미지
    	</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" align="center">시작일</td>
    <td>
		<input type="text" name="startdate" size="10" maxlength=10 readonly value="<%=startdate%>"> 00:00:00
		<a href="javascript:calendarOpen(regfrm.startdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" align="center">종료일</td>
    <td>
		<input type="text" name="enddate" size="10" maxlength=10 readonly value="<%=enddate%>"> 00:00:00
		<a href="javascript:calendarOpen(regfrm.enddate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
    </td>
</tr>
<tr bgcolor="#FFFFFF" id="clickURL">
    <td width="100" align="center">clickURL</td>
    <td><input type="text" name="clickURL" size="60" value="<%=clickURL%>"></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" align="center">사용유무</td>
    <td>
    	<input type="radio" name="isusing" value="Y" <%=chkiif(isusing="Y","checked","")%>>Y
    	<input type="radio" name="isusing" value="N" <%=chkiif(isusing="N","checked","")%>>N
    </td>
</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right">
		<input type="image" src="/images/icon_confirm.gif">
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->