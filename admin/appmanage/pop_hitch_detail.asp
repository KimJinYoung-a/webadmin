<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/appmanage/hitchhiker.asp" -->
<%
Dim midx, mode, vol, rev, hitch, device
Dim detailseq, dgunm
midx = request("midx")
vol = request("vol")
rev = request("rev")
mode = request("mode")
detailseq = request("detailseq")
dgunm = request("dgunm")
device = request("device")
Dim ctgbnname, linkURL, isusing, orderNo, musicTitle, musician
If mode = "U" Then
Set hitch = new Hitchhiker
	hitch.Midx = midx
	hitch.Ctseq = detailseq
	hitch.Ctgbnname = dgunm
	hitch.HitchDetailView

	ctgbnname = hitch.Dctgbnname
	musicTitle = hitch.DmusicTitle
	musician = hitch.Dmusician
	linkURL = hitch.DlinkURL
	isusing = hitch.Disusing
	orderNo = hitch.DorderNo
Set hitch = nothing
End If
%>
<script language="javascript">
<!--
	//document.domain ="10x10.co.kr";
	function jsUpload(){
//		if(!document.regfrm.detailFile.value){
//			alert("찾아보기 버튼을 눌러 업로드할 이미지를 선택해 주세요.");
//			return false;
//		}
		if(document.getElementById("lkurl").style.display == "block" && document.regfrm.linkUrl.value == ""){
			alert("배너이미지 클릭시 이동할 URL을 입력하세요");
			document.regfrm.linkUrl.focus();
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
		if(document.regfrm.orderNo.value == ""){
			alert("정렬번호를 입력하세요");
			document.regfrm.orderNo.focus();
			return false;
		}
	}
	function chgGBN(g){
		if(g=="bgSound"){
			document.getElementById("lkurl").style.display = "none";
			document.getElementById("muTitle").style.display = "block";
			document.getElementById("muCian").style.display = "block";
			document.getElementById("device").style.display = "block";
			document.regfrm.linkUrl.value = "";
		}else{
			document.getElementById("lkurl").style.display = "block";
			document.getElementById("muTitle").style.display = "none";
			document.getElementById("muCian").style.display = "none";
			document.getElementById("device").style.display = "none";
		}
	}
//-->
</script>
<!-- staticImgUrl ==> staticUploadUrl -->
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 파일 업로드 처리</div>
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="regfrm" method="post" action="<%=staticUploadUrl%>/linkweb/appmanage/package_upload2.asp?vol=<%=vol%>&rev=<%=rev%>&detailseq=<%=detailseq%>" enctype="multipart/form-data" onSubmit="return jsUpload();">
<input type="hidden" name="midx" value="<%=midx%>">
<input type="hidden" name="mode" value="<%=mode%>">
<tr bgcolor="#FFFFFF">
    <td width="100" align="center">mIdx</td>
    <td><%=mIdx%></td>
</tr>

<tr bgcolor="#FFFFFF" <%=chkiif(mode="U","style=display:none","")%> >
	<td width="100" align="center">파일</td>
	<td bgcolor="#FFFFFF"><input type="file" name="detailFile" size= "35"></td>
</tr>
<tr bgcolor="#FFFFFF" <%=chkiif(mode="U","style=display:none","")%>>
    <td width="100" align="center">구분</td>
    <td>
    	<select name="ctgbnname" class="select" onchange="javascript:chgGBN(this.value);">
    		<option value="bgImage" <%=chkiif(ctgbnname="bgImage","selected","")%>>배너이미지
   			<option value="bgSound" <%=chkiif(ctgbnname="bgSound","selected","")%>>배경음악
    	</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF" id="lkurl" style="display:block;">
    <td width="100" align="center">linkURL</td>
    <td><input type="text" name="linkUrl" size="60" value="<%=linkURL%>"></td>
</tr>

<tr bgcolor="#FFFFFF" id="device" style="display:none;">
    <td width="100" align="center">운영체제</td>
    <td>
        <INPUT TYPE="hidden" name="device" value=""> <!-- ALL -->
        ALL
        <!--
    	<select name="device" class="select">
    		<option value="ios" <% If device = "" OR device = "IOS" Then response.write "selected" End If %> >IOS
   			<option value="android" <% If device = "android" Then response.write "selected" End If %> >ANDROID
    	</select>
    	-->
    </td>
</tr>
<tr bgcolor="#FFFFFF" id="muTitle" style="display:none;">
    <td width="100" align="center">음악제목</td>
    <td><input type="text" name="musicTitle" size="60" value="<%=musicTitle%>"></td>
</tr>
<tr bgcolor="#FFFFFF" id="muCian" style="display:none;">
    <td width="100" align="center">음악가</td>
    <td><input type="text" name="musician" size="40" value="<%=musician%>"></td>
</tr>
<script language="javascript">
<% If ctgbnname = "bgSound" Then %>
document.getElementById("lkurl").style.display = "none";
document.getElementById("muTitle").style.display = "block";
document.getElementById("muCian").style.display = "block";
document.getElementById("device").style.display = "block";
<% End If %>
</script>
<tr bgcolor="#FFFFFF">
    <td width="100" align="center">사용유무</td>
    <td>
    	<input type="radio" name="isusing" value="Y" <%=chkiif(isusing="Y","checked","")%>>Y
    	<input type="radio" name="isusing" value="N" <%=chkiif(isusing="N","checked","")%>>N
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" align="center">정렬번호</td>
    <td><input type="text" name="orderNo" size="3" maxlength="3" value="<%=orderNo%>"></td>
</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right">
		<input type="image" src="/images/icon_confirm.gif">
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->