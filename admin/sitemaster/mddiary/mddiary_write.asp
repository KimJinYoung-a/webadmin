<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/mddiaryCls.asp"-->

<%
	Dim mddiary, vMgzID, vMenuImg, vMenuImgOn, vMainImg, vOpenDate, vRegDate, vUseYN, vUseMap
	vMgzID 	= NullFillWith(Request("mgzId"),"")


	If vMgzID <> "" Then	
		Set mddiary = new Clsmddiary
		mddiary.FMgzID = vMgzID
		mddiary.FmddiaryCont
		
		vMenuImg = mddiary.FMenuImg
		vMenuImgOn = mddiary.FMenuImg_On
		vMainImg = mddiary.FMainImg
		vOpenDate = mddiary.FOpenDate
		vRegDate = mddiary.FRegdate
		vUseYN = mddiary.FUseYN
		vUseMap = db2html(mddiary.FUseMap)
		set mddiary = nothing
	Else
		vUseYN = "Y"
		vOpenDate = FormatDateTime(now(),2)
	End If
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
<!--
	function jsUpload(){
		//if(!document.frmImg.sfImg.value){
		//	alert("찾아보기 버튼을 눌러 업로드할 이미지를 선택해 주세요.");			
		//	return false;
		//}
	}
	
//-->
</script>

<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/sitemaster/uploadMDDiary.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="mgzId" value="<%=vMgzID%>">
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">ID</td>
	<td><%=vMgzID%> <% If vMgzID <> "" Then Response.Write "&nbsp;(등록일:" & vRegDate & ")" End If %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">메뉴이미지 Off</td>
	<td>
		<input type="file" name="menuimg">
		<% If vMgzID <> "" Then %><br><img src="<%=vMenuImg%>"><% End If %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">메뉴이미지 On</td>
	<td>
		<input type="file" name="menuimg_on">
		<% If vMgzID <> "" Then %><br><img src="<%=vMenuImgOn%>"><% End If %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">본문이미지</td>
	<td>
		<input type="file" name="mainimg">
		<% If vMgzID <> "" Then %><br><img src="<%=vMainImg%>" width="50" height="50"><% End If %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">오픈일</td>
	<td>
		<input id="iSD" name="iSD" value="<%=vOpenDate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "iSD", trigger    : "iSD_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">사용여부</td>
	<td>
		<input type="radio" name="useyn" value="Y" <% If vUseYN = "Y" Then %>checked<% End If %>>Y&nbsp;&nbsp;&nbsp;
		<input type="radio" name="useyn" value="N" <% If vUseYN = "N" Then %>checked<% End If %>>N
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">Map</td>
	<td>
		<textarea name="usemap" rows="11" cols="55"><%=vUseMap%></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="right">
		<input type="image" src="/images/icon_confirm.gif">
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->