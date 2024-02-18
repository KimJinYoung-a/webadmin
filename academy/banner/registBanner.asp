<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/academy/lib/academy_function.asp" -->
<!-- #include virtual="/academy/lib/classes/banner_cls.asp" -->
<%
'#############################################
' PageName : /academy/banner/registBanner.asp	
' Description : 핑거스 배너 등록
' History : 2006.11.16 정윤정 생성
'#############################################
	Dim iBId : iBId = RequestCheckvar(Request("iBId"),10)
	Dim iMaxSize :	iMaxSize = 10
	Dim sType, sImg, sLinkUrl, sCommCd, sIsUsing, sRegdate, sAdminId
	Dim simgUrl, sWidth, sHeight,sWindow
	Dim cBCont
	
	simgUrl = imgFingers & "/contents/banner/"	
	sWindow = "_self"
	IF iBId <> "" THEN
		sType = "U"	
	ELSE
		sType = "I"
	END If	
	
	IF sType ="U" THEN
		set cBCont = new ClsBannerCont
		cBCont.FBannerId = iBId
		cBCont.sbGetBannerView
		sImg = cBCont.FImgUrl
		sLinkUrl = cBCont.FLink
		sCommCd = cBCont.FCommCd
		sIsUsing = cBCont.FisUsing
		sRegdate = cBCont.FRegdate
		sAdminId = cBCont.FAdminId
		sWidth = cBCont.FWidth
		sHeight = cBCont.FWidth
		sWindow = cBCont.FWindow
		set cBCont = nothing
	END IF		
%>
<script language="javascript" src="/academy/lib/js/common.js"></script>
<script language="javascript">
	function jsSubmit(frm, sType){
	var iMaxSize;
	iMaxSize = <%=iMaxSize%>;
	
		if(!frm.selLoc.value){
			alert("배너위치를 선택해 주세요");
			frm.selLoc.focus();
			return false;
		}
		
	  
		if(!frm.sImg.value && sType == "I" ){
			alert("배너이미지를 등록해 주세요");
			frm.sImg.focus();
			return false;
		}
		
		if(!fnChkFile(frm.sImg.value, iMaxSize)){
			return false;
		}
		
		if(!frm.sLUrl.value){
			alert("배너 링크주소를 등록해 주세요");
			frm.sLUrl.focus();
			return false;
		}
	}
</script>
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<form name="frmSubmit" method="post" action="<%=imgFingers%>/linkweb/processBanner.asp" enctype="multipart/form-data" onsubmit="return jsSubmit(this,'<%=sType%>');">
<input type="hidden" name="sType" value="<%=sType%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="adminId" value="<%=Session("ssBctId")%>">
<input type="hidden" name="iMaxSize" value="<%=iMaxSize%>">
<input type="hidden" name="iBId" value="<%=iBId%>">
<input type="hidden" name="oldImg" value="<%=sImg%>">
	<tr>
		<td colspan="2" bgcolor="#F4F4F4" height="30">&nbsp;<b>배너 신규등록</b></td>
	</tr>
	<tr>
		<td width="120" align="center" bgcolor="#E6E6E6">배너 위치</td>
		<td bgcolor="#FFFFFF">
			<select name="selLoc">
			<option value="">--선택--</option>
			<%Call sbOptCommCd(sCommCd,"I000")%>
			</select>
		</td>
	</tr>		
	<tr>
		<td  align="center" bgcolor="#E6E6E6">배너 이미지</td>
		<td bgcolor="#FFFFFF">
		<%IF sImg <> "" THEN%>
			<img src="<%=simgUrl&sCommCd&"/"&sImg%>">
		<%END IF%>	
		<input type="file" name="sImg"> (10MB 이하 jpg, gif) 
		</td>
	</tr>
	<tr>
		<td  align="center" bgcolor="#E6E6E6"> Size </td>
		<td bgcolor="#FFFFFF">
		Width:
		<input type="text" name="sW" size="3" value="<%=sWidth%>"> px
		, Height:
		<input type="text" name="sH" size="3" value="<%=sHeight%>"> px
		</td>
	</tr>
	<tr>
		<td  align="center" bgcolor="#E6E6E6">링크 URL</td>
		<td bgcolor="#FFFFFF"><input type="text" name="sLUrl" size="50" maxlenght="80" value="<%=sLinkUrl%>"></td>
	</tr>
	<tr>
		<td align="center" bgcolor="#E6E6E6">링크 Target</td>
		<td  bgcolor="#FFFFFF"><input type="text" name="sWD" size="10" maxlenght="10" value="<%=sWindow%>"></td>
	</tr>
	<tr>
		<td  align="center" bgcolor="#E6E6E6">사용여부</td>
		<td bgcolor="#FFFFFF"><input type="checkbox" name="chkUse" <%IF sIsUsing <> "N" THEN%>checked<%END IF%>> 사용</td>
	</tr>
	<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
	<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="location.href='index.asp'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->