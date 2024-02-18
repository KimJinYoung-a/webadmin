<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : scm 로그인 페이지 백이미지 관리 
'#############################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"--> 
<!-- #include virtual="/lib/classes/hitchhiker/scmMngCls.asp" -->
 <%
 dim CScmMng
 dim idx, sMode
 dim sfimg, suserid, dregdate, susername
 idx = requestCheckVar(Request("idx"),10) 
 sMode = "I"
 if idx <> "" then
 	sMode ="U"
 	set CScmMng = new ClsScmMng 
 	CScmMng.FRectIdx = idx
 	CScmMng.fnGetScmMngConts
 	sfimg = CScmMng.FImgUrl
 	suserid = CScmMng.Fuserid
 	dregdate = CScmMng.Fregdate
 	susername = CScmMng.Fusername
	set CScmMng = nothing
end if
 %>
 <script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript">
	function jsImgReg(){
		 var popImg = window.open("loginImgMng_regImg.asp?menupos=<%=menupos%>","winImg","width=400, height=150, scrollbars=yes,resizable=yes");
		 popImg.focus();
	}
	
	function jsCancel(){
	location.href = "loginImgMng.asp?menupos=<%=menupos%>";
	}
	
	function jsSubmit(){
		 if(!document.frm.sfimg.value){
		 	alert("파일(Image)을 등록해 주세요");
		 	return;
		 }
		 
		 if(confirm("저장하시겠습니까? scm로그인 화면에 바로 적용됩니다.")){
		 document.frm.submit();
		 }
	} 
	
	function jsDelete(){
		if(confirm("삭제하시겠습니까?")){
			document.frm.hidM.value ="D";
			document.frm.submit();
		} 
	} 
	
	 function jsPreview(){
	 	var sfimg =document.frm.sfimg.value;
	 	var selPV =$("#selPV").val();
	 	var sSize="";
	 	if (selPV=="MN"){
	 		sSize = "width=400, height=600,scrollbars=yes,resizable=yes"
	 	}else if(selPV=="MW"){
	 		sSize = "width=600, height=400,scrollbars=yes,resizable=yes"
	 	}else{
	 		sSize = "width=1024, height=768,scrollbars=yes,resizable=yes"
	 	}
	 	var popView = window.open("/adminIndex_pv.asp?sBGImg="+sfimg,"winV",sSize);
	 	popView.focus();
	 }
</script> 
 <table width="100%" align="center">
	<tr>
		<td></td>
	</tr>
</table>
<form name="frm" method="post" action="loginImgMng_proc.asp">
	<input type="hidden" name="hidM" value="<%=sMode%>">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="idx" value="<%=idx%>">
<table width="600" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=adminColor("tablebg")%>"> 
	<%if sMode ="U" then%>
	<tr>
		<td  align="center" bgcolor="<%=adminColor("tabletop")%>"  width="120"><b>idx</b></td>
		<td  bgcolor="#FFFFFF" colspan="3"><%=idx%></td>
	</tr>	
	<tr>
		<td  align="center" bgcolor="<%=adminColor("tabletop")%>"><b>작성자</b></td>
		<td  bgcolor="#FFFFFF"><%=susername%>(<%=suserid%>)</td> 
		<td  align="center" bgcolor="<%=adminColor("tabletop")%>" width="120"><b>작성일</b></td>
		<td  bgcolor="#FFFFFF"><%=dregdate%></td>
	</tr>	
	<%end if%>
	<tr>
		<td  align="center" bgcolor="<%=adminColor("tabletop")%>"><b>배경화면 이미지</b></td>
		<td bgcolor="#FFFFFF" colspan="3"><input type="button" class="button" value="파일(Image)등록" onClick="jsImgReg();"> </td> 
	</tr>
	<tr>
		<td colspan="4" bgcolor="#FFFFFF"><input type="hidden" name="sfimg" id="sfimg"  value="<%=sfimg%>">
			<div id="dvFUrl"><%=sfimg%></div> 
		</td>
	</tr>
</table> 
</form>
 <table width="600" cellpadding="3" cellspacing="1" border="0" style="padding-top:20px;">
	<tr>
		<td width="200"> 
			<select name="selPV" id="selPV" class="select">
			<option value="PW">PC WEB</option>
			<option value="MN">M(normal)</option>
			<option value="MW">M(width mode)</option>
			</select>
			<input type="button" value="미리보기" class="button" onClick="jsPreview();">
		</td>
		<td><input type="button" value="저장" class="button" style="width:80px;color:red;" onClick="jsSubmit();">
			<%if sMode ="U" then%>
			<input type="button" value="삭제" class="button" style="width:80px;color:blue;" onClick="jsDelete();">
			<%end if%>
			<input type="button" value="취소" class="button"  style="width:80px;" onClick="jsCancel();">
			</td>
	</tr>
</table> 

<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
