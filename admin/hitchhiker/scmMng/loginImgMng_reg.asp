<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : scm �α��� ������ ���̹��� ���� 
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
		 	alert("����(Image)�� ����� �ּ���");
		 	return;
		 }
		 
		 if(confirm("�����Ͻðڽ��ϱ�? scm�α��� ȭ�鿡 �ٷ� ����˴ϴ�.")){
		 document.frm.submit();
		 }
	} 
	
	function jsDelete(){
		if(confirm("�����Ͻðڽ��ϱ�?")){
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
		<td  align="center" bgcolor="<%=adminColor("tabletop")%>"><b>�ۼ���</b></td>
		<td  bgcolor="#FFFFFF"><%=susername%>(<%=suserid%>)</td> 
		<td  align="center" bgcolor="<%=adminColor("tabletop")%>" width="120"><b>�ۼ���</b></td>
		<td  bgcolor="#FFFFFF"><%=dregdate%></td>
	</tr>	
	<%end if%>
	<tr>
		<td  align="center" bgcolor="<%=adminColor("tabletop")%>"><b>���ȭ�� �̹���</b></td>
		<td bgcolor="#FFFFFF" colspan="3"><input type="button" class="button" value="����(Image)���" onClick="jsImgReg();"> </td> 
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
			<input type="button" value="�̸�����" class="button" onClick="jsPreview();">
		</td>
		<td><input type="button" value="����" class="button" style="width:80px;color:red;" onClick="jsSubmit();">
			<%if sMode ="U" then%>
			<input type="button" value="����" class="button" style="width:80px;color:blue;" onClick="jsDelete();">
			<%end if%>
			<input type="button" value="���" class="button"  style="width:80px;" onClick="jsCancel();">
			</td>
	</tr>
</table> 

<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
