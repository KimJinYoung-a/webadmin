<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이벤트 그룹등록
' History : 2010.09.28 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/Event_cls.asp"-->

<script language="javascript">

 function jsGroupSubmit(frm){
 	if(!frm.sGD.value){
 	alert("그룹명을 입력해주세요");
 	return false;
 	}
 } 


</script>

<%
Dim eCode : eCode = RequestCheckvar(Request("eC"),10)
Dim eGCode : eGCode = RequestCheckvar(Request("eGC"),10)
Dim cEGroup, arrP,intP,sM
Dim gpcode, gdesc, gsort, gimg,gdepth,gpdesc,glink
Dim arrImg, slen, sImgName
 set cEGroup = new ClsEventGroup
 	cEGroup.FECode = eCode
  	arrP = cEGroup.fnGetRootGroup
  	sM = "I"
  	IF (eGCode <> "" and eGCode <> "0" and not isnull(eGCode)) THEN
	  	cEGroup.FEGCode = eGCode
	  	cEGroup.fnGetEventItemGroupCont	
	  		
	  	gpcode 	= cEGroup.FGPCode
	  	gdesc  	= cEGroup.FGDesc
	  	gsort	= cEGroup.FGSort
	  	gimg	= cEGroup.FGImg
	  	gdepth	= cEGroup.FGDepth
	  	gpdesc  = cEGroup.FGPDesc
	  	glink	= cEGroup.FGlink
	  	sM = "U"
	END IF  	
 set cEGroup = nothing
 
 
IF gimg <> "" THEN
	arrImg = split(gimg,"/")
	slen = ubound(arrImg)
	sImgName = arrImg(slen)
END IF	

if gsort = "" then gsort = 0

%>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 이벤트 그룹 등록</div>
<table width="580" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
<form name="frmG" method="post" action="<%= imgFingers %>/linkweb/eventgroup_process.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsGroupSubmit(this);">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="eGC" value="<%=eGCode%>">
<input type="hidden" name="mode" value="<%=sM%>">
<input type="hidden" name="sOGimg" value="<%=gimg%>">
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">상위그룹</td>
				<td bgcolor="#FFFFFF">
				<%IF gdepth = "" THEN%>
				<select name="selPC">
				<option value="0">최상위</option>
				<%IF isArray(arrP) THEN
					For intP =0 To UBound(arrP,2)
					%>
				<option value="<%=arrP(0,intP)%>" <%IF Cstr(gpcode) = CStr(arrP(0,intP)) THEN%>selected<%END IF%>><%=arrP(1,intP)%></option>	
			<%  Next
				END IF%>	
				</select>
				<%ELSE%>
				<input type="hidden" name="selPC" value="<%=gpcode%>">
				<%=gpdesc%>
				<%END IF%>
				</td>
			</tr>
			<tr>
				<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">그룹명</td>
				<td bgcolor="#FFFFFF"><input type="text" name="sGD" size="20" value="<%=db2html(gdesc)%>"></td>
			</tr>		
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">정렬순서</td>
				<td bgcolor="#FFFFFF"><input type="text" size="2" name="sGS"  value="<%=gsort%>"></td>
			</tr>	
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">이미지</td>
				<td bgcolor="#FFFFFF"><input type="file" name="sGimg"><br><%IF gimg <> "" THEN%><%=sImgName%> <input type="checkbox" name="delI">삭제<%END IF%></td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">맵링크</td>
				<td bgcolor="#FFFFFF">				
					<font color="red">+ 꼭! 맵코딩은 맵명칭을 뺀 &lt;area shape="rect" ~ 만 입력해주세요. </font><br>
					<font color="blue">이벤트 그룹 페이지로 링크시<br>
					&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoEventGroupMain('<font color="blue">이벤트코드</font>','<font color="blue">그룹코드</font>');" onfocus="this.blur();"&gt;<br>		    						
		   			<input type="text" value="<map name='mapGroup'>" style="border:0" size="30"><br>
						<textarea name="tGMap" rows="13" cols="60"><%=db2html(glink)%></textarea>  	
					<input type="text" value="</map>" style="border:0">	
					
				</td>
			</tr>	
		</table>
	</td>		
</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right" height="40">
		<input type="image" src="/images/icon_confirm.gif">
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>	
</form>	
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->