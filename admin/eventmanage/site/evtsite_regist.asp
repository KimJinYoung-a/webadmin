<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event_regist.asp
' Description :  이벤트 개요 등록
' History : 2007.02.07 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventSiteCls.asp"-->
<%
Dim siteidx, eMode
Dim cEvtSiteCont
Dim slocation, stype, scont,sltype, slink, sw, sh, sdo, susing, strlink, strMap
Dim sFolder

siteidx = Request.Querystring("idx")
strlink=""
strMap =""
IF siteidx = "" THEN
	 eMode = "I"	 
ELSE
	eMode = "U"
	Set cEvtSiteCont = new ClsEvtSite
	 cEvtSiteCont.FSIdx = siteidx
	 cEvtSiteCont.fnGetContent
	 
	 slocation = cEvtSiteCont.FSLocation
	 stype = cEvtSiteCont.FSType
	 scont = cEvtSiteCont.FSCont 
	 sltype = cEvtSiteCont.FSLType
	 slink = cEvtSiteCont.FSLink 
	 sw = cEvtSiteCont.FSW
	 sh = cEvtSiteCont.FSH
	 sdo = cEvtSiteCont.FSDo
	 susing = cEvtSiteCont.FSUsing
	Set cEvtSiteCont = nothing
	
	If sltype = "L" THEN strlink = slink
	If sltype = "M" THEN strMap =  slink 
	
END IF	
'sFolder = "sitemanage/"&slocation

%>
<script language="javascript">
<!--
	function jsSetImg(sFolder, sImg, sName, sSpan){
		document.domain ="10x10.co.kr";
		var winImg,sl;			
		sl = document.frmEvt.sitelocation.value;
		
		winImg = window.open('/admin/eventmanage/common/pop_event_uploadimg.asp?sF=sitemanage/'+sl+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}
	
	function jsDelImg(sName, sSpan){
		if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){		
		   eval("document.all."+sName).value = "";	
		   eval("document.all."+sSpan).style.display = "none";
		}
	}
	
	function jsChangeLink(){
	  var frm;  frm = document.frmEvt;
	  	  
	  if(frm.selLinkType.options[frm.selLinkType.selectedIndex].value == "M") {
	  	document.all.divL.style.display = "none";
	  	document.all.divM.style.display = "";	  
	  }else{
	  	document.all.divL.style.display = "";
	  	document.all.divM.style.display = "none";
	  } 
	}
	
	function jsEvtSubmit(frm){
		if(frm.selType.value=="image"||frm.selType.value=="flash"){
			if(frm.evtImg.value.length!=0){
				if(!frm.sW.value||frm.sW.value==0) {
					alert("이미지 너비를 입력해주세요.");
					return false;
				}
				if(!frm.sH.value||frm.sH.value==0) {
					alert("이미지 높이를 입력해주세요.");
					return false;
				}
			}
		}

	 if(confirm("저장하시겠습니까?")){
	   return;
	 }	 
	 return false;
	}
	
	function jsChTypeView(iVal){
		if(iVal == "text"){
			document.all.divimg.style.display = "none";
		  	document.all.divtxt.style.display = "";
		}else{
			document.all.divimg.style.display = "";
		  	document.all.divtxt.style.display = "none";
		}
	}
//-->
</script>
<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0" >
<form name="frmEvt" method="post" action="evtsite_process.asp" onSubmit="return jsEvtSubmit(this);">
<input type="hidden" name="imod" value="<%=eMode%>">
<input type="hidden" name="evtImg" value="<%=scont%>">
<input type="hidden" name="siteidx" value="<%=siteidx%>">
<tr>
	<td>
		<table width="800" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		   <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">위치</td>
		   		<td bgcolor="#FFFFFF">		   			
		   		<%sbGetOptEventCodeValue "sitelocation",slocation,False,""%>
		   		</td>
		   	</tr>	
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">종류</td>
		   		<td bgcolor="#FFFFFF">
		   			<select name="selType" onChange="jsChTypeView(this.value);">
		   			<option value="image" <%IF stype = "image" THEN%>selected<%END IF%>>image</option>
		   			<option value="flash" <%IF stype = "flash" THEN%>selected<%END IF%>>flash</option>
		   			<option value="text" <%IF stype = "text" THEN%>selected<%END IF%>>text</option>
		   			</select>
		   		</td>
		   	</tr>
		</table>
	</td>			   	
</tr>	
<tr>
	<td><div id="divimg" style="display:;">
		<table width="800" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		   	<tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">이미지</td>
		   		<td bgcolor="#FFFFFF">
		   			<span id="spanImg">
		   				<%IF scont <> "" THEN %>
		   				<img src="<%=scont%>" <%IF Cint(sw) > 200 THEN%>width="200" <%END IF%> <%IF Cint(sh) > 200 THEN%>height="200" <%END IF%>> 
		   				<%END IF%>
		   			</span>
		   			<input type="button" name="btnImg" value="이미지 등록" onClick="javascript:jsSetImg('<%=sFolder%>','','evtImg','spanImg');">
		   		</td>
		   	</tr>
		   	   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">사이즈</td>
		   		<td bgcolor="#FFFFFF">
		   			width : <input type="text" name="sW" size="4" maxlength="4" value="<%=sw%>"> &nbsp; height : <input type="text" name="sH" size="4" maxlength="4" value="<%=sh%>">
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">링크종류</td>
		   		<td bgcolor="#FFFFFF">
		   			<select name="selLinkType" onChange="jsChangeLink();">
		   			<option value="L" <%IF sltype = "L" THEN%>selected<%END IF%>>Link</option>
		   			<option value="M" <%IF sltype = "M" THEN%>selected<%END IF%>>Map</option>
		   			</select>
		   		</td>
		   	</tr>
		 </table>
		 </div>
		 <div id="divtxt" style="display:none;">
		<table width="800" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		   	<tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">텍스트</td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="stxt" size="30" maxlength="60" value="<%=scont%>">
		   		</td>
		   	</tr>
		 </table>
		 </div>
	</td>	   	
</tr>	
<tr>
	<td>
		<table width="800" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">   
		   	<tr>
		   		<td  align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">링크/맵</td>
		   		<td bgcolor="#FFFFFF">
		   			<div id="divL" style="display:;">
		   			<input type="text" name="sL" size="60" value="<%=strlink%>">
		   			</div>
		   			<div id="divM" style="display:none;">
		   			<font color="red">+ 꼭! 맵코딩은 맵명칭을 뺀 &lt;area shape="rect" ~ 만 입력해주세요. </font>
		   			<br>
		   			<input type="text" value="<map name='Map'>" style="border:0"><br>
		   			<textarea name="sM" cols="60" rows="3"><%=strMap%></textarea><br>
		   			<input type="text" value="</map>" style="border:0">
		   			</div>
		   		</td>
		   	</tr>	
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">전시순서</td>
		   		<td  bgcolor="#FFFFFF">		
		   			<input type="text" name="sDO" size="4" maxlength="4" value="<%=sdo%>">
		   		</td>
		   	</tr>	
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">사용유무</td>
		   		<td  bgcolor="#FFFFFF">		
		   			<input type="radio" name="rdoUse" value="Y" <%IF susing="" OR susing="Y" THEN%>checked<%END IF%>>사용함 
		   			<input type="radio" value="N" name="rdoUse"  <%IF susing="N" THEN%>checked<%END IF%>>사용안함
		   		</td>
		   	</tr>	   
		</table>	
	</td>		
</tr>
<tr>
	<td width="800" align="right" height="40">
		<input type="image" src="/images/icon_save.gif"> 
		<a href="index.asp?menupos=<%=menupos%>"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>	
</form>
</table>
 <script language="javascript">
 <!--
 	jsChangeLink();
 //-->
 </script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->