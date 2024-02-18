<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/pop_eventitem_groupImage.asp
' Description :  이벤트 그룹 이미지 수정
' History : 2007.02.22 정윤정 생성
'			2015.02.12 정윤정 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v3.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"--> 
<%
Dim eCode : eCode = requestCheckVar(Request("eC"),10)
Dim eGCode : eGCode = requestCheckVar(Request("eGC"),10) 
Dim vYear : vYear = requestCheckVar(Request("yr"),4)
dim eChannel : eChannel = requestCheckVar(Request("eCh"),1)
Dim cEGroup, arrP,intP, sM
Dim gpcode, gdesc, gsort, gimg,gdepth,gpdesc,glink, gdisp
Dim arrImg, slen, sImgName, eFolder, cEvtCont, eregdate

 gdisp = True

IF eCode <> "" THEN
    set cEvtCont = new ClsEvent
    cEvtCont.FECode = eCode	'이벤트 코드
	cEvtCont.fnGetEventCont
	eregdate = cEvtCont.FERegdate
    set cEvtCont = nothing
end if 

 set cEGroup = new ClsEventGroup
 	cEGroup.FECode = eCode
 	cEGroup.FEChannel = eChannel
 	cEGroup.FGDisp = 1
  	arrP = cEGroup.fnGetRootGroup
  	sM = "GU"
  	IF (eGCode <> "" and eGCode <> "0" and not isnull(eGCode)) THEN
	  	cEGroup.FEGCode = eGCode
	  	cEGroup.fnGetEventItemGroupCont		
	  	gpcode 	= cEGroup.FGPCode
	  	gdesc  	= cEGroup.FGDesc
	  	gsort	= cEGroup.FGSort
	  	gdepth	= cEGroup.FGDepth
	  	gpdesc  = cEGroup.FGPDesc 
		gimg	= cEGroup.FGImg  
	  	glink	= cEGroup.FGlink
		gdisp  = cEGroup.FGDisp 
	END IF  	
 set cEGroup = nothing
 
 
IF gimg <> "" THEN
	arrImg = split(gimg,"/")
	slen = ubound(arrImg)
	sImgName = arrImg(slen)
END IF	

if gsort = "" then gsort = 0
eFolder = eCode
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script>
$(function(){
    // 창 리사이즈시 testarea 높이 조정
    $(window).resize(function() { 
        $('#tGMap').css('height', $(window).height()-340); 
    }); 
});

function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(eregdate)%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
		eval("document.all."+sName).value = "";
		eval("document.all."+sSpan).style.display = "none";
	}
}
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 이벤트 그룹 이미지 등록</div>
<form name="frmG" method="post" action="eventgroup_process.asp">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="eGC" value="<%=eGCode%>">
<input type="hidden" name="mode" value="<%=sM%>">
<input type="hidden" name="sOGimg" value="<%=gimg%>">
<input type="hidden" name="yr" value="<%=vYear%>">
<input type="hidden" name="eCh" value="<%=eChannel%>">
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr> 
				<%IF eChannel ="M" then%>
				<td bgcolor="#e3f1fb" align="center"  colspan="2"><b>Mobile / App</b></td>
				<%ELSE%>
				<td bgcolor="#FAECC5" align="center" colspan="2"><b>PC-WEB</b></td>
				<%END IF%>
			</tr>
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
				<td bgcolor="#FFFFFF"><input type="text" name="sGD" size="40" value="<%=db2html(gdesc)%>" maxlength="32"></td>
			</tr>		
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">정렬순서</td>
				<td bgcolor="#FFFFFF"><input type="text" size="2" name="sGS"  value="<%=gsort%>"></td>
			</tr> 
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">전시여부</td>
				<td bgcolor="#FFFFFF"><input type="radio" name="eIsDisp" value="1" <%if gdisp then%>checked<%end if%>>Y <input type="radio" name="eIsDisp" value="0" <%if not gdisp then%>checked<%end if%>>N </td>
			</tr> 
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">이미지</td>
				<td bgcolor="#FFFFFF">
					<input type="button" name="etcitem" value="파일선택" onClick="jsSetImg('<%=eFolder%>','<%=gimg%>','sGimg','gbanner')" class="button">
					<div id="gbanner" style="padding: 5 5 5 5">
						<%IF gimg <> "" THEN %>
						<img src="<%=gimg%>" border="0">
						<a href="javascript:jsDelImg('sGimg','gbanner');"><img src="/images/icon_delete2.gif" border="0"></a>
						<%END IF%>
					</div>
					<input type="hidden" name="sGimg" value="<%=gimg%>">
					<p style="color:red;font-size:11px;">[ 권장 이미지 : JPEG, 60%, 정방형 750px × 750px ]</p>
				</td>
			</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">맵링크</td>
				<td bgcolor="#FFFFFF">				
					<font color="red">+ 꼭! 맵코딩은 맵명칭을 뺀 &lt;area shape="rect" ~ 만 입력해주세요. </font><br>
					<font color="blue">이벤트 그룹 이동시<br>
					&lt;a href="#mapGroup그룹코드">홈인테리어&lt;/a&gt; <br><br>
					<font color="blue">이벤트 그룹 페이지로 링크시<br>
					&lt;area shape="rect" coords="0,0,0,0" onclick="TnGotoEventGroupMain('<font color="blue">이벤트코드</font>','<font color="blue">그룹코드</font>');" onfocus="this.blur();"&gt;<br><br>
					<font color="blue">GNB 보드 링크시 아래스크립트 사용 (모바일웹/웹뷰 공통)<br>
					&lt;a href= "/event/eventmain.asp?eventid=이벤트코드" onclick="jsEventlinkURL(<font color="blue">이벤트코드</font>);return false;"&gt;&nbsp;&lt/a&gt;<br>		    
		   			<div style="padding-right:10px;">
			   			<input type="text" value="<map name='mapGroup'>" style="border:0" size="30"><br>
							<textarea id="tGMap" name="tGMap" style="width:100%;height:280px;"><%=db2html(glink)%></textarea>  	
						<input type="text" value="</map>" style="border:0">	
					</div>
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
</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->