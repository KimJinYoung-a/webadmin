<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : main_manager.asp
' Discription : 이미지 영역 링크
' History : 2019.08.06 원승현 : 신규작성
'			2022.07.07 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/imageLinkCls.asp" -->
<%
Dim isusing, fixtype, validdate, prevDate
Dim idx, poscode, reload, gubun, edid
Dim culturecode
	idx = requestCheckVar(getNumeric(request("idx")),10)
	poscode = request("poscode")
	reload = request("reload")
	gubun = request("gubun")

	isusing = requestCheckVar(request("isusing"),1)
	fixtype = request("fixtype")
	validdate= request("validdate")
	prevDate = request("prevDate")

	culturecode = request("eC")

	if idx="" then idx=0

	Response.write culturecode

	if reload="on" then
	    response.write "<script>opener.location.reload(); window.close();</script>"
	    dbget.close()	:	response.End
	end if

	dim oLinkContents
		set oLinkContents = new CimageLink
		oLinkContents.FRectIdx = idx
		oLinkContents.GetOneContents

	If gubun = "" Then
		gubun = "index"
	End If

%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript'>

	function SaveMainContents(frm){

		if (frm.title.value==""){
	        alert('타이틀을 입력해주세요.');
	        frm.title.focus();
	        return;
	    }

		if (frm.Link_Image.value==""){
	        alert('이미지를 업로드 하세요.');
	        frm.Link_Image.focus();
	        return;
	    }

	    if (confirm('저장 하시겠습니까?')){
	        frm.submit();
	    }
	}

	function ChangeLinktype(comp){
	    if (comp.value=="M"){
	       document.all.link_M.style.display = "";
	       document.all.link_L.style.display = "none";
	    }else{
	       document.all.link_M.style.display = "none";
	       document.all.link_L.style.display = "";
	    }
	}

	//function getOnLoad(){
	//    ChangeLinktype(frmcontents.linktype.value);
	//}

	//window.onload = getOnLoad;

	function ChangeGubun(comp){
	    location.href = "?gubun=<%=gubun%>&poscode=" + comp.value;
	    // nothing;
	}


	function ChangeGroupGubun(comp){
	    location.href = "?gubun=" + comp.value;
	    // nothing;
	}

	function cultureloadpop(){
		winLast = window.open('pop_culturelist.asp','pLast','width=1200,height=600, scrollbars=yes')
		winLast.focus();
	}

	//색상코드 선택
	function selColorChip(bg,cd) {
		var i;
		document.frmcontents.BGColor.value= bg;
		for(i=1;i<=11;i++) {
			document.all("cline"+i).bgColor='#DDDDDD';
		}
		if(!cd) document.all("cline0").bgColor='#DD3300';
		else document.all("cline"+cd).bgColor='#DD3300';
	}

	//-- jsLastEvent : 지난 이벤트 불러오기 --//
	function jsLastEvent(num){
	  winLast = window.open('pop_event_lastlist.asp?num='+num,'pLast','width=800,height=600, scrollbars=yes')
	  winLast.focus();
	}

	//-- jsImgView : 이미지 확대화면 새창으로 보여주기 --//
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}


	function jsSetImg(sFolder, sImg, sName, sSpan){ 
		var winImg;
		winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	function jsDelImg(sName, sSpan){
		if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
		   eval("document.all."+sName).value = "";
		   eval("document.all."+sSpan).style.display = "none";
		}
	}
</script>

<form name="frmcontents" method="post" action="doLinkImageReg.asp" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="Link_Image" value="<%=oLinkContents.FOneItem.Fimage%>">
<input type="hidden" name="Link_Admin_Image" value="<%=oLinkContents.FOneItem.FRegUserImage%>">
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF" align="center">Idx</td>
    <td>
        <% if oLinkContents.FOneItem.Fidx<>"" then %>
        <%= oLinkContents.FOneItem.Fidx %>
        <input type="hidden" name="idx" value="<%= oLinkContents.FOneItem.Fidx %>">
        <% else %>

        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="200" bgcolor="#DDDDFF" align="center">타이틀</td>
    <td >
		<input type="text" name="title" value="<%= ReplaceBracket(oLinkContents.FOneItem.Ftitle) %>" maxlength="60">
    </td>
</tr>
<tr bgcolor="#FFFFFF" id="tmpstyle1">
	<td bgcolor="#DDDDFF" align="center" width="15%">링크 이미지</td>
	<td><input type="button" name="limg" value="이미지 등록" onClick="jsSetImg('linkimageutil','<%=oLinkContents.FOneItem.FImage%>','Link_Image','newlinkimg')" class="button">
		<div id="newlinkimg" style="padding: 5 5 5 5">
			<%IF oLinkContents.FOneItem.FImage <> "" THEN %>
			<a href="javascript:jsImgView('<%=oLinkContents.FOneItem.FImage%>')"><img  src="<%=oLinkContents.FOneItem.FImage%>" width="400" border="0"></a>
			<a href="javascript:jsDelImg('Link_Image','newlinkimg');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
		<%=oLinkContents.FOneItem.FImage%>
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="tmpstyle1">
	<td bgcolor="#DDDDFF" align="center" width="15%">등록자 사진(브랜드 이미지 또는 등록자 사진)</td>
	<td><input type="button" name="limg" value="이미지 등록" onClick="jsSetImg('linkimageutil','<%=oLinkContents.FOneItem.FRegUserImage%>','Link_Admin_Image','newlinkimgadmin')" class="button">
		<div id="newlinkimgadmin" style="padding: 5 5 5 5">
			<%IF oLinkContents.FOneItem.FImage <> "" THEN %>
			<a href="javascript:jsImgView('<%=oLinkContents.FOneItem.FRegUserImage%>')"><img  src="<%=oLinkContents.FOneItem.FRegUserImage%>" width="400" border="0"></a>
			<a href="javascript:jsDelImg('Link_Admin_Image','newlinkimgadmin');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
		<%=oLinkContents.FOneItem.FRegUserImage%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="200" bgcolor="#DDDDFF" align="center">등록자 프론트 노출이름</td>
    <td >
		<input type="text" name="reguserfrontname" value="<%= ReplaceBracket(oLinkContents.FOneItem.FRegUserFrontName) %>" maxlength="20">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF" align="center">사용여부</td>
  <td>
  	<input type="radio" name="Isusing" value="Y"<% If oLinkContents.FOneItem.FIsusing="Y" Or oLinkContents.FOneItem.FIsusing="" Then Response.write " checked" %>> 사용함
	<input type="radio" name="Isusing" value="N"<% If oLinkContents.FOneItem.FIsusing="N" Then Response.write " checked" %>> 사용안함
  </td>
</tr>
<% If oLinkContents.FOneItem.Fadminid<>"" Then %>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">작업자</td>
  <td>
  	작업자 : <%=oLinkContents.FOneItem.Fadminid %><br>
	최종작업자 : <%=oLinkContents.FOneItem.Flastadminid %>
  </td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" 저 장 " onClick="SaveMainContents(frmcontents);"></td>
</tr>
</table>
</form>

<%
set oLinkContents = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
