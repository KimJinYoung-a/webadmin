<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' Discription : 웨딩 기획전 등록페이지(모바일)
' History : 2018.04.16 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/wedding_ContentsManageCls.asp" -->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
Dim isusing, fixtype, validdate, prevDate
Dim idx, poscode, reload, gubun, edid
Dim culturecode
	idx = request("idx")
	poscode = request("poscode")
	reload = request("reload")
	gubun = request("gubun")

	isusing = request("isusing")
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

	dim oPlanEvent
		set oPlanEvent = new CWeddingContents
		oPlanEvent.FRectIdx = idx
		oPlanEvent.GetOneKitContents

	If gubun = "" Then
		gubun = "index"
	End If

%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>

	function SaveMainContents(frm){
	    
		if (frm.copy1.value==""){
	        alert('상단 카피를 입력 하세요.');
	        frm.copy1.focus();
	        return;
	    }

		if (frm.copy2.value==""){
	        alert('메인 카피를 입력 하세요.');
	        frm.copy2.focus();
	        return;
	    }

		if (frm.copy3.value==""){
	        alert('하단 카피를 입력 하세요.');
	        frm.copy3.focus();
	        return;
	    }

		if (frm.copy4.value==""){
	        alert('PC 롤오버 카피를 입력 하세요.');
	        frm.copy4.focus();
	        return;
	    }
		
		if (frm.itemid.value==""){
	        alert('아이템 번호를 입력 하세요.');
	        frm.itemid.focus();
	        return;
	    }

		if (frm.upload_img1.value==""){
	        alert('이미지 1번을 업로드 해주세요.');
	        frm.upload_img1.focus();
	        return;
	    }

		if (frm.DispOrder.value==""){
	        alert('뷰 순번을 입력 하세요.');
	        frm.DispOrder.focus();
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

	function jsSetImg(sImg, sName, sSpan){ 
		var winImg;
		var sFolder=document.frmcontents.Evt_Code.value;
		if (sFolder=="")
		{
			alert("이벤트 검색 후 이미지를 등록해주세요.");
		}
		else
		{
		winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
		}
	}

	// 새상품 추가 팝업
	function addnewItem(itemnum){
			var popwin; 
			popwin = window.open("pop_additemlist.asp?itemnum="+itemnum, "popup_item", "width=1024,height=768,scrollbars=yes,resizable=yes");
			popwin.focus();
	}

	function jsDelImg(sName, sSpan){
		if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
		   eval("document.all."+sName).value = "";
		   eval("document.all."+sSpan).style.display = "none";
		}
	}
</script>

<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="doWeddingKitReg.asp" onsubmit="return false;">
<input type="hidden" name="Evt_Code" value="30000">
<input type="hidden" name="idx" value="<%= oPlanEvent.FOneItem.FIdx %>">
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">상단카피</td>
    <td>
        <input type="text" name="copy1" size="50" value="<%=oPlanEvent.FOneItem.FCopy1%>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">메인카피</td>
    <td>
        <input type="text" name="copy2" size="50" value="<%=oPlanEvent.FOneItem.FCopy2%>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">하단카피</td>
    <td>
        <input type="text" name="copy3" size="50" value="<%=oPlanEvent.FOneItem.FCopy3%>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">PC롤오버 카피</td>
    <td>
		<textarea name="copy4" rows="5" cols="50"><%=oPlanEvent.FOneItem.FCopy4%></textarea>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">아이템</td>
    <td>
		<input type="text" name="itemid" value="<%=oPlanEvent.FOneItem.FItemid%>"> <a href="javascript:addnewItem('');">불러오기</a>&nbsp;<span id="iteminfo"></span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">아이템 이미지1 업로드</td>
  <td>
  	<input type="button" name="etcitem1" value="이미지 등록" onClick="jsSetImg('<%=oPlanEvent.FOneItem.FUpload_img1%>','upload_img1','item1')" class="button"> (PC,모바일 겸용 직사각 이미지 없으면 모바일 제외)
	<input type="hidden" name="upload_img1" value="<%=oPlanEvent.FOneItem.FUpload_img1%>">
		<div id="item1" style="padding: 5 5 5 5">
			<%IF oPlanEvent.FOneItem.FUpload_img1 <> "" THEN %>
			<img  src="<%=oPlanEvent.FOneItem.FUpload_img1%>" width="50%" border="0">
			<a href="javascript:jsDelImg('upload_img1','item1');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">아이템 이미지2 업로드</td>
  <td>
  	<input type="button" name="etcitem2" value="이미지 등록" onClick="jsSetImg('<%=oPlanEvent.FOneItem.FUpload_img2%>','upload_img2','item2')" class="button"> (PC 정사각 이미지)
	<input type="hidden" name="upload_img2" value="<%=oPlanEvent.FOneItem.FUpload_img2%>">
		<div id="item2" style="padding: 5 5 5 5">
			<%IF oPlanEvent.FOneItem.FUpload_img2 <> "" THEN %>
			<img  src="<%=oPlanEvent.FOneItem.FUpload_img2%>" width="50%" border="0">
			<a href="javascript:jsDelImg('upload_img2','item2');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">뷰 순서</td>
    <td>
		<input type="text" name="DispOrder" value="<%=oPlanEvent.FOneItem.FDispOrder %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" 저 장 " onClick="SaveMainContents(frmcontents);"></td>
</tr>
</form>
</table>
<%
set oPlanEvent = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
