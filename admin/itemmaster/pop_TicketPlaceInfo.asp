<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/items/ticketItemCls.asp"-->
<%

dim pidx

pidx  = requestCheckvar(request("ticketPlaceIdx"),10)
if (pidx="") then pidx=0

'==============================================================================

Dim oticketItem
set oticketItem = new CTicketPlace
oticketItem.FRectTicketPlaceIdx = pidx
oticketItem.GetOneTicketPLace

Dim brd_content , parkingGuide
brd_content = oticketItem.FOneItem.FplaceContents
parkingGuide = oticketItem.FOneItem.FparkingGuide
%>
<script type="text/javascript">
function regImg(sFolder, sImg, sName, sSpan){
    var popWin = window.open('/admin/itemmaster/pop_TicketPlace_uploadimg.asp?yr=<%= Left(now(),4) %>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
    popWin.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}

function saveContents(frm){
    if (frm.ticketPlaceName.value.length<1){
        alert('공연장소 를 입력하세요.');
        frm.ticketPlaceName.focus();
        return;
    }
    
    if (frm.tPAddress.value.length<1){
        alert('공연장 주소 를 입력하세요.');
        frm.tPAddress.focus();
        return;
    }
    
    if (frm.placeImg.value.length<1&&frm.brd_content.value.length<1){
        alert('약도 이미지 또는 공연장 설명을 입력하세요.');
        frm.placeImg.focus();
        return;
    }
	
    if(confirm('저장 하시겠습니까?')){
        frm.submit();
    }
    
}
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <form name="frmContents" method="post" action="ticketItem_process.asp">
  <input type="hidden" name="mode" value= "ticketPlace"> 
  <input type="hidden" name="ticketPlaceIdx" value= "<%= pidx %>">
  <tr align="left" bgcolor="F4F4F4">
    <td height="30" colspan="4">
    공연장소 정보
    </td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">* 공연장소 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	<input type="text" name="ticketPlaceName" value="<%= oticketItem.FOneItem.FticketPlaceName %>" size="64" class="text" maxlength="64" >
  	<br>(ex 국립중앙박물관,지산 포레스트 리조트  등)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">* 공연장 주소 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	<input type="text" name="tPAddress" value="<%= oticketItem.FOneItem.FtPAddress %>" size="100" class="text" maxlength="200" >
  	<br>(ex 서울시 강남구 대치동 1002 코스모타워 3층 등)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">공연장 전화 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	<input type="text" name="tPTel" value="<%= oticketItem.FOneItem.FtPTel %>" size="16" class="text" maxlength="16" >
  	(ex 02-000-0000 등)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">홈페이지 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	<input type="text" name="tPHomeURL" value="<%= oticketItem.FOneItem.FtPHomeURL %>" size="60" class="text" maxlength="100" >
  	(ex http://www.jisanresort.co.kr 등)
  	</td>
  </tr>
  <!--
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">약도 link URL :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	<input type="text" name="placeLinkURL" value="<%= oticketItem.FOneItem.FplaceLinkURL %>" size="60" class="text" maxlength="100" >
  	</td>
  </tr>
  -->
  
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">*약도 이미지 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	<input type="text" name="placeImg" value="<%= oticketItem.FOneItem.FplaceImgURL %>" size="60" class="text" maxlength="100" readOnly >
  	<input type="button" value="약도 이미지 등록" onClick="regImg('placeImg','<%= oticketItem.FOneItem.FplaceImgURL %>','placeImg','spanplaceImg');">
  	<div id="spanplaceImg" style="padding: 5 5 5 5">
  	    <% if oticketItem.FOneItem.FplaceImgURL<>"" then %>
		<img  src="<%= oticketItem.FOneItem.FplaceImgURL %>" border="0">
		<a href="javascript:jsDelImg('placeImg','spanplaceImg');"><img src="/images/icon_delete2.gif" border="0"></a>
		<% end if %>
	</div>
					   			
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">관련 이미지1 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	<input type="text" name="contentsImage1" value="<%= oticketItem.FOneItem.FplacecontentsImage1 %>" size="60" class="text" maxlength="100" readOnly >
  	<input type="button" value="관련이미지1등록" onClick="regImg('contentsImage1','<%= oticketItem.FOneItem.FplacecontentsImage1 %>','contentsImage1','spanpcontentsImage1');">
  	<div id="spancontentsImage1" style="padding: 5 5 5 5">
  	    <% if oticketItem.FOneItem.FplacecontentsImage1<>"" then %>
		<img  src="<%= oticketItem.FOneItem.FplacecontentsImage1 %>" border="0">
		<a href="javascript:jsDelImg('contentsImage1','spancontentsImage1');"><img src="/images/icon_delete2.gif" border="0"></a>
		<% end if %>
	</div>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">관련 이미지2 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	<input type="text" name="contentsImage2" value="<%= oticketItem.FOneItem.FplacecontentsImage2 %>" size="60" class="text" maxlength="100" readOnly >
  	<input type="button" value="관련이미지2등록" onClick="regImg('contentsImage2','<%= oticketItem.FOneItem.FplacecontentsImage2 %>','contentsImage2','spanpcontentsImage2');">
  	<div id="spancontentsImage2" style="padding: 5 5 5 5">
  	    <% if oticketItem.FOneItem.FplacecontentsImage2<>"" then %>
		<img  src="<%= oticketItem.FOneItem.FplacecontentsImage2 %>" border="0">
		<a href="javascript:jsDelImg('contentsImage2','spancontentsImage2');"><img src="/images/icon_delete2.gif" border="0"></a>
		<% end if %>
	</div>
	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">관련 이미지3 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	<input type="text" name="contentsImage3" value="<%= oticketItem.FOneItem.FplacecontentsImage3 %>" size="60" class="text" maxlength="100" readOnly >
  	<input type="button" value="관련이미지3등록" onClick="regImg('contentsImage3','<%= oticketItem.FOneItem.FplacecontentsImage3 %>','contentsImage3','spanpcontentsImage3');">
  	<div id="spancontentsImage3" style="padding: 5 5 5 5">
  	    <% if oticketItem.FOneItem.FplacecontentsImage3<>"" then %>
		<img  src="<%= oticketItem.FOneItem.FplacecontentsImage3 %>" border="0">
		<a href="javascript:jsDelImg('contentsImage3','spancontentsImage3');"><img src="/images/icon_delete2.gif" border="0"></a>
		<% end if %>
	</div>
	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">공연장 설명 </td>
  	<td bgcolor="#FFFFFF" colspan="3">
  		<textarea name="brd_content" class="textarea" style="width:98%; height:400px;"><%=brd_content%></textarea>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">주차 안내 </td>
  	<td bgcolor="#FFFFFF" colspan="3">
  		<textarea name="parkingguide" class="textarea" style="width:98%; height:200px;"><%=parkingGuide%></textarea>
  	</td>
  </tr>
  <tr>
    <td colspan="4" height="30" align="center" bgcolor="#FFFFFF" >
        <input type="button" value=" 저 장 " onclick="saveContents(frmContents);">
    </td>
  </tr>
  </form>
</table>

<%
set oticketItem = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->