<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  히치하이커 어드민 프리뷰내 Iframe 페이지
' History : 2014.08.04 유태욱 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/hitchhiker_previewCls.asp"-->
<%
Dim idx, olist, page, i
	idx = request("idx")
	page = request("page")

If page = "" Then page = 1
	
SET olist = new CHitchhikerPreview
	olist.FCurrPage		= page
	olist.FPageSize		= 10
	olist.FrectIdx		= idx
	olist.FrectDevice	= "W"
	olist.sbpreviewDetaillist
%>

<script language="javascript">

var ichk;
ichk = 1;
function jsChkAll(){
	var frm, blnChk;
	frm = document.fitem;
	if(!frm.chkI) return;
	if ( ichk == 1 ){
		blnChk = true;
		ichk = 0;
	}else{
		blnChk = false;
		ichk = 1;
	}
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];
		if ((e.type=="checkbox")) {
			e.checked = blnChk ;
		}
	}
}
//이미지 등록(WWW)
function jsSetImg(idx, sFolder, sImg, sName, sSpan, device){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/admin/hitchhiker/preview/pop_hitchhikerpreview_uploadimg.asp?idx='+idx+'&mode=NEW&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan+'&device='+device,'popImg','width=370,height=150');
	winImg.focus();
}
//이미지 삭제
function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
//이미지 새창 확대보기
function jsImgView(sImgUrl){
	var wImgView;
	wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	wImgView.focus();
}

//노출순서,사용여부 저장
function jsSortIsusing() {
	var frm;
	var sValue, sortNo, isusing;
	frm = document.fitem;
	sValue = "";
	sortNo = "";
	isusing = "";
	chkSel	= 0;

	if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){
			if(frm.chkI[i].checked) chkSel++;

			if (frm.isusing[i].value ==''){
				alert('사용여부를 선택하세요.');
				frm.isusing[i].focus();
				return;
			}
			if (frm.chkI[i].checked){
				if (sValue==""){
					sValue = frm.chkI[i].value;
				}else{
					sValue =sValue+","+frm.chkI[i].value;
				}
				// 노출순서
				if (sortNo==""){
					sortNo = frm.sortNo[i].value;
				}else{
					sortNo =sortNo+","+frm.sortNo[i].value;
				}
				// 사용여부
				if (isusing==""){
					isusing = frm.isusing[i].value;
				}else{
					isusing =isusing+","+frm.isusing[i].value;
				}
			}
		}
	}else{
		if(frm.chkI.checked) chkSel++;
		if(frm.chkI.checked){
			sValue = frm.chkI.value;
			if(!IsDigit(frm.sortNo.value)){
				alert("순서지정은 숫자만 가능합니다.");
				frm.sortNo.focus();
				return;
			}
			sortNo 	=  frm.sortNo.value;
			isusing =  frm.isusing.value;
		}
	}
	if(chkSel<=0) {
		alert("선택한 이미지가 없습니다.");
		return;
	}
	document.frmSortIsusing.sortnoarr.value = sortNo;
	document.frmSortIsusing.isusingarr.value = isusing;
	document.frmSortIsusing.detailidxarr.value = sValue;
	document.frmSortIsusing.submit();
}

//사용여부 전체 조작
function jsIsusingChg(selv) {
    var frm, blnChk;
	frm = document.fitem;
	if (frm.chkI.length > 1){
		for (var i=0;i<frm.isusing.length;i++){
			frm.isusing[i].value=selv;
		}
	}else{
		frm.isusing.value=selv;
	}
}

//페이징
function gosubmit(page){
    var frm = document.frm;
    frm.page.value=page;
	frm.submit();
}

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="frm" method="get" action="" style="margin:0px;">
	<input type="hidden" name="idx" value="<%=idx%>">
	<input type="hidden" name="page" value="<%=page%>">
</form>
<form name="frmSortIsusing" method="post" action="/admin/hitchhiker/preview/hitchhiker_preview_sortIsusing_proc.asp" style="margin:0px;">
	<input type="hidden" name="sortnoarr" value="">
	<input type="hidden" name="isusingarr" value="">
	<input type="hidden" name="idx" value="<%=idx%>">
	<input type="hidden" name="detailidxarr" value="">
	<input type="hidden" name="device" value="W">
	<input type="hidden" name="mode" value="sortisusingedit">
</form>
<tr>
	<td align="left">
		<input class="button" type="button" id="btnEditSel" value="노출순서,사용여부 수정" onClick="jsSortIsusing();">
		<font color="red">※사용여부 및 노출순서를 수정하신 후에 버튼을 눌러주셔야 저장 및 반영이 완료됩니다.</font>
	</td>
	<td align="right">
		<input type="button" name="btnBan" value="Preview이미지등록" onClick="jsSetImg('<%=idx%>','preview','','imgU','spanimgU','W')" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fitem" method="post" style="margin:0px;">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%=olist.FTotalCount %></b>&nbsp;
		페이지 : <b><%= page %>/ <%= olist.FTotalPage %></b>		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="20"><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>
	<td>상세번호</td>
	<td>이미지</td>
	<td>노출순서</td>
	<td>
		전체사용여부
		<select name="selisusing" onchange="jsIsusingChg(this.value)" class="select">
			<option value="N">N</option>
			<option value="Y">Y</option>
		</select>
		<font color="red">(N->삭제)</font>
	</td>
</tr>

<% If olist.FResultCount > 0 Then %>
	<% For i = 0 to olist.Fresultcount -1 %>
		<% if olist.FItemList(i).FIsusing="Y" then %>
			<tr height="25" bgcolor="FFFFFF" align="center">
		<% else %>
			<tr height="25" bgcolor="f1f1f1" align="center">
		<% end if %>
		<td><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= olist.FItemlist(i).Fdetailidx %>"></td>
		<td><%= olist.FItemlist(i).Fdetailidx %></td>
		<td>
			<img src="<%=uploadUrl%>/hitchhiker/preview/detail/<%= olist.FItemlist(i).Fpreviewimg %>" width="50" height="50" onClick="jsImgView('<%=uploadUrl%>/hitchhiker/preview/detail/<%=olist.FItemlist(i).Fpreviewimg%>')" style="cursor:pointer" >
		</td>
		<td><input type="text" size="2" maxlength="2" name="sortNo" value="<%=olist.FItemlist(i).Fsortnum%>" class="text"></td>
		<td>
			<input type="hidden" value="<%=olist.FItemList(i).FIsusing%>" name="orgisusing">
			<% drawSelectBoxUsingYN "isusing", olist.FItemlist(i).FIsusing %>
		</td>
	</tr>
	<% Next %>
	<tr height="25" bgcolor="FFFFFF" >
		<td colspan="15" align="center">
	       	<% If olist.HasPreScroll Then %>
				<span class="olist_link"><a href="javascript:gosubmit('<%= olist.StartScrollPage-1 %>');">[pre]</a></span>
			<% Else %>
			[pre]
			<% End If %>
			
			<% For i = 0 + olist.StartScrollPage to olist.StartScrollPage + olist.FScrollCount - 1 %>
				<% If (i > olist.FTotalpage) Then Exit for %>
				<% If CStr(i) = CStr(olist.FCurrPage) Then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% Else %>
				<a href="javascript:gosubmit('<%= i %>');" class="olist_link"><font color="#000000"><%= i %></font></a>
				<% End if %>
			<% Next %>
			
			<% If olist.HasNextScroll Then %>
				<span class="olist_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
			<% Else %>
			[next]
			<% End If %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</form>
</table>
<% 
SET olist = nothing 
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->