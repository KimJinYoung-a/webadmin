<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/street/shopcls.asp"-->
<%
Dim idx, olist, page, i, makerid, ocollection
	page = requestCheckVar(request("page"),10)
	idx = requestCheckVar(request("idx"),10)
	makerid = requestCheckVar(request("makerid"),50)
	
If page = "" Then page = 1
	
makerid = session("ssBctID")
	
SET olist = new ccollection
	olist.FCurrPage		= page
	olist.FPageSize		= 10
	olist.FrectIdx			= idx
	olist.frectmakerid = makerid
	olist.sbcollectionitemlist

SET ocollection = new ccollection
	ocollection.FrectIdx = idx
	ocollection.frectmakerid = makerid
	
	if idx <> "" then
		ocollection.sbcollectionmodify
	end if	
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

function jsImgView(sImgUrl){
	var wImgView;

	wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	wImgView.focus();
}

//순서,사용여부 저장
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
			if(!IsDigit(frm.sortNo[i].value)){
				alert("순서지정은 숫자만 가능합니다.");
				frm.sortNo[i].focus();
				return;
			}
			if (frm.chkI[i].checked){
				if (sValue==""){
					sValue = frm.chkI[i].value;
				}else{
					sValue =sValue+","+frm.chkI[i].value;
				}
				// 정렬순서
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
		alert("선택한 상품이 없습니다.");
		return;
	}
	
	var state = '<%= ocollection.FOneItem.Fstate %>';
	var message;
	if (state=='7'){
		message = '오픈상태에서 수정을 하실경우, 상태가 등록중 상태가 되며,\n텐바이텐에 승인요청을 하셔야 합니다.\n\n저장하시겠습니까?';
	}else{
		message = '저장하시겠습니까?';
	}

	if(confirm(message)){
		document.frmSortIsusing.detailidxarr.value = sValue;
		document.frmSortIsusing.sortnoarr.value = sortNo;
		document.frmSortIsusing.isusingarr.value = isusing;
		document.frmSortIsusing.submit();
	}
}

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

function gosubmit(page){
    var frm = document.frm;
    frm.page.value=page;
	frm.submit();
}

function pop_collectionitemreg(idx){
	var state = '<%= ocollection.FOneItem.Fstate %>';
	var message;
	if (state=='7'){
		message = '오픈상태에서 수정을 하실경우, 상태가 등록중 상태가 되며,\n텐바이텐에 승인요청을 하셔야 합니다.\n\n저장하시겠습니까?';
	}else{
		message = '저장하시겠습니까?';
	}

	if(confirm(message)){
		var pop_collectionitemreg = window.open('/designer/brand/shop/collection/pop_collection_itemAddInfo.asp?idx='+idx+'&makerid=<%=makerid%>','pop_collectionitemreg','width=1024,height=768,scrollbars=yes,resizable=yes');
		pop_collectionitemreg.focus();
	}
}

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="frm" method="get" action="" style="margin:0px;">
	<input type="hidden" name="page" value="<%=page%>">
	<input type="hidden" name="idx" value="<%=idx%>">
</form>
<form name="frmSortIsusing" method="post" action="/designer/brand/shop/collection/collectionsortIsusingProcess.asp" style="margin:0px;">
	<input type="hidden" name="detailidxarr" value="">
	<input type="hidden" name="sortnoarr" value="">
	<input type="hidden" name="isusingarr" value="">
	<input type="hidden" name="idx" value="<%=idx%>">
	<input type="hidden" name="mode" value="sortisusingedit">
</form>
<tr>
	<td align="left">
		<input class="button" type="button" id="btnEditSel" value="순서/사용여부 수정" onClick="jsSortIsusing();">
		&nbsp;&nbsp;
		※노출순서&사용여부를 설정하신 후에 버튼을 눌러주셔야 저장 및 반영이 완료됩니다.
	</td>
	<td align="right">
		<input type="button" name="btnBan" value="상품등록" onClick="pop_collectionitemreg('<%= idx %>')" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fitem" method="post" style="margin:0px;">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%=olist.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= olist.FTotalPage %></b>
		&nbsp;&nbsp;
		<font color="red"><b>※총 40개 상품까지 등록 가능 합니다.</b></font>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="20"><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>
	<td>이미지</td>
	<td>상품코드</td>
	<td>판매여부</td>
	<td>노출순서</td>
	<td>
		사용여부
		<select name="selisusing" onchange="jsIsusingChg(this.value)" class="select">
			<option value="N">N</option>
			<option value="Y">Y</option>
		</select>	
	</td>
</tr>
<% If olist.FResultCount > 0 Then %>
<% For i = 0 to olist.fresultcount -1 %>
<% if olist.FItemList(i).FIsusing="Y" then %>
	<tr height="25" bgcolor="FFFFFF" align="center">
<% else %>
	<tr height="25" bgcolor="f1f1f1" align="center">
<% end if %>
	<td><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= olist.FItemlist(i).fdetailidx %>"></td>
	<td>
		<img src="<%= olist.FItemlist(i).FimageSmall %>" width="50" height="50" onClick="jsImgView('<%=olist.FItemlist(i).FimageSmall%>')" style="cursor:pointer" >
	</td>	
	<td><%= olist.FItemlist(i).fitemid %></td>
	<td><%= olist.FItemlist(i).fsellyn %></td>
	<td><input type="text" size="2" maxlength="2" name="sortNo" value="<%=olist.FItemlist(i).FSortNo%>" class="text"></td>
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
set ocollection = nothing
SET olist = nothing
%>
<!-- #include virtual="/designer/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->