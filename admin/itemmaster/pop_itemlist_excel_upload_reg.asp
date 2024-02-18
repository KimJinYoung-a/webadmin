<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  상품신규등록. 엑셀 일괄
' History : 2019.12.13 한용민 생성
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemedit_temp_cls.asp"-->
<%
dim mode, i, failtype, chk_idx, chk_idx_fail
	mode 			= requestCheckVar(request("mode"),32)
	chk_idx 		= request("chk_idx")
	chk_idx_fail 		= request("chk_idx_fail")

dim oCManualMeachul
set oCManualMeachul = new Citemedit_templist
	oCManualMeachul.FPageSize = 1000
	oCManualMeachul.FCurrPage = 1
	oCManualMeachul.FRectRegAdminID = session("ssBctId")
	oCManualMeachul.FRectExcludeRegFinish = "Y"
	oCManualMeachul.GetsuccessitemregList

dim oCFailManualMeachul
set oCFailManualMeachul = new Citemedit_templist
	oCFailManualMeachul.FPageSize = 1000
	oCFailManualMeachul.FCurrPage = 1
	oCFailManualMeachul.FRectRegAdminID = session("ssBctId")
	oCFailManualMeachul.GetregFailList
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript">

function fnChkFile(sFile, arrExt){
    //파일 업로드 유무확인
     if (!sFile){
    	 return true;
    	}

    var blnResult = false;

    //파일 확장자 확인
	var pPoint = sFile.lastIndexOf('.');
	var fPoint = sFile.substring(pPoint+1,sFile.length);
	var fExet = fPoint.toLowerCase();

	for (var i = 0; i < arrExt.length; i++)
	   	{
	    	if (arrExt[i].toLowerCase() == fExet)
	    	{
	   			blnResult =  true;
	   		}
		}

	return blnResult;
}

function XLSumbit(){
	document.domain = '10x10.co.kr';
	var frm = document.frmFile;
    
	arrFileExt = new Array();			
	arrFileExt[arrFileExt.length]  = "xls";     //csv
	
	if (frm.sFile.value==''){
		alert('파일을 입력해 주세요');
		frm.sFile.focus();
		return;
	}
	
	//파일유효성 체크
	if (!fnChkFile(frm.sFile.value, arrFileExt)){
		alert("파일은 xls파일만 업로드 가능합니다.");
		return;
	}

	frm.target='view';
	frm.submit();
}

function CheckAll(chk) {
	for (var i = 0; ; i++) {
		var v = document.getElementById("chk_" + i);
		if (v == undefined) {
			return;
		}

		if (v.disabled != true) {
			v.checked = chk.checked;
		}
	}
}

// 업로드 삭제하기
function delClick(mode) {
	var frm = document.frmdetail;

	if (mode=='delregitem_fail'){
		if ($('input[name="chk_idx_fail"]:checked').length == 0) {
			alert('선택 아이템이 없습니다.');
			return;
		}
	}else{
		if ($('input[name="chk_idx"]:checked').length == 0) {
			alert('선택 아이템이 없습니다.');
			return;
		}
	}
	if (confirm("삭제 하시겠습니까?") == true) {
		frm.mode.value=mode;
		frm.action="/admin/itemmaster/pop_itemlist_excel_upload_process.asp"
		frm.target="view";
		frm.submit();
	}
}

function toggleChecked(status) {
    $('[name="chk_idx"]').each(function () {
        $(this).prop("checked", status);
    });
}
function toggleChecked_fail(status) {
    $('[name="chk_idx_fail"]').each(function () {
        $(this).prop("checked", status);
    });
}

// 실제적용
function saveClick() {
	var frm = document.frmdetail;

	if ($('input[name="chk_idx"]:checked').length == 0) {
		alert('선택 아이템이 없습니다.');
		return;
	}
	if ($('input[name="chk_idx"]:checked').length > 100) {
		alert('한번에 100줄씩 적용 가능 합니다.');
		return;
	}

	if (confirm("검토 하셨습니까?\n실제 상품으로 적용 하시겠습니까?") == true) {
		frm.mode.value='regtemporder';
		frm.action="/admin/itemmaster/pop_itemlist_excel_upload_process.asp"
		frm.target="view";
		frm.submit();
	}
}

$(document).ready(function () {
    var checkAllBox = $("#chkall");
    checkAllBox.click(function () {
        var status = checkAllBox.prop('checked');
        toggleChecked(status);
    });
    var checkAllBox_fail = $("#chkall_fail");
    checkAllBox_fail.click(function () {
        var status_fail = checkAllBox_fail.prop('checked');
        toggleChecked_fail(status_fail);
    });
});

</script>

<form name="frmFile" method="post" action="<%= uploadUrl %>/linkweb/item/upload_itemlistreg_excel.asp"  enctype="MULTIPART/FORM-DATA" style="margin:0px;">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#999999">
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2">
		<b>온라인 상품 신규등록 엑셀 일괄 업로드</b>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td width="60">샘플</td>
	<td align="left">
		업로드 양식 : <a href="<%= uploadUrl %>/offshop/sample/item/sample_product_new_v4.xls" target="_blank">다운로드</a>
        <br>수정후 <font color="red"><b>Save As Excel 97 -2003 통합문서</b></font>로 저장후 업로드 해주세요.
        <!--<br>최대 <font color="red"><strong>1000개</strong></font> 상품씩 업로드 제한-->
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td><font color="red"><b>주의사항</b></font></td>
	<td align="left">
		* 맨 상단의 <font color="red"><b>첫줄</b></font>은 수정하지 마세요.
		<br><br>* <font color="red"><b>옵션이 있는 상품</b></font>의 경우 엑셀 샘플을 참조 하셔서, 동일 상품 옵션 순서대로 밑에줄에 이어서 입력해 주시면 됩니다.
        <br>&nbsp;&nbsp;
        같은상품의 두번째 옵션부터는 상품정보(브랜드ID,상품명,기본전시카테고리코드,소비자가,매입가,거래구분,배송구분,원산지,제조사,검색키워드,표시브랜드)는 공란으로 비워 두시고,
        <br>&nbsp;&nbsp;
        옵션명,범용바코드,업체관리코드,상품사이즈,상품무게,[매입용]상품명,[매입용]통화화폐,[매입용]매입가,[매입용]옵션명 만 입력해 주시면 됩니다.
		<br>&nbsp;&nbsp;
		공란으로 비워두신 정보는 윗줄의 상품정보가 그대로 복사 됩니다.
		<br><br>* 전시카테고리코드를 입력할경우, 자동으로 매칭되어 관리카테고리코드가 입력 됩니다.
		<br><br>* 상세 설명 : http://confluence.tenbyten.kr:8090/pages/editpage.action?pageId=59021677
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>파일명:</td>
	<td align="left"><input type="file" name="sFile" class="button"></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2">
	    <input type="button" class="button" value="업로드" onClick="XLSumbit();">
	    <input type="button" class="button" value="취소" onClick="self.close();">
	</td>
</tr>
</table>
</form>

<form name="frmdetail" method="post" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<% if oCManualMeachul.FResultCount > 0 then %>
	<Br>
	[업로드내역]
	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"><input type="checkbox" name="chkall" id="chkall"></td>
		<td width="90">임시상품코드<Br>[임시옵션코드]</td>
		<td width="100">브랜드ID</td>
		<td width="100">표시브랜드</td>
		<td width="110">전시카테고리코드<br>[관리카테고리코드]</td>
		<td>상품명<br>[옵션명]</td>
		<td width="60">소비자가<br>[매입가]</td>
		<td width="90">거래구분<br>[배송구분]</td>
		<td width="100">범용바코드<br>[업체관리코드]</td>
		<td width="60">상태</td>
		<td width="80">등록자</td>
	</tr>

	<% if oCManualMeachul.FResultCount > 0 then %>
		<% For i = 0 To oCManualMeachul.FResultCount - 1 %>
		<% if IsNull(oCManualMeachul.FItemList(i).Ffailtype) then %>
		<tr align="center" bgcolor="#FFFFFF">
		<% else %>
		<tr align="center" bgcolor="#CCCCCC">
		<% end if %>
			<td><input type="checkbox" name="chk_idx" value="<%= oCManualMeachul.FItemList(i).Fidx %>" ></td>
			<td><%= oCManualMeachul.FItemList(i).Ftempitemid %><br>[<%= oCManualMeachul.FItemList(i).Ftempitemoption %>]</td>
			<td><%= oCManualMeachul.FItemList(i).Fmakerid %></td>
			<td><%= oCManualMeachul.FItemList(i).ffrontmakerid %></td>
			<td align="left">
				<%= oCManualMeachul.FItemList(i).Fdispcatecode %>
				<br>[<% if oCManualMeachul.FItemList(i).fcate_large<>"" and not(isnull(oCManualMeachul.FItemList(i).fcate_large)) then %>
					<%= oCManualMeachul.FItemList(i).fcate_large %>
				<% end if %>
				<% if oCManualMeachul.FItemList(i).fcate_mid<>"" and not(isnull(oCManualMeachul.FItemList(i).fcate_mid)) then %>
					<%= oCManualMeachul.FItemList(i).fcate_mid %>
				<% end if %>
				<% if oCManualMeachul.FItemList(i).fcate_small<>"" and not(isnull(oCManualMeachul.FItemList(i).fcate_small)) then %>
					<%= oCManualMeachul.FItemList(i).fcate_small %>
				<% end if %>]
			</td>
			<td align="left">
				<%= oCManualMeachul.FItemList(i).Fitemname %>
				<% if oCManualMeachul.FItemList(i).Fitemoptionname<>"" and not(isnull(oCManualMeachul.FItemList(i).Fitemoptionname)) then %>
					<br>[<%= oCManualMeachul.FItemList(i).Fitemoptionname %>]
				<% end if %>
			</td>
			<td align="left"><%= FormatNumber(oCManualMeachul.FItemList(i).forgprice, 0) %><br>[<%= FormatNumber(oCManualMeachul.FItemList(i).Fbuycash, 0) %>]</td>
			<td><%= mwdivName(oCManualMeachul.FItemList(i).Fmwdiv) %><br>[<%= getdeliverytypename(oCManualMeachul.FItemList(i).Fdeliverytype) %>]</td>
			<td align="left">
				<%= oCManualMeachul.FItemList(i).Fbarcode %>
				<% if oCManualMeachul.FItemList(i).Fupchemanagecode<>"" and not(isnull(oCManualMeachul.FItemList(i).Fupchemanagecode)) then %>
					<br>[<%= oCManualMeachul.FItemList(i).Fupchemanagecode %>]
				<% end if %>
			</td>
			<td>
				<%= oCManualMeachul.FItemList(i).GetOrderTempStatusName %>
				<% if oCManualMeachul.FItemList(i).GetFailTypeName<>"" then %>
					<br><%= oCManualMeachul.FItemList(i).GetFailTypeName %>
				<% end if %>
			</td>
			<td>
				<%= oCManualMeachul.FItemList(i).Fregadminid %>
				<Br><%= left(oCManualMeachul.FItemList(i).Fregdate,10) %>
				<Br><%= mid(oCManualMeachul.FItemList(i).Fregdate,11,12) %>
			</td>
		</tr>
		<% next %>
		<tr bgcolor="#FFFFFF" align="center">
			<td colspan="16">
				<input type="button" class="button" value="선택상품 실제 상품으로 등록(100줄씩)" onclick="saveClick()">	
				<input type="button" class="button" value="삭제하기" onclick="delClick('delregitem');">
			</td>
		</tr>
	<% else %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td colspan="16" height="35">
				업로드내역 없음
			</td>
		</tr>
	<% end if %>
	</table>
<% end if %>

<% if oCFailManualMeachul.FResultCount > 0 then %>
	<br>
	[업로드실패]
	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"><input type="checkbox" name="chkall_fail" id="chkall_fail"></td>
		<td width="90">임시>상품코드<Br>[임시옵션코드]</td>
		<td width="100">브랜드ID</td>
		<td width="100">표시브랜드</td>
		<td width="110">전시카테고리코드<br>[관리카테고리코드]</td>
		<td>상품명<br>[옵션명]</td>
		<td width="60">소비자가<br>[매입가]</td>
		<td width="90">거래구분<br>[배송구분]</td>
		<td width="100">범용바코드<br>[업체관리코드]</td>
		<td width="60">상태</td>
		<td width="80">등록자</td>
	</tr>

	<% if oCFailManualMeachul.FResultCount > 0 then %>
		<% For i = 0 To oCFailManualMeachul.FResultCount - 1 %>
		<% if IsNull(oCFailManualMeachul.FItemList(i).Ffailtype) then %>
		<tr align="center" bgcolor="#FFFFFF">
		<% else %>
		<tr align="center" bgcolor="#CCCCCC">
		<% end if %>
			<td><input type="checkbox" name="chk_idx_fail" value="<%= oCFailManualMeachul.FItemList(i).Fidx %>" ></td>
			<td><%= oCFailManualMeachul.FItemList(i).Ftempitemid %><Br>[<%= oCFailManualMeachul.FItemList(i).Ftempitemoption %>]</td>
			<td><%= oCFailManualMeachul.FItemList(i).Fmakerid %></td>
			<td><%= oCFailManualMeachul.FItemList(i).ffrontmakerid %></td>
			<td align="left">
				<%= oCFailManualMeachul.FItemList(i).Fdispcatecode %>
				<br>[<% if oCFailManualMeachul.FItemList(i).fcate_large<>"" and not(isnull(oCFailManualMeachul.FItemList(i).fcate_large)) then %>
					<%= oCFailManualMeachul.FItemList(i).fcate_large %>
				<% end if %>
				<% if oCFailManualMeachul.FItemList(i).fcate_mid<>"" and not(isnull(oCFailManualMeachul.FItemList(i).fcate_mid)) then %>
					<%= oCFailManualMeachul.FItemList(i).fcate_mid %>
				<% end if %>
				<% if oCFailManualMeachul.FItemList(i).fcate_small<>"" and not(isnull(oCFailManualMeachul.FItemList(i).fcate_small)) then %>
					<%= oCFailManualMeachul.FItemList(i).fcate_small %>
				<% end if %>]
			</td>
			<td align="left">
				<%= oCFailManualMeachul.FItemList(i).Fitemname %>
				<% if oCFailManualMeachul.FItemList(i).Fitemoptionname<>"" and not(isnull(oCFailManualMeachul.FItemList(i).Fitemoptionname)) then %>
					<br>[<%= oCFailManualMeachul.FItemList(i).Fitemoptionname %>]
				<% end if %>
			</td>
			<td align="left"><%= FormatNumber(oCFailManualMeachul.FItemList(i).forgprice, 0) %><br>[<%= FormatNumber(oCFailManualMeachul.FItemList(i).Fbuycash, 0) %>]</td>
			<td><%= mwdivName(oCFailManualMeachul.FItemList(i).Fmwdiv) %><br>[<%= getdeliverytypename(oCFailManualMeachul.FItemList(i).Fdeliverytype) %>]</td>
			<td align="left">
				<%= oCFailManualMeachul.FItemList(i).Fbarcode %>
				<% if oCFailManualMeachul.FItemList(i).Fupchemanagecode<>"" and not(isnull(oCFailManualMeachul.FItemList(i).Fupchemanagecode)) then %>
					<br>[<%= oCFailManualMeachul.FItemList(i).Fupchemanagecode %>]
				<% end if %>
			</td>
			<td>
				<%= oCFailManualMeachul.FItemList(i).GetOrderTempStatusName %>
				<% if oCFailManualMeachul.FItemList(i).GetFailTypeName<>"" then %>
					<br><%= oCFailManualMeachul.FItemList(i).GetFailTypeName %>
				<% end if %>
			</td>
			<td>
				<%= oCFailManualMeachul.FItemList(i).Fregadminid %>
				<Br><%= left(oCFailManualMeachul.FItemList(i).Fregdate,10) %>
				<Br><%= mid(oCFailManualMeachul.FItemList(i).Fregdate,11,12) %>
			</td>
		</tr>
		<% next %>
		<tr bgcolor="#FFFFFF" align="center">
			<td colspan="16">
				<input type="button" class="button" value="삭제하기" onclick="delClick('delregitem_fail');">
			</td>
		</tr>
	<% else %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td colspan="16" height="35">
				업로드내역 없음
			</td>
		</tr>
	<% end if %>
	</table>
<% end if %>

</form>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width=1280 height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width=1280 height=0 frameborder="0" scrolling="no"></iframe>
<% end if %>

<%
function getdeliverytypename(vdeliverytype)
    dim deliverytypename

    if vdeliverytype="1" then
        deliverytypename="텐바이텐배송"
    elseif vdeliverytype="2" then
        deliverytypename="업체배송"
    else
        deliverytypename=""
    end if

    getdeliverytypename=deliverytypename
end function
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->