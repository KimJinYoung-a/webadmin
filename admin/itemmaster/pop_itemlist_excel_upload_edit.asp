<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  상품 엑셀 업로드 일괄 수정
' History : 2019.04.18 한용민 생성
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
	oCManualMeachul.GetsuccessitemList

dim oCFailManualMeachul
set oCFailManualMeachul = new Citemedit_templist
	oCFailManualMeachul.FPageSize = 1000
	oCFailManualMeachul.FCurrPage = 1
	oCFailManualMeachul.FRectRegAdminID = session("ssBctId")
	oCFailManualMeachul.GetFailList
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
	//arrFileExt[arrFileExt.length]  = "csv";     //xls
	arrFileExt[arrFileExt.length]  = "xls";     //csv
	
	if (frm.sFile.value==''){
		alert('파일을 입력해 주세요');
		frm.sFile.focus();
		return;
	}
	
	//파일유효성 체크
	if (!fnChkFile(frm.sFile.value, arrFileExt)){
		//alert("파일은 csv파일만 업로드 가능합니다.");   // xls
		alert("파일은 xls파일만 업로드 가능합니다.");   // csv
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

	if (mode=='delitem_fail'){
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

	if (confirm("검토 하셨습니까?\n실제 적용 하시겠습니까?") == true) {
		frm.mode.value='edittemporder';
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

<form name="frmFile" method="post" action="<%= uploadUrl %>/linkweb/item/upload_itemlist_excel.asp"  enctype="MULTIPART/FORM-DATA" style="margin:0px;">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#999999">
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2">
		<b>온라인 상품 엑셀 일괄 업로드 수정</b>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td width="60">샘플</td>
	<td align="left">
		업로드 양식은 상품목록에 있는 "상품다운로드(엑셀)" 입니다. 수정후 <font color="red"><b>Save As Excel 97 -2003 통합문서</b></font>로 저장후 업로드 해주세요.
		<% '<br>최대 <font color="red"><strong>1000개</strong></font> 상품씩 업로드 제한 %>
		<!--<a href="<%'= uploadUrl %>/offshop/sample/item/item_list_sample_v1.csv" target="_blank">다운로드</a>-->
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td><font color="red"><b>주의사항</b></font></td>
	<td align="left">
		※ 엑셀에서 저장형식이 <font color="red"><b>Save As Excel 97 -2003 통합문서</b></font> 형태만 인식 합니다.
		<br><br>* 맨 상단의 상품코드, 브랜드 등이 있는 <font color="red"><b>1줄은 그대로 두세요.</b></font>
		<br><br>* 상품코드를 기준으로 업데이트 되기 때문에 <font color="red"><strong>상품코드는 절대 공란</strong></font>이거나, 틀리면 안됩니다.
		<br><br><font color="red"><b>* 상품명,소비자가(매입가는 브랜드 기본마진에 따라 자동계산),ISBN13,표시브랜드</b></font> 필드를 입력하시면 되며, 그대로 저장 됩니다.
		<br><br><font color="red"><strong>* 할인중</strong></font>인 상품은 등록되지 않습니다.
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
		<td width="60">번호</td>
		<td width="60">상품코드</td>
		<td>현재상품명</td>
		<td>수정상품명</td>
		<td width="60">현재<br>소비자가</td>
		<td width="60">수정<br>소비자가</td>
		<td width="100">현재<br>isbn13</td>
		<td width="100">수정<br>isbn13</td>
		<td width="100">현재<br>표시브랜드</td>
		<td width="100">수정<br>표시브랜드</td>
		<td width="80">등록자</td>
		<td width="70">상태</td>
		<td width="70">비고</td>
	</tr>

	<% if oCManualMeachul.FResultCount > 0 then %>
		<% For i = 0 To oCManualMeachul.FResultCount - 1 %>
		<% if IsNull(oCManualMeachul.FItemList(i).Ffailtype) then %>
		<tr align="center" bgcolor="#FFFFFF">
		<% else %>
		<tr align="center" bgcolor="#CCCCCC">
		<% end if %>
			<td><input type="checkbox" name="chk_idx" value="<%= oCManualMeachul.FItemList(i).Fidx %>" ></td>
			<td><%= oCManualMeachul.FItemList(i).Fidx %></td>
			<td><%= oCManualMeachul.FItemList(i).Fitemid %></td>
			<td align="left"><%= oCManualMeachul.FItemList(i).Fitemname_10x10 %></td>
			<td align="left"><%= oCManualMeachul.FItemList(i).Fitemname %></td>
			<td align="right"><%= FormatNumber(oCManualMeachul.FItemList(i).forgprice_10x10, 0) %></td>
			<td align="right">
				<%= FormatNumber(oCManualMeachul.FItemList(i).forgprice, 0) %>

				<% if oCManualMeachul.FItemList(i).forgprice <= oCManualMeachul.FItemList(i).forgprice_10x10*0.1 then %>
					<Br><font color="red"><strong>90%이상할인</strong></font>
				<% elseif oCManualMeachul.FItemList(i).forgprice <= oCManualMeachul.FItemList(i).forgprice_10x10*0.2 then %>
					<Br><font color="red"><strong>80%이상할인</strong></font>
				<% elseif oCManualMeachul.FItemList(i).forgprice <= oCManualMeachul.FItemList(i).forgprice_10x10*0.3 then %>
					<Br><font color="red"><strong>70%이상할인</strong></font>
				<% elseif oCManualMeachul.FItemList(i).forgprice <= oCManualMeachul.FItemList(i).forgprice_10x10*0.4 then %>
					<Br><font color="red"><strong>60%이상할인</strong></font>
				<% elseif oCManualMeachul.FItemList(i).forgprice <= oCManualMeachul.FItemList(i).forgprice_10x10*0.5 then %>
					<Br><font color="red"><strong>50%이상할인</strong></font>
				<% end if %>
			</td>
			<td align="left"><%= oCManualMeachul.FItemList(i).fisbn13_10x10 %></td>
			<td align="left"><%= oCManualMeachul.FItemList(i).fisbn13 %></td>
			<td><%= oCManualMeachul.FItemList(i).ffrontmakerid_10x10 %></td>
			<td><%= oCManualMeachul.FItemList(i).ffrontmakerid %></td>
			<td><%= oCManualMeachul.FItemList(i).Fregadminid %></td>
			<td><%= oCManualMeachul.FItemList(i).GetOrderTempStatusName %></td>
			<td>
				<%= oCManualMeachul.FItemList(i).GetFailTypeName %>
			</td>
		</tr>
		<% next %>
		<tr bgcolor="#FFFFFF" align="center">
			<td colspan="17">
				<input type="button" class="button" value="선택상품 실제 상품에 적용" onclick="saveClick()">	
				<input type="button" class="button" value="삭제하기" onclick="delClick('delitem');">
			</td>
		</tr>
	<% else %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td colspan="17" height="35">
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
		<td width="60">번호</td>
		<td width="20"><input type="checkbox" name="chkall_fail" id="chkall_fail"></td>
		<td width="60">상품코드</td>
		<td>현재상품명</td>
		<td>수정상품명</td>
		<td width="60">현재<br>소비자가</td>
		<td width="80">수정<br>소비자가</td>
		<td width="100">현재<br>isbn13</td>
		<td width="100">수정<br>isbn13</td>
		<td width="100">현재<br>표시브랜드</td>
		<td width="100">수정<br>표시브랜드</td>
		<td width="80">등록자</td>
		<td width="70">상태</td>
		<td width="70">비고</td>
	</tr>

	<% if oCFailManualMeachul.FResultCount > 0 then %>
		<% For i = 0 To oCFailManualMeachul.FResultCount - 1 %>
		<% if IsNull(oCFailManualMeachul.FItemList(i).Ffailtype) then %>
		<tr align="center" bgcolor="#FFFFFF">
		<% else %>
		<tr align="center" bgcolor="#CCCCCC">
		<% end if %>
			<td><%= oCFailManualMeachul.FItemList(i).Fidx %></td>
			<td><input type="checkbox" name="chk_idx_fail" value="<%= oCFailManualMeachul.FItemList(i).Fidx %>" ></td>
			<td><%= oCFailManualMeachul.FItemList(i).Fitemid %></td>
			<td align="left"><%= oCFailManualMeachul.FItemList(i).Fitemname_10x10 %></td>
			<td align="left"><%= oCFailManualMeachul.FItemList(i).Fitemname %></td>
			<td align="right"><%= FormatNumber(oCFailManualMeachul.FItemList(i).forgprice_10x10, 0) %></td>
			<td align="right">
				<%= FormatNumber(oCFailManualMeachul.FItemList(i).forgprice, 0) %>

				<% if oCFailManualMeachul.FItemList(i).forgprice <= oCFailManualMeachul.FItemList(i).forgprice_10x10*0.1 then %>
					<Br><font color="red"><strong>90%이상할인</strong></font>
				<% elseif oCFailManualMeachul.FItemList(i).forgprice <= oCFailManualMeachul.FItemList(i).forgprice_10x10*0.2 then %>
					<Br><font color="red"><strong>80%이상할인</strong></font>
				<% elseif oCFailManualMeachul.FItemList(i).forgprice <= oCFailManualMeachul.FItemList(i).forgprice_10x10*0.3 then %>
					<Br><font color="red"><strong>70%이상할인</strong></font>
				<% elseif oCFailManualMeachul.FItemList(i).forgprice <= oCFailManualMeachul.FItemList(i).forgprice_10x10*0.4 then %>
					<Br><font color="red"><strong>60%이상할인</strong></font>
				<% elseif oCFailManualMeachul.FItemList(i).forgprice <= oCFailManualMeachul.FItemList(i).forgprice_10x10*0.5 then %>
					<Br><font color="red"><strong>50%이상할인</strong></font>
				<% end if %>
			</td>
			<td align="left"><%= oCFailManualMeachul.FItemList(i).fisbn13_10x10 %></td>
			<td align="left"><%= oCFailManualMeachul.FItemList(i).fisbn13 %></td>
			<td><%= oCFailManualMeachul.FItemList(i).ffrontmakerid_10x10 %></td>
			<td><%= oCFailManualMeachul.FItemList(i).ffrontmakerid %></td>
			<td><%= oCFailManualMeachul.FItemList(i).Fregadminid %></td>
			<td><%= oCFailManualMeachul.FItemList(i).GetOrderTempStatusName %></td>
			<td>
				<%= oCFailManualMeachul.FItemList(i).GetFailTypeName %>
			</td>
		</tr>
		<% next %>
		<tr bgcolor="#FFFFFF" align="center">
			<td colspan="17">
				<input type="button" class="button" value="삭제하기" onclick="delClick('delitem_fail');">
			</td>
		</tr>
	<% else %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td colspan="17" height="35">
				업로드내역 없음
			</td>
		</tr>
	<% end if %>
	</table>
<% end if %>

</form>

<% IF application("Svr_Info")="Dev" or C_ADMIN_AUTH THEN %>
	<iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<% end if %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->