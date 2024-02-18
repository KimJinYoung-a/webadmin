<%@ language=vbscript %>
<% option explicit %>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/common/commonCls.asp"-->
<%
Dim itemid, i, mallgubun, oCommon, arrRows, arrRows2
Dim searchArea
searchArea = request("searchArea")
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function setCode(id, area){
	var text;
	text = "선택 코드 : " + id + "(" + area + ")";
	$("#sCode").text(text);
	$("#sid").val(id);
	$("#savebtn").show();
}

function delCode(id, area){
	document.frm.mode.value = "D";
	document.frm.target = "xLink";
	document.frm.action = "/admin/etc/common/procSourcearea.asp?sid="+id+"&sname="+area;
	document.frm.submit();
}


function keyWordsProcess() {
	var chkSel=0;
	var keywords = "";
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) {
					chkSel++;
					keywords = keywords + frmSvArr.keywords[i].value + "*(^!";
				}
			}
		} else {
			if(frmSvArr.cksel.checked){
				 chkSel++;
				 keywords = frmSvArr.keywords.value;
			}
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm(chkSel + '개의 상품 키워드 변경을 적용하시겠습니까?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.mode.value = "REG";
		document.frmSvArr.arrkeywords.value = keywords;
		document.frmSvArr.action = "/admin/etc/common/procKeywords.asp"
		document.frmSvArr.submit();
	}
}
function frmSave(f){
	if($("#sname").val() == ""){
		alert("원산지명을 입력하세요");
		return;
	}

	if (confirm('원산지명을 적용하시겠습니까?')){
		f.target = "xLink";
		f.action = "/admin/etc/common/procSourcearea.asp"
		f.submit();
	}
}

function goPage(pg){
    frm2.page.value = pg;
    frm2.submit();
}

</script>

<div style="width:100%;float:left;">
<form name="frm" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="I">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<th colsapn="2">원산지등록</th>
</tr>
<tr align="center"  bgcolor="#FFFFFF">
	<td width="50%">원산지명
		<input type="text" class="text" size="50" id="sname" name="sname" value=""> <br />
		<input type="hidden" class="text" size="50" id="sid" name="sid">
		<span id="sCode"></span>
		<input type="button" id="savebtn" class="button" style="display:none;" value="저장" onclick="frmSave(this.form);"; >
	</td>
</tr>
</table>
</form>
</div>
<br />
<div style="width:49%;float:left;">
	<span align="center"><h4>텐바이텐 원산지 목록</h4></span>
	<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
	<input type="hidden" name="arrkeywords" value="" />
	<input type="hidden" name="mode" value="" />
	<input type="hidden" name="mallgubun" value="<%= mallgubun %>" />
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<th width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
		<th width="120">원산지코드</td>
		<th>원산지명</td>
		<th width="20">비고</td>
	</tr>
<%
SET oCommon = new CCommon
	oCommon.FCurrPage		= 1
	oCommon.FPageSize		= 1000
	oCommon.getOutmallSSGSourceAreaMappList
	If oCommon.FResultCount > 0 Then
		For k=0 to oCommon.FResultCount - 1
%>
	<tr align="center"  bgcolor="#FFFFFF">
		<td><input type="checkbox" name="cksel" id="cksel<%= k %>" onClick="AnCheckClick(this);"  value="<%= oCommon.FItemList(k).FId %>"></td>
		<td><%= oCommon.FItemList(k).FId %></td>
		<td align="LEFT"><%= oCommon.FItemList(k).FSourceArea %></td>
		<td align="LEFT">
			<input type="button" id="del" class="button" value="삭제" onclick="delCode('<%= oCommon.FItemList(k).FId %>', '<%= oCommon.FItemList(k).FSourceArea %>')">
		</td>
	</tr>
<%
	Next
Else
%>
	<tr align="center" height="50" bgcolor="#FFFFFF">
		<td colspan="3">검색 된 데이터가 없습니다.</td>
	</tr>
<%
End If
%>
	</table>
	</form>
</div>
<%
SET oCommon = nothing

Dim sItemid
sItemid = request("sItemid")

If sItemid<>"" then
	Dim iA2, arrTemp2, arrItemid2
	sItemid = replace(sItemid,",",chr(10))
	sItemid = replace(sItemid,chr(13),"")
	arrTemp2 = Split(sItemid,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid2 = arrItemid2 & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	sItemid = left(arrItemid2,len(arrItemid2)-1)
End If

Dim page, k
page = request("page")
If page = "" Then page = 1
SET oCommon = new CCommon
	oCommon.FCurrPage			= page
	oCommon.FPageSize			= 1000
	oCommon.FRectSourceArea		= searchArea
	oCommon.getOutmallOrgSSGSourceAreaList
%>
<div style="width:49%;float:right;">
	<span align="center"><h4><%= mallgubun %> 원산지 목록</h4></span>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm2" method="get" action="">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			원산지명 : <input type="text" name="searchArea" value="<%= searchArea %>">
		</td>
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm2.submit();">
		</td>
	</table>
	</form>
	<br />

	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<th width="120">원산지코드</td>
		<th>원산지명</td>
	</tr>
<%
If oCommon.FResultCount > 0 Then
	For k=0 to oCommon.FResultCount - 1
%>
	<tr align="center" bgcolor="#FFFFFF" >
		<td style="cursor:pointer;" onclick="setCode('<%= oCommon.FItemList(k).FId %>', '<%= oCommon.FItemList(k).FSourceArea %>')"><%= oCommon.FItemList(k).FId %></td>
		<td align="left"><%= oCommon.FItemList(k).FSourceArea %></td>
	</tr>
<%
	Next
%>
	<tr height="20">
		<td colspan="19" align="center" bgcolor="#FFFFFF">
		<% If oCommon.HasPreScroll Then %>
			<a href="javascript:goPage('<%= oCommon.StartScrollPage-1 %>');">[pre]</a>
		<% Else %>
			[pre]
		<%
			End If
			For i = 0 + oCommon.StartScrollPage to oCommon.FScrollCount + oCommon.StartScrollPage - 1
				If i > oCommon.FTotalpage Then Exit For
				If CStr(page)=CStr(i) Then
		%>
				<font color="red">[<%= i %>]</font>
		<%		Else %>
				<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<%		End If
			Next
			If oCommon.HasNextScroll Then
		%>
				<a href="javascript:goPage('<%= i %>');">[next]</a>
		<%	Else %>
				[next]
		<%	End If %>
		</td>
	</tr>
<%
Else
%>
	<tr align="center" height="50" bgcolor="#FFFFFF">
		<td colspan="3">등록 된 데이터가 없습니다.</td>
	</tr>
<%
End If
%>
	</table>
</div>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="<%= CHKIIF(request("auto") <> "Y",300,100) %>"></iframe>
<% SET oCommon = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->