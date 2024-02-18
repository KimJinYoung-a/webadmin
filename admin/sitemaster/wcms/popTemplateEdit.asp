<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/mainWCMSCls.asp" -->
<%
'###############################################
' PageName : popTemplateEdit.asp
' Discription : 템플릿 등록/수정
' History : 2013.04.01 허진원 : 신규 생성
'###############################################

'// 변수 선언
Dim siteDiv, pageDiv, i, page
Dim oTemplate
Dim tplIdx, tplType, tplName, isTimeUse, isIconUse, isSubNumUse, isTopImgUse, isTopLinkUse
Dim isImageUse, isTextUse, isLinkUse, isItemUse, isVideoUse, isBGColorUse, isExtDataUse, isImgDescUse, tplinfoDesc, tplSortNo

'// 파라메터 접수
siteDiv = request("site")
pageDiv = request("pDiv")
tplIdx = request("tplIdx")
page = request("page")

if siteDiv="" then siteDiv="P"		'기본값 PC웹(P:PC웹, M:모바일)
if pageDiv="" then pageDiv="10"		'기본값 사이트메인(10:사이트메인, 20:이벤트메인...)
if page="" then page="1"

'// 템플릿 내용
	set oTemplate = new CCMSContent
	oTemplate.FRectTplIdx = tplIdx
    if tplIdx<>"" then
    	oTemplate.GetOneTemplate
		if oTemplate.FResultCount>0 then
			tplType			= oTemplate.FOneItem.FtplType
			tplName			= oTemplate.FOneItem.FtplName
			siteDiv			= oTemplate.FOneItem.FsiteDiv
			pageDiv			= oTemplate.FOneItem.FpageDiv
			isTimeUse		= oTemplate.FOneItem.FisTimeUse
			isIconUse		= oTemplate.FOneItem.FisIconUse
			isSubNumUse		= oTemplate.FOneItem.FisSubNumUse
			isTopImgUse		= oTemplate.FOneItem.FisTopImgUse
			isTopLinkUse	= oTemplate.FOneItem.FisTopLinkUse
			isImageUse		= oTemplate.FOneItem.FisImageUse
			isTextUse		= oTemplate.FOneItem.FisTextUse
			isLinkUse		= oTemplate.FOneItem.FisLinkUse
			isItemUse		= oTemplate.FOneItem.FisItemUse
			isVideoUse		= oTemplate.FOneItem.FisVideoUse
			isBGColorUse	= oTemplate.FOneItem.FisBGColorUse
			isExtDataUse	= oTemplate.FOneItem.FisExtDataUse
			isImgDescUse	= oTemplate.FOneItem.FisImgDescUse
			tplinfoDesc		= oTemplate.FOneItem.FtplinfoDesc
			tplSortNo		= oTemplate.FOneItem.FtplSortNo
		end if
    else
    	tplSortNo = "0"
    end if
    set oTemplate = Nothing

'// 템플릿 목록
	set oTemplate = new CCMSContent
	oTemplate.FRectSiteDiv = siteDiv
	oTemplate.FRectPageDiv = pageDiv
	oTemplate.FCurrPage = page
    oTemplate.GetTemplateList
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
//템플릿 유형(preset,group)
function chgTplType(v) {
	switch(v) {
		case "A" :
			$("#tplTpDesc").html("썸네일+키워드 유형");
			$("select:not(select[name='tplType'])").val("N");
			$("select[name='isImageUse']").val("Y");
			$("select[name='isTextUse']").val("Y");
			$("select[name='isLinkUse']").val("Y");
			$("select[name='isImgDescUse']").val("Y");
			break;
		case "B" :
			$("#tplTpDesc").html("텍스트 링크 유형");
			$("select:not(select[name='tplType'])").val("N");
			$("select[name='isTimeUse']").val("Y");
			$("select[name='isTextUse']").val("Y");
			$("select[name='isLinkUse']").val("Y");
			$("select[name='isBGColorUse']").val("Y");
			break;
		case "C" :
			$("#tplTpDesc").html("이미지 링크 유형");
			$("select:not(select[name='tplType'])").val("N");
			$("select[name='isTimeUse']").val("Y");
			$("select[name='isImageUse']").val("Y");
			$("select[name='isLinkUse']").val("Y");
			$("select[name='isBGColorUse']").val("Y");
			$("select[name='isImgDescUse']").val("Y");
			break;
		case "D" :
			$("#tplTpDesc").html("카피/아이콘 상품링크 유형");
			$("select:not(select[name='tplType'])").val("N");
			$("select[name='isTimeUse']").val("Y");
			$("select[name='isIconUse']").val("Y");
			$("select[name='isSubNumUse']").val("Y");
			$("select[name='isTextUse']").val("Y");
			$("select[name='isItemUse']").val("Y");
			break;
		case "E" :
			$("#tplTpDesc").html("베스트 상품 유형");
			$("select:not(select[name='tplType'])").val("N");
			$("select[name='isExtDataUse']").val("Y");
			break;
		case "F" :
			$("#tplTpDesc").html("이미지, 상품 링크 유형");
			$("select:not(select[name='tplType'])").val("N");
			$("select[name='isTimeUse']").val("Y");
			$("select[name='isSubNumUse']").val("Y");
			$("select[name='isTopImgUse']").val("Y");
			$("select[name='isTopLinkUse']").val("Y");
			$("select[name='isItemUse']").val("Y");
			break;
		case "G" :
			$("#tplTpDesc").html("동영상 링크 유형");
			$("select:not(select[name='tplType'])").val("N");
			$("select[name='isTimeUse']").val("Y");
			$("select[name='isTextUse']").val("Y");
			$("select[name='isLinkUse']").val("Y");
			$("select[name='isVideoUse']").val("Y");
			break;
		default :
			$("#tplTpDesc").html("");
			$("select:not(select[name='tplType'])").val("");
	}
}

// 폼검사
function SaveTemplate(frm) {
	var selChk=true;
	$("select").each(function(){
		if($(this).val()=="") {
			alert($(this).attr("title")+"을(를) 선택해주세요");
			$(this).focus();
			selChk=false;
			return false;
		}
	});
	if(!selChk) return;

	if($("input[name='tplName']").val()=="") {
		alert("템플릿명을 입력해주세요.");
		$("input[name='tplName']").focus();
		selChk=false;
	}
	if(selChk) {
		frm.submit();
	} else {
		return;
	}
}
</script>
<center>
<form name="frmTemplate" method="post" action="doTemplate.asp" style="margin:0px;">
<input type="hidden" name="page" value="" />
<table width="690" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<tr bgcolor="#FFFFFF">
    <td height="25" colspan="4" bgcolor="#F8F8F8"><b>템플릿 등록/수정</b></td>
</tr>
<% if tplIdx<>"" then %>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">템플릿번호</td>
    <td width="610" colspan="3">
        <%=tplIdx %>
        <input type="hidden" name="tplIdx" value="<%=tplIdx %>" />
    </td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">템플릿유형</td>
    <td width="610" colspan="3">
        <select name="tplType" class="select" onchange="chgTplType(this.value)" title="템플릿유형">
        	<option value="">::선택::</option>
        	<option value="A" <%=chkIIF(tplType="A","selected","")%>>A Type</option>
        	<option value="B" <%=chkIIF(tplType="B","selected","")%>>B Type</option>
        	<option value="C" <%=chkIIF(tplType="C","selected","")%>>C Type</option>
        	<option value="D" <%=chkIIF(tplType="D","selected","")%>>D Type</option>
        	<option value="E" <%=chkIIF(tplType="E","selected","")%>>E Type</option>
        	<option value="F" <%=chkIIF(tplType="F","selected","")%>>F Type</option>
        	<option value="G" <%=chkIIF(tplType="G","selected","")%>>G Type</option>
        </select>
        &nbsp;<span id="tplTpDesc"></span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">사이트구분</td>
    <td width="230">
    	<%=chkIIF(siteDiv="P","PC웹","모바일")%>
    	<input type="hidden" name="site" value="<%=siteDiv%>" />
    </td>
    <td width="100" bgcolor="#DDDDFF">사용처</td>
    <td width="230">
        <select name="pageDiv" class="select" title="사용처">
        	<option value="">::선택::</option>
        	<option value="10" <%=chkIIF(pageDiv="10","selected","")%>>사이트 메인</option>
        	<option value="20" <%=chkIIF(pageDiv="20","selected","")%>>이벤트 메인</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">템플릿명</td>
    <td width="610" colspan="3">
        <input type="text" name="tplName" value="<%= tplName %>" maxlength="64" size="64" title="템플릿명">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">시간표시 여부</td>
    <td>
        <select name="isTimeUse" class="select" title="시간표시 여부">
        	<option value="">::선택::</option>
        	<option value="Y" <%=chkIIF(isTimeUse="Y","selected","")%>>사용</option>
        	<option value="N" <%=chkIIF(isTimeUse="N","selected","")%>>사용안함</option>
        </select>
    </td>
    <td bgcolor="#DDDDFF">아이콘 사용</td>
    <td>
        <select name="isIconUse" class="select" title="아이콘 사용여부">
        	<option value="">::선택::</option>
        	<option value="Y" <%=chkIIF(isIconUse="Y","selected","")%>>사용</option>
        	<option value="N" <%=chkIIF(isIconUse="N","selected","")%>>사용안함</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">소재개수 제한</td>
    <td>
        <select name="isSubNumUse" class="select" title="소재개수 제한 여부">
        	<option value="">::선택::</option>
        	<option value="Y" <%=chkIIF(isSubNumUse="Y","selected","")%>>사용</option>
        	<option value="N" <%=chkIIF(isSubNumUse="N","selected","")%>>사용안함</option>
        </select>
    </td>
    <td bgcolor="#DDDDFF">외부자료 사용</td>
    <td>
        <select name="isExtDataUse" class="select" title=외부자료 사용 여부">
        	<option value="">::선택::</option>
        	<option value="Y" <%=chkIIF(isExtDataUse="Y","selected","")%>>사용</option>
        	<option value="N" <%=chkIIF(isExtDataUse="N","selected","")%>>사용안함</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">탑이미지 여부</td>
    <td>
        <select name="isTopImgUse" class="select" title="탑이미지 사용여부">
        	<option value="">::선택::</option>
        	<option value="Y" <%=chkIIF(isTopImgUse="Y","selected","")%>>사용</option>
        	<option value="N" <%=chkIIF(isTopImgUse="N","selected","")%>>사용안함</option>
        </select>
    </td>
    <td bgcolor="#DDDDFF">탑링크 여부</td>
    <td>
        <select name="isTopLinkUse" class="select" title="탑이미지 링크 사용여부">
        	<option value="">::선택::</option>
        	<option value="Y" <%=chkIIF(isTopLinkUse="Y","selected","")%>>사용</option>
        	<option value="N" <%=chkIIF(isTopLinkUse="N","selected","")%>>사용안함</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">이미지 사용</td>
    <td>
        <select name="isImageUse" class="select" title="이미지 사용여부">
        	<option value="">::선택::</option>
        	<option value="Y" <%=chkIIF(isImageUse="Y","selected","")%>>사용</option>
        	<option value="N" <%=chkIIF(isImageUse="N","selected","")%>>사용안함</option>
        </select>
    </td>
    <td bgcolor="#DDDDFF">이미지설명 사용</td>
    <td>
        <select name="isImgDescUse" class="select" title="이미지설명 사용여부(이미지가 있는 경우)">
        	<option value="">::선택::</option>
        	<option value="Y" <%=chkIIF(isImgDescUse="Y","selected","")%>>사용</option>
        	<option value="N" <%=chkIIF(isImgDescUse="N","selected","")%>>사용안함</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">텍스트 사용</td>
    <td>
        <select name="isTextUse" class="select" title="텍스트 사용여부">
        	<option value="">::선택::</option>
        	<option value="Y" <%=chkIIF(isTextUse="Y","selected","")%>>사용</option>
        	<option value="N" <%=chkIIF(isTextUse="N","selected","")%>>사용안함</option>
        </select>
    </td>
    <td bgcolor="#DDDDFF">링크 사용</td>
    <td>
        <select name="isLinkUse" class="select" title="링크 사용여부">
        	<option value="">::선택::</option>
        	<option value="Y" <%=chkIIF(isLinkUse="Y","selected","")%>>사용</option>
        	<option value="N" <%=chkIIF(isLinkUse="N","selected","")%>>사용안함</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">상품코드 사용</td>
    <td>
        <select name="isItemUse" class="select" title="상품코드 사용여부">
        	<option value="">::선택::</option>
        	<option value="Y" <%=chkIIF(isItemUse="Y","selected","")%>>사용</option>
        	<option value="N" <%=chkIIF(isItemUse="N","selected","")%>>사용안함</option>
        </select>
    </td>
    <td bgcolor="#DDDDFF">동영상 사용</td>
    <td>
        <select name="isVideoUse" class="select" title="동영상URL 사용여부">
        	<option value="">::선택::</option>
        	<option value="Y" <%=chkIIF(isVideoUse="Y","selected","")%>>사용</option>
        	<option value="N" <%=chkIIF(isVideoUse="N","selected","")%>>사용안함</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">배경색 사용</td>
    <td colspan="3">
        <select name="isBGColorUse" class="select" title="배경색 사용여부">
        	<option value="">::선택::</option>
        	<option value="Y" <%=chkIIF(isBGColorUse="Y","selected","")%>>사용</option>
        	<option value="N" <%=chkIIF(isBGColorUse="N","selected","")%>>사용안함</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">템플릿 사용안내</td>
    <td colspan="3">
        <textarea name="tplinfoDesc" class="textarea" style="width:100%; height:60px;" title="템플릿 사용안내"><%=tplinfoDesc%></textarea>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">정렬순서</td>
    <td colspan="3">
        <input type="text" name="tplSortNo" class="text" size="4" value="<%=tplSortNo%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="4" align="center"><input type="button" value=" 저 장 " onClick="SaveTemplate(this.form);"></td>
</tr>
</table>
</form>
<br>
<!-- // 등록된 템플릿 목록 --------->
<table width="690" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<tr bgcolor="#FFFFFF">
    <td colspan="5" align="right"><a href="?site=<%=siteDiv%>"><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
    <td width="60">code</td>
    <td width="80">템플릿유형</td>
    <td width="100">템플릿유형</td>
    <td>템플릿명</td>
    <td width="80">정렬</td>
</tr>
<% for i=0 to oTemplate.FResultCount-1 %>
<% if (CStr(oTemplate.FItemList(i).FtplIdx)=tplIdx) then %>
<tr bgcolor="#9999CC">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
    <td align="center"><a href="?site=<%=siteDiv%>&tplIdx=<%= oTemplate.FItemList(i).FtplIdx %>&page=<%= page %>"><%= oTemplate.FItemList(i).FtplIdx %></a></td>
    <td align="center"><%= oTemplate.FItemList(i).FtplType %> Type</td>
    <td align="center"><%= oTemplate.FItemList(i).getPageDiv %></td>
    <td ><a href="?site=<%=siteDiv%>&tplIdx=<%= oTemplate.FItemList(i).FtplIdx %>&page=<%= page %>"><%= oTemplate.FItemList(i).FtplName %></a></td>
    <td align="center"><%= oTemplate.FItemList(i).FtplSortNo %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="5" align="center">
    <% if oTemplate.HasPreScroll then %>
		<a href="?site=<%=siteDiv%>&page=<%= oTemplate.StartScrollPage-1 %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oTemplate.StartScrollPage to oTemplate.FScrollCount + oTemplate.StartScrollPage - 1 %>
		<% if i>oTemplate.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?site=<%=siteDiv%>&page=<%= i %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oTemplate.HasNextScroll then %>
		<a href="?site=<%=siteDiv%>&page=<%= i %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
</center>
<%	set oTemplate = Nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->