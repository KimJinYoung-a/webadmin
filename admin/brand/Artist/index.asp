<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.19 김진영 생성
'			2013.08.29 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/artistGalleryCls.asp" -->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
Dim page, isusing, gal_div, designerid, lp, oGallery, oGalleryitem
dim catecode, standardCateCode, mduserid, brandgubun
	catecode	= request("catecode")
	standardCateCode	= request("standardCateCode")
	mduserid	= request("mduserid")
	brandgubun	= request("brandgubun")
	page = request("page")
	isusing = request("isusing")
	gal_div = request("gal_div")
	designerid = request("designerid")

If page = "" Then page = 1
If isusing = "" Then isusing="Y"

'// 목록 접수
Set oGallery = New CGallery
	oGallery.FCurrPage = page
	oGallery.FPageSize=20
	oGallery.FRectGal_div = gal_div
	oGallery.FRectDesignerId = designerid
	oGallery.FRectIsusing = isusing
	oGallery.Frectbrandgubun		= brandgubun
	oGallery.Frectcatecode = catecode
	oGallery.FrectstandardCateCode = standardCateCode
	oGallery.Frectmduserid = mduserid	
	oGallery.GetGalleryList

'//메인페이지 배너 6개 리스트
Set oGalleryitem = New CGallery
	'oGalleryitem.getgalleryitem
%>
<script language="javascript">

	//메인배너 등록상품 상품찾기
	function popItemWindow(tgf){
		var popup_item = window.open("/common/pop_singleItemSelect.asp?target=" + tgf, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
		popup_item.focus();
	}
	
	//메인배너 상품 등록
	function regmainbanneritem()
	{
		if (searchForm.itemid.value==''){
			alert('상품코드를입력하세요');
			searchForm.itemid.focus();
		}else{
			moveForm.action="/admin/artist/artist_process.asp";
			moveForm.mode.value="mainbanneritem";
			moveForm.itemid.value = searchForm.itemid.value;
			moveForm.submit();
		}
	}

	function goPage(pg)
	{
		frm = document.moveForm;
		frm.action="";
		frm.page.value=pg;
		frm.submit();
	}

	function addItem()
	{
		frm = document.moveForm;
		frm.action="artist_gallery_edit.asp";
		frm.mode.value="add";
		frm.submit();
	}
	
	//입점문의
	function inquiry(){
		var inquiry = window.open('/admin/artist/artist_inquiry.asp','inquiry','width=1024,height=768,scrollbars=yes,resizable=yes');
		inquiry.focus();
	}	
	
	//아티스트추천관리
	function recommend(){
		var recommend = window.open('/admin/artist/artist_recommend.asp','recommend','width=1024,height=768,scrollbars=yes,resizable=yes');
		recommend.focus();
	}	

	function editItem(sn)
	{
		frm = document.moveForm;
		frm.action="artist_gallery_edit.asp";
		frm.mode.value="edit";
		frm.page.value="<%=page%>";
		frm.gal_sn.value=sn;
		frm.submit();
	}

</script>
<!-- #include virtual="/admin/brand/inc_streetHead.asp"-->

<img src="/images/icon_arrow_link.gif"> <b>Artist Work</b>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="searchForm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 브랜드 : 
		<%' Call DrawSelectBoxUseBrand("designerid",designerid) %>
		<% drawSelectBoxDesignerwithName "designerid",designerid %>		
		&nbsp;&nbsp;
		* 브랜드구분 : <% drawSelectBoxbrandgubun "brandgubun",brandgubun , "" %>		
		&nbsp; &nbsp;
		* 아티스트구분 :
		<select name="gal_div" class="select">
			<option value=""<% if gal_div="" then Response.Write " selected" %>>선택</option>
			<option value="W"<% if gal_div="W" then Response.Write " selected" %>>Work</option>
			<option value="D"<% if gal_div="D" then Response.Write " selected" %>>Drawing</option>
			<option value="P"<% if gal_div="P" then Response.Write " selected" %>>Photo</option>
		</select>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="searchForm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 대표카테고리 : 
		기능<% SelectBoxBrandCategory "catecode", catecode %>
		전시<%= fnStandardDispCateSelectBox(1,"", "standardCateCode", standardCateCode, "")%>
		&nbsp;&nbsp;
		* 담당MD : <% drawSelectBoxCoWorker_OnOff "mduserid", mduserid, "on" %>
		&nbsp; &nbsp;
		* 사용유무 : <select name="isusing" class="select"><option value="Y">Yes</option><option value="N">No</option></select>
		<script language="javascript">
			document.searchForm.isusing.value="<%=isusing%>";
		</script>		
	</td>
</tr>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<!--
<tr>
	<td align="left">
		<font color="red">※Mainpage 하단배너6개(최근등록상품이첫번째로노출)<br></font>
		<% for lp = 0 to oGalleryitem.ftotalcount - 1 %>
		<img src="<%= oGalleryitem.fitemlist(lp).flistimage120 %>" border=0 width=40 height=40>
		<% next %>
		상품코드 : <input type="text" name="itemid" size=10>
		<input type="button" class="button" value="찾기" onClick="popItemWindow('searchForm')">			
		<input type="button" class="button" value="저장" onClick="regmainbanneritem()">					
	</td>
	<td align="right">	
	</td>
</tr>
-->
<tr>
	<td align="left">	
	</td>
	<td align="right">	
		<input type="button" value="아이템 추가" onclick="addItem()" class="button">
	</td>
</tr>	
</form>
<form name="moveForm" method="GET">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="gal_sn" value="">
	<input type="hidden" name="isusing" value="<%=isusing%>">
	<input type="hidden" name="gal_div" value="<%=gal_div%>">
	<input type="hidden" name="designerid" value="<%=designerid%>">
	<input type="hidden" name="itemid" size=10>
</form>
</table>
<!-- 액션 끝 -->
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oGallery.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oGallery.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50" align="center">번호</td>
	<td width="100" align="center">구분</td>
	<td width="250" align="center">업체명</td>
	<td align="center">이미지</td>
	<td width="50" align="center">사용유무</td>
	<td width="80" align="center">등록일</td>
</tr>
<% if oGallery.FResultCount>0 then %>
<% for lp=0 to oGallery.FResultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
		<td align="center"><%= oGallery.FItemList(lp).Fgal_sn %></td>
		<td align="center"><%= oGallery.FItemList(lp).getGalDivName %></td>
		<td align="center"><%= oGallery.FItemList(lp).Fsocname_kor & "(" & oGallery.FItemList(lp).Fsocname & ")" %></td>
		<td align="center">
			<a href="javascript:editItem(<%= oGallery.FItemList(lp).Fgal_sn %>)">
			<img src="<%= oGallery.FItemList(lp).Fgal_img400 %>" width=50 height="50" border="0">
			</a>
		</td>
		<td align="center"><%= oGallery.FItemList(lp).Fgal_isusing %></td>
		<td align="center"><%= FormatDateTime(oGallery.FItemList(lp).Fgal_regdate,2) %></td>
    </tr>   
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if oGallery.HasPreScroll then %>
			<a href="javascript:goPage(<%= oGallery.StartScrollPage-1 %>)">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
	
		<% for lp=0 + oGallery.StartScrollPage to oGallery.FScrollCount + oGallery.StartScrollPage - 1 %>
			<% if lp>oGallery.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(lp) then %>
			<font color="red">[<%= lp %>]</font>
			<% else %>
			<a href="javascript:goPage(<%= lp %>)">[<%= lp %>]</a>
			<% end if %>
		<% next %>
	
		<% if oGallery.HasNextScroll then %>
			<a href="javascript:goPage(<%= lp %>)">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
	set oGallery = Nothing
	set oGalleryitem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
