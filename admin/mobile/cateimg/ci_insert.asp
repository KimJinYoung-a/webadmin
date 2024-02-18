<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : nb_insert.asp
' Discription : 모바일 사이트 알림배너
' History : 2013.04.01 이종화
'			2013.12.15 한용민 수정
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/catebanner.asp" -->
<%
Dim subImage1 , isusing , mode, oCateImgOne, idx, dispCate
Dim kword1 , kword2 , kword3 , kwordurl1 , kwordurl2 , kwordurl3
	idx = requestCheckvar(request("idx"),16)
	menupos = requestCheckvar(request("menupos"),10)

If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 

set oCateImgOne = new CMainbanner
	oCateImgOne.FRectIdx = idx
	
	if idx<>"" then
		oCateImgOne.GetOneContents()
	end if
	
	if oCateImgOne.FResultCount > 0 then
		dispCate = oCateImgOne.FOneItem.Fcatecode
		isusing = oCateImgOne.FOneItem.Fisusing
		subImage1 = oCateImgOne.FOneItem.Fcateimg
		idx = oCateImgOne.FOneItem.fidx
		kword1 = oCateImgOne.FOneItem.fkword1
		kword2 = oCateImgOne.FOneItem.fkword2
		kword3 = oCateImgOne.FOneItem.fkword3
		kwordurl1 = oCateImgOne.FOneItem.fkwordurl1
		kwordurl2 = oCateImgOne.FOneItem.fkwordurl2
		kwordurl3 = oCateImgOne.FOneItem.fkwordurl3
	end if
set oCateImgOne = Nothing
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>

	function jsSubmit(){
		var frm = document.frm;
	
		if (!frm.disp.value){
			alert('카테고리를 선택해주세요');
			frm.disp.focus();
			return;
		}

		if (confirm('저장 하시겠습니까?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	
	function jsgolist(){
		self.location.href="/admin/mobile/cateimg/";
	}
	
	function putLinkText(key,gubun) {
		var frm = document.frm;
		var kword
		var urllink
		if (gubun == "1" )
		{
			urllink = frm.kwordurl1;
			kword = frm.kword1.value;
		}else if( gubun == "2"){
			urllink = frm.kwordurl2;
			kword = frm.kword2.value;
		}else{
			urllink = frm.kwordurl3;
			kword = frm.kword3.value;
		}
		switch(key) {
			case 'search':
				urllink.value='/search/search_result.asp?rect='+kword;
				break;
			case 'event':
				urllink.value='/event/eventmain.asp?eventid=이벤트번호';
				break;
			case 'itemid':
				urllink.value='/category/category_itemprd.asp?itemid=상품코드';
				break;
			case 'category':
				urllink.value='/category/category_list.asp?disp=카테고리';
				break;
			case 'brand':
				urllink.value='/street/street_brand.asp?makerid=브랜드아이디';
				break;
		}
	}
</script>
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/doCateimage.asp" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="menupos" value="<%=menupos%>">

<% If mode = "modify" then%>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="100">번호</td>
	<td>
		<%= idx %>	<font color="red">※수정시에는 카테고리 변경이 불가능 합니다 . 이미지만 변경 해주세요 ※</font>
	</td>
</tr>
<% End If %>

<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="100">카테고리</td>
	<td>
		<!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">카테고리이미지</td>
	<td>
		<input type="file" name="subImage1" class="file" title="이미지 #1" require="N" style="width:80%;" />
		<% if subImage1<>"" then %>
		<br>
		<img src="<%= subImage1 %>" width="100" /><br><%= subImage1 %>
		<% end if %>		
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">키워드1</td>
	<td><input type="text" name="kword1" value="<%=kword1%>" size="20" maxlength="20"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">키워드1 URL</td>
	<td><input type="text" name="kwordurl1" size="80" value="<%=kwordurl1%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('search','1')">검색결과 링크 : /search/search_result.asp?rect=<font color="darkred">검색어</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('event','1')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','1')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('category','1')">카테고리 링크 : /category/category_list.asp?disp=<font color="darkred">카테고리</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('brand','1')">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">키워드2</td>
	<td><input type="text" name="kword2" value="<%=kword2%>" size="20" maxlength="20"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">키워드2 URL</td>
	<td><input type="text" name="kwordurl2" size="80" value="<%=kwordurl2%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('search','2')">검색결과 링크 : /search/search_result.asp?rect=<font color="darkred">검색어</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('event','2')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','2')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('category','2')">카테고리 링크 : /category/category_list.asp?disp=<font color="darkred">카테고리</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('brand','2')">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">키워드3</td>
	<td><input type="text" name="kword3" value="<%=kword3%>" size="20" maxlength="20"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">키워드3 URL</td>
	<td><input type="text" name="kwordurl3" size="80" value="<%=kwordurl3%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('search','3')">검색결과 링크 : /search/search_result.asp?rect=<font color="darkred">검색어</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('event','3')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','3')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('category','3')">카테고리 링크 : /category/category_list.asp?disp=<font color="darkred">카테고리</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('brand','3')">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">사용여부</td>
	<td><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />사용함 &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>사용안함</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2"><input type="button" value=" 취 소 " onClick="jsgolist();" class="button" /><input type="button" value=" 저 장 " onClick="jsSubmit();" class="button" /></td>
</tr>
</form>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->