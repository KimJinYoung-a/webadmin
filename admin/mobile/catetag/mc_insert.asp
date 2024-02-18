<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : mc_insert.asp
' Discription : 모바일 사이트 카테고리 태그
' History : 2014-09-02 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/catetag.asp" -->
<%
Dim subImage1 , isusing , mode, oCatetagOne, idx, dispCate , appdiv , appcate
Dim kword1 , kword2 , kword3 , kwordurl1 , kwordurl2 , kwordurl3
	idx = requestCheckvar(request("idx"),16)
	menupos = requestCheckvar(request("menupos"),10)

If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 

set oCatetagOne = new CMaincatetag
	oCatetagOne.FRectIdx = idx
	
	if idx<>"" then
		oCatetagOne.GetOneContents()
	end if
	
	if oCatetagOne.FResultCount > 0 then
		dispCate = oCatetagOne.FOneItem.Fcatecode
		isusing = oCatetagOne.FOneItem.Fisusing
		idx = oCatetagOne.FOneItem.fidx
		kword1 = oCatetagOne.FOneItem.fkword1
		kwordurl1 = oCatetagOne.FOneItem.fkwordurl1
		kwordurl2 = oCatetagOne.FOneItem.fkwordurl2
		appdiv = oCatetagOne.FOneItem.fappdiv
		appcate = oCatetagOne.FOneItem.fappcate
	end if
set oCatetagOne = Nothing
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
		self.location.href="/admin/mobile/catetag/?menupos=<%=menupos%>";
	}
	
	function putLinkText(key) {
		var frm = document.frm;
		var kword
		var urllink
			urllink = frm.kwordurl1;
			kword = frm.kword1.value;
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
				urllink.value='/category/category_list.asp?cdl=카테고리';
				break;
			case 'brand':
				urllink.value='/street/street_brand.asp?makerid=브랜드아이디';
				break;
		}
	}

	//url 자동 생성
	function chklink(v){
		if (v == "1"){
			document.frm.kwordurl2.value = "/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid=상품코드";
			$("#catesel").css("display","none");
			$("#kwordurl2").prop('disabled',false);
		}else if (v == "2"){
			document.frm.kwordurl2.value = "/apps/appcom/wish/web2014/event/eventmain.asp?eventid=이벤트코드&rdsite=rdsite명(필수아님)";
			$("#catesel").css("display","none");
			$("#kwordurl2").prop('disabled',false);
		}else if (v == "3"){
			document.frm.kwordurl2.value = "makerid=브랜드명";
			$("#catesel").css("display","none");
			$("#kwordurl2").prop('disabled',false);
		}else if (v == "4"){
			chgDispCate2('');
			document.frm.kwordurl2.value = "cd1=&nm1=";
			$("#catesel").css("display","block");
			$("#kwordurl2").attr('readonly','readonly');
		}else{
			document.frm.kwordurl2.value = "APP URL 구분을 선택 해주세요.";
			$("#catesel").css("display","none");
			$("#kwordurl2").prop('disabled',false);
		}
	}
</script>
<script>
function chgDispCate2(dc) {
	$.ajax({
		url: "dispCateSelectBox_response.asp?disp="+dc,
		cache: false,
		async: false,
		success: function(message) {
			// 내용 넣기
			$("#lyrDispCtBox2").empty().html(message);
			if (dc.length == 3){
				document.frm.kwordurl2.value = $("#dispcateval1 option:selected").val()+"||"+$("#dispcateval1 option:selected").text();
				$("#appcate").val(dc);
			}else if (dc.length == 6){
				document.frm.kwordurl2.value = $("#dispcateval1 option:selected").val()+"||"+$("#dispcateval2 option:selected").val()+"||"+$("#dispcateval1 option:selected").text()+"||"+$("#dispcateval2 option:selected").text();
				$("#appcate").val(dc);
			}else if (dc.length == 9){
				document.frm.kwordurl2.value = $("#dispcateval1 option:selected").val()+"||"+$("#dispcateval2 option:selected").val()+"||"+$("#dispcateval3 option:selected").val()+"||"+$("#dispcateval1 option:selected").text()+"||"+$("#dispcateval2 option:selected").text()+"||"+$("#dispcateval3 option:selected").text();
				$("#appcate").val(dc);
			}else{
				
			}

		}
	});
}
$(function(){
	chgDispCate2('<%=appcate%>');
});
</script>
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="docatetag.asp" onSubmit="return jsRegCode();">
<input type="hidden" name="mode" value="<%=mode%>"/>
<input type="hidden" name="idx" value="<%=idx%>"/>
<input type="hidden" name="menupos" value="<%=menupos%>"/>
<input type="hidden" name="appcate" id="appcate"/>
<% If mode = "modify" then%>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="100">번호</td>
	<td>
		<%= idx %>
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
	<td bgcolor="#FFF999" align="center" width="10%">인기 키워드</td>
	<td><input type="text" name="kword1" value="<%=kword1%>" size="20" maxlength="20"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">모바일웹용 URL</td>
	<td><input type="text" name="kwordurl1" size="80" value="<%=kwordurl1%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('search')">검색결과 링크 : /search/search_result.asp?rect=<font color="darkred">검색어</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('event')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid')">상품코드 링크 : /category/category_itemprd.asp?itemid=<font color="darkred">상품코드 (O)</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('category')">카테고리 링크 : /category/category_list.asp?cdl=<font color="darkred">카테고리</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('brand')">브랜드아이디 링크 : /street/street_brand.asp?makerid=<font color="darkred">브랜드아이디</font></span>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">APP용 URL</td>
	<td>
		<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="#3d3d3d">
			<tr>
				<td bgcolor="#FFF999" width="100" align="center">APP URL 구분</td>
				<td bgcolor="#FFFFFF">
					<select name='appdiv' class='select' onchange="chklink(this.value);">
						<option value="0">선택하세요</option>
						<option value="1" <% if appdiv = "1" then response.write " selected" %>>상품상세</option>
						<option value="2" <% if appdiv = "2" then response.write " selected" %>>이벤트</option>
						<option value="3" <% if appdiv = "3" then response.write " selected" %>>브랜드</option>
						<option value="4" <% if appdiv = "4" then response.write " selected" %>>카테고리</option>
					</select>
				</td>
			</tr>
			<tr id="catesel" style="display:<%=chkiif(idx<>"" And appdiv = "4","block","none")%>">
				<td bgcolor="#FFF999" width="100" align="center">전시카테고리 선택</td>
				<td bgcolor="#FFFFFF">
					<span id="lyrDispCtBox2"></span>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFF999" width="100" align="center">코드내용</td>
				<td bgcolor="#FFFFFF"><textarea name="kwordurl2" class="textarea" id="kwordurl2" style="width:100%; height:40px;"><%=kwordurl2%></textarea></td>
			</tr>
		</table>
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