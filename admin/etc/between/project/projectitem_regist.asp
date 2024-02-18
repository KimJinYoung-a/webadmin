<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/between/projectcls.asp"-->
<%
Dim opjt, pjt_code
pjt_code = Request("pjt_code")

If pjt_code = "" Then
%>
<script language="javascript">
	alert("전달값에 문제가 발생하였습니다. 관리자에게 문의해주십시오");
	history.back();
</script>
<%	dbget.close()	:	response.End
End If

Dim pjt_name, pjt_gender, pjt_state, pjt_sortType
Dim strG, strSort

strG  = Request("selG")
strSort  = Request("selSort")

SET opjt = new cProject
	opjt.FRectPjt_code = pjt_code
	opjt.getProjectCont()
	pjt_name		= opjt.FItemList(0).FPjt_name
	pjt_gender		= opjt.FItemList(0).FPjt_gender
	pjt_state		= opjt.FItemList(0).FPjt_state
	pjt_sortType	= opjt.FItemList(0).FPjt_sortType
%>

<script language="javascript">
// 새상품 추가 팝업
function addnewItem(){
	var popwin;
	popwin = window.open("/admin/etc/between/project/pop_project_additemlist.asp?pjt_code=<%=pjt_code%>", "popup_item", "width=1500,height=768,scrollbars=yes,resizable=yes");
	popwin.focus();
}

//전체선택
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
		//check optioon
		var e = frm.elements[i];
	
		//check itemEA		
		if ((e.type=="checkbox")) {				
			e.checked = blnChk ;
		}
	}
}

//삭제
function jsDel(sType, iValue){	
	var frm;		
	var sValue;		
	frm = document.fitem;
	sValue = "";
	
	if (sType ==0) {
		if(!frm.chkI) return;
		
		if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){
			if(frm.chkI[i].checked){
			   	if (sValue==""){
					sValue = frm.chkI[i].value;		
			   	}else{
					sValue =sValue+","+frm.chkI[i].value;		
			   	}	
			}
		}	
		}else{
			if(frm.chkI.checked){
				sValue = frm.chkI.value;
			}	
		}
	
		if (sValue == "") {
			alert('선택 상품이 없습니다.');
			return;
		}
		document.frmDel.itemidarr.value = sValue;
	}else{
		document.frmDel.itemidarr.value = iValue;
	}	
	 
	if(confirm("선택하신 상품을 삭제하시겠습니까?")){		
		document.frmDel.submit();
	}
}
//그룹검색
function jsSearchGroup(){
	document.fitem.submit();	
}
//정렬
function jsChSort(){
	document.fitem.submit();	
}

//그룹이동	

function addGroup(){
	var frm, sValue, sGroup;

	frm = document.fitem;
	sValue = "";
	sGroup =frm.eG.options[frm.eG.selectedIndex].value ;
			
	if(!frm.chkI) return;
	if(!sGroup){
	 alert("이동할 그룹이 없습니다.");
	 return;
	}
	
	if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){
			if(frm.chkI[i].checked){
			   if (sValue==""){
				sValue = frm.chkI[i].value;		
				}else{
				sValue =sValue+","+frm.chkI[i].value;		
				}
			}
		}	
	}else{
		sValue = frm.chkI.value;
	}
	
	if (sValue == "") {
		alert('선택 상품이 없습니다.');
		return;
	}
	document.frmG.selGroup.value = frm.eG.options[frm.eG.selectedIndex].value;
	document.frmG.itemidarr.value = sValue;
	document.frmG.submit();
}

// 상품 순서 일괄 저장
function jsSortSize() {
	var frm;
	var sValue, sSort
	frm = document.fitem;
	sValue = "";
	sSort = "";
		
	if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){
			if(!IsDigit(frm.sSort[i].value)){
				alert("순서지정은 숫자만 가능합니다.");
				frm.sSort[i].focus();
				return;
			}

			if (sValue==""){
				sValue = frm.chkI[i].value;		
			}else{
				sValue =sValue+","+frm.chkI[i].value;		
			}	
			
			// 정렬순서
			if (sSort==""){
				sSort = frm.sSort[i].value;		
			}else{
				sSort =sSort+","+frm.sSort[i].value;		
			}
		}
	}else{
		sValue = frm.chkI.value;
		if(!IsDigit(frm.sSort.value)){
			alert("순서지정은 숫자만 가능합니다.");
			frm.sSort.focus();
			return;
		}
		sSort =  frm.sSort.value; 
	}
	document.frmSortSize.itemidarr.value = sValue;
	document.frmSortSize.sortarr.value = sSort;
	document.frmSortSize.submit();
}
function goPage(pg) {
    document.fitem.page.value = pg;
    document.fitem.submit();
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td style="padding-bottom:10">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">기획전코드</td>
			<td width="30%" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%= pjt_code %></td>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">기획전명</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%= pjt_name %></td>
		</tr>
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">종류</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%= getDBcodeByName(opjt.FItemList(0).FPjt_kind) %></td>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">상태</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%= getDBcodeByName(opjt.FItemList(0).FPjt_state) %></td>
		</tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">성별</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<% 
					Select Case pjt_gender
						Case "A"	response.write "전체"
						Case "M"	response.write "남자"
						Case "F"	response.write "여자"
					End Select
				%>
			</td>
			<td colspan="2" bgcolor="#FFFFFF"></td>
		</tr>
		</table>
	</td>
</tr>
<%
SET opjt = nothing

Dim cPjtGroup, i
SET cPjtGroup = new cProject
	cPjtGroup.FRectPjt_code = pjt_code
	cPjtGroup.getProjectItemGroup()
%>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="0" class="a">
		<form name="fitem" method="post" action="projectitem_regist.asp">
		<input type="hidden" name="page" value="">
		<input type="hidden" name="pjt_code" value="<%=pjt_code%>">
		<input type="hidden" name="menupos" value="<%=menupos%>">
		<input type="hidden" name="mode" value="">
		<input type="hidden" name="selGroup" value="">
		<tr align="center"  >
			<td align="left">
	        	 그룹검색
	        	<select name="selG" onChange="jsSearchGroup();">
	        	<option value="">전체</option>
	       	<% If cPjtGroup.FResultCount > 0 Then %>
	       		<option value="0"  <%IF Cstr(strG) = "0" THEN%>selected<%END IF%>>미지정</option>
	       	<%
	       		For i = 0 to cPjtGroup.FResultCount - 1
	       	%>
	       		<option value="<%=cPjtGroup.FItemList(i).FPjtgroup_code%>" <%IF Cstr(strG) = Cstr(cPjtGroup.FItemList(i).FPjtgroup_code) THEN %> selected<%END IF%>> <%=cPjtGroup.FItemList(i).FPjtgroup_code%>(<%=cPjtGroup.FItemList(i).FPjtgroup_desc%>)</option>
	    	<%	Next
	    	END IF%>
	       	</select>
	        </td>
	        <td align="right">
	         정렬 : <select name="selSort" onchange="jsChSort();">
	       		<option value="sitemid" >신상품순</option>
	       		<option value="sevtitem" <%IF Cstr(strSort) = "sevtitem" THEN %>selected<%END IF%>>순서순</option>
	       		<option value="sbest" <%IF Cstr(strSort) = "sbest" THEN %>selected<%END IF%>>베스트셀러순</option>
	       		<option value="shsell" <%IF Cstr(strSort) = "shsell" THEN %>selected<%END IF%>>높은가격순</option>
	       		<option value="slsell" <%IF Cstr(strSort) = "slsell" THEN %>selected<%END IF%>>낮은가격순</option>
	       		<option value="sevtgroup" <%IF Cstr(strSort) = "sevtgroup" THEN %>selected<%END IF%>>그룹순</option>
	       		<option value="sbrand" <%IF Cstr(strSort) = "sbrand" THEN %>selected<%END IF%>>브랜드</option>
	       		</select>
	        </td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
		<tr height="35">
			<td align="left">
				<input type="button" value="선택삭제" onClick="jsDel(0,'');" class="button">&nbsp;&nbsp;&nbsp;
				<select name="eG">
		<%
			If cPjtGroup.FResultCount > 0 Then
				For i = 0 to cPjtGroup.FResultCount - 1
		%>
					<option value=" <%=cPjtGroup.FItemList(i).FPjtgroup_code%>" ><%IF cPjtGroup.FItemList(i).FPjtgroup_pcode <> 0 THEN%>└&nbsp;<%END IF%><%=cPjtGroup.FItemList(i).FPjtgroup_code%>(<%=cPjtGroup.FItemList(i).FPjtgroup_desc%>)</option>
		<%
				Next
			ELSE
		%>
					<option value=""> --그룹없음--</option>
		<% END IF %>
				</select>
				<input type="button" value="선택그룹이동" onClick="addGroup();" class="button">
			</td>
			<td align="right">
				<input type="button" value="순서 저장" onClick="jsSortSize();" class="button">&nbsp;
				<input type="button" value="새상품 추가" onclick="addnewItem();" class="button">
			</td>
		</tr>
		</table>
	</td>
</tr>
<%
SET cPjtGroup = nothing

Dim cPjtGroupItem, page
page    = request("page")
If page = "" Then page = 1

SET cPjtGroupItem = new cProject
	cPjtGroupItem.FPageSize 	= 20
	cPjtGroupItem.FCurrPage		= page
	cPjtGroupItem.FRectPjt_code = pjt_code
	cPjtGroupItem.FRectSGroup 	= strG
	cPjtGroupItem.FRectSort		= strSort
	cPjtGroupItem.getProjectItem()
%>
<tr>
	<td>
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF">
			<td colspan="16" align="left">검색결과 : <b><%= FormatNumber(cPjtGroupItem.FTotalCount,0) %></b>&nbsp;&nbsp;페이지 : <b><%= FormatNumber(page,0) %> / <%= FormatNumber(cPjtGroupItem.FTotalPage,0) %></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>
			<td>그룹코드</td>
			<td align="center">상품ID</td>
			<td align="center">이미지</td>
			<td align="center">브랜드</td>
			<td align="center">상품명</td>
			<td align="center">비트윈 상품명</td>
			<td align="center">판매가</td>
			<td align="center">매입가</td>
			<td align="center">배송</td>
			<td align="center">마진</td>
			<td align="center">판매여부</td>
			<td align="center">사용여부</td>
			<td align="center">한정여부</td>
			<td>순서</td>
			<td>처리</td>
		</tr>
<%
	If cPjtGroupItem.FResultCount > 0 Then
    	For i = 0 to cPjtGroupItem.FResultCount - 1
%>
		<tr align="center" bgcolor="#FFFFFF">
			<td><input type="checkbox" name="chkI" value="<%= cPjtGroupItem.FItemList(i).FItemid %>"></td>
			<td><%= Chkiif(cPjtGroupItem.FItemList(i).FPjtgroup_code <> 0, cPjtGroupItem.FItemList(i).FPjtgroup_code, "") %></td>
			<td>
				<A href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= cPjtGroupItem.FItemList(i).FItemid %>" target="_blank"><%= cPjtGroupItem.FItemList(i).FItemid %></a>
			<% If cPjtGroupItem.IsSoldOut(cPjtGroupItem.FItemList(i).FSellyn, cPjtGroupItem.FItemList(i).FLimityn, cPjtGroupItem.FItemList(i).FLimitno, cPjtGroupItem.FItemList(i).FLimitsold) Then %>
				<br><img src="http://webadmin.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
			<% End If %>
			</td>
	    	<td>
	    		<% If (Not IsNull(cPjtGroupItem.FItemList(i).FSmallimage) ) and (cPjtGroupItem.FItemList(i).FSmallimage <> "") Then %>
					<img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid( cPjtGroupItem.FItemList(i).FItemid )%>/<%=cPjtGroupItem.FItemList(i).FSmallimage%>">
				<% End If %>
			</td>
			<td><%=db2html(cPjtGroupItem.FItemList(i).FMakerid)%></td>
			<td align="left">&nbsp;<%=db2html(cPjtGroupItem.FItemList(i).FItemname)%></td>
			<td align="left">&nbsp;<%=db2html(cPjtGroupItem.FItemList(i).FChgItemname)%></td>
			<td>
			<%
				Response.Write FormatNumber(cPjtGroupItem.FItemList(i).FOrgprice,0)
				'할인가
				If cPjtGroupItem.FItemList(i).FSailyn="Y" then
					Response.Write "<br><font color=#F08050>(할)" & FormatNumber(cPjtGroupItem.FItemList(i).FSailprice,0) & "</font>"
				End If
				'쿠폰가
				If cPjtGroupItem.FItemList(i).FItemcouponyn = "Y" Then
					Select Case cPjtGroupItem.FItemList(i).FItemcoupontype
						Case "1"
							Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(cPjtGroupItem.FItemList(i).FOrgprice * ((100 - cPjtGroupItem.FItemList(i).FItemcouponvalue) / 100), 0) & "</font>"
						Case "2"
							Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(cPjtGroupItem.FItemList(i).FOrgprice - cPjtGroupItem.FItemList(i).FItemcouponvalue, 0) & "</font>"
					End Select
				End If
			%>
			</td>
	    	<td>
			<%
				Response.Write FormatNumber(cPjtGroupItem.FItemList(i).FOrgsuplycash,0)
				'할인가
				If cPjtGroupItem.FItemList(i).FSailyn = "Y" Then
					Response.Write "<br><font color=#F08050>" & FormatNumber(cPjtGroupItem.FItemList(i).FSailsuplycash,0) & "</font>"
				End If
				'쿠폰가
				If cPjtGroupItem.FItemList(i).FItemcouponyn = "Y" Then
					If cPjtGroupItem.FItemList(i).FItemcoupontype = "1" OR cPjtGroupItem.FItemList(i).FItemcoupontype = "2" Then
					End If
				End If
			%>
			</td>
	    	<td><%= fnColor(cPjtGroupItem.IsUpcheBeasong(cPjtGroupItem.FItemList(i).FDeliverytype),"delivery")%></td>
	    	<td>
    		<%
				If cPjtGroupItem.FItemList(i).Fsellcash<>0 Then
					response.write CLng(10000-cPjtGroupItem.FItemList(i).Fbuycash/cPjtGroupItem.FItemList(i).Fsellcash*100*100)/100 & "%"
				End If
			%>
	    	</td>
	    	<td><%= fnColor(cPjtGroupItem.FItemList(i).FSellyn, "yn") %></td>
	    	<td><%= fnColor(cPjtGroupItem.FItemList(i).FIsusing, "yn") %></td>
	    	<td><%= fnColor(cPjtGroupItem.FItemList(i).FLimityn, "yn") %></td>
	    	<td><input type="text" name="sSort" value="<%=cPjtGroupItem.FItemList(i).FPjtitem_sort%>" size="4" style="text-align:right;"></td>
	    	<td><input type="button" value="삭제" onClick="jsDel(1,<%= cPjtGroupItem.FItemList(i).FItemid %>);" class="button"></td>
	    </tr>
<%
	   Next
%>
		<tr height="20">
			<td colspan="17" align="center" bgcolor="#FFFFFF">
			<% If cPjtGroupItem.HasPreScroll Then %>
				<a href="javascript:goPage('<%= cPjtGroupItem.StartScrollPage-1 %>');">[pre]</a>
			<% Else %>
				[pre]
			<% End If %>
			<% For i=0 + cPjtGroupItem.StartScrollPage to cPjtGroupItem.FScrollCount + cPjtGroupItem.StartScrollPage - 1 %>
				<% If i > cPjtGroupItem.FTotalpage Then Exit For %>
				<% If CStr(page) = CStr(i) Then %>
					<font color="red">[<%= i %>]</font>
				<% Else %>
					<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
				<% End If %>
			<% Next %>
			<% If cPjtGroupItem.HasNextScroll Then %>
				<a href="javascript:goPage('<%= i %>');">[next]</a>
			<% Else %>
				[next]
			<% End If %>
			</td>
		</tr>
<%
	   	ELSE
%>
		<tr align="center" bgcolor="#FFFFFF">
			<td height="50" colspan="16">등록된 내용이 없습니다.</td>
		</tr>
	   <% END IF %>
		</form>
		</table>
	</td>
</tr>
</table>
<% SET cPjtGroupItem = nothing %>
<!-- 선택삭제--->
<form name="frmDel" method="post" action="projectitem_process.asp">
<input type="hidden" name="mode" value="D">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="pjt_code" value="<%=pjt_code%>">
<input type="hidden" name="selG" value="<%=strG%>">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<!-- 그룹이동--->
<form name="frmG" method="post" action="projectitem_process.asp">
<input type="hidden" name="mode" value="G">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="pjt_code" value="<%=pjt_code%>">
<input type="hidden" name="selG" value="<%=strG%>">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="selGroup" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<!-- 순서 및 이미지크기 변경--->
<form name="frmSortSize" method="post" action="projectitem_process.asp">
<input type="hidden" name="mode" value="S">
<input type="hidden" name="pjt_code" value="<%=pjt_code%>">
<input type="hidden" name="selG" value="<%=strG%>">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="sortarr" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->