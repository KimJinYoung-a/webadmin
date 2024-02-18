<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/common/uploadimage/uploadimageCls.asp"-->
<%
	Dim i, cUpImg, vItemID, vItemName, keyword, vMakerID, vSiteGubun, vCurrpage, vNoUp, vSortNo
	vCurrpage = requestCheckVar(NullFillWith(Request("cpg"), "1"),10)
	vItemID = requestCheckVar(request("itemid"),200)
	vItemName = requestCheckVar(request("itemname"),150)
	vMakerID = requestCheckVar(request("makerid"),50)
	vNoUp = requestCheckVar(NullFillWith(Request("noup"), ""),1)
	vSortNo = requestCheckVar(NullFillWith(Request("sortno"), ""),1)

	vSiteGubun = "china"

	if vItemID<>"" then
		dim iA ,arrTemp,arrItemid
		vItemID = replace(vItemID,",",chr(10))
		vItemID = replace(vItemID,chr(13),"")
		arrTemp = Split(vItemID,chr(10))

		iA = 0
		do while iA <= ubound(arrTemp) 
			if trim(arrTemp(iA))<>"" then
				'상품코드 유효성 검사(2008.08.05;허진원)
				if Not(isNumeric(trim(arrTemp(iA)))) then
					Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
					dbget.close()	:	response.End
				else
					arrItemid = arrItemid & trim(arrTemp(iA)) & ","
				end if
			end if
			iA = iA + 1
		loop

		vItemID = left(arrItemid,len(arrItemid)-1)
	end if
	
	Set cUpImg = New cUploadImage
	cUpImg.FPageSize = 20
	cUpImg.FCurrPage = vCurrpage
	cUpImg.FRectSiteGubun = vSiteGubun
	cUpImg.FRectItemID = vItemID
	cUpImg.FRectMakerId = vMakerID
	cUpImg.FRectItemName = vItemName
	cUpImg.FRectNoUp = vNoUp
	cUpImg.FRectSortNo = vSortNo
	cUpImg.sbUploadImageMngList
%>

<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script>
function jsGoCommonImgUp(idx){
	location.href = "<%=uploadImgUrl%>/linkweb/common/common_upload_image.asp?sitegubun=<%=vSiteGubun%>&topidx="+idx+"";
}

function searchFrm(p){
	frmitem.cpg.value = p;
	frmitem.submit();
}

function jsGoSortNo(s){
	frmitem.cpg.value = 1;
	frmitem.sortno.value = s;
	frmitem.submit();
}
</script>
<form name="frmitem" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
<input type="hidden" name="search" value="o">
<input type="hidden" name="cpg" value="1">
<input type="hidden" name="sortno" value="<%=vSortNo%>">
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr>
	<td bgcolor="#FFFFFF">
		<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
		<tr>
			<td bgcolor="#FFFFFF">
				<table class="a">
				<tr>
					<td>브랜드 : <% drawSelectBoxDesignerWithName "makerid", vMakerID %></td> 
					<td><span style="padding-left:10px">상품명 :</span> <input type="text" class="text" name="itemname" value="<%= vItemName %>" size="35" maxlength="20">
						<span style="font-size:11px; color:gray;padding-left:5px;">(주의:느릴수있습니다.)</span>
						<!--<span style="padding-left:10px">검색키워드 :</span> <input type="text" class="text" name="keyword" value="<%=keyword%>" size="35">
						<span style="font-size:11px; color:gray;padding-left:5px;">(주의:느릴수있습니다.)</span>//-->
						<span style="padding-left:10px"><label id="nola"><input type="checkbox" name="noup" id="nola" value="o" <%=CHKIIF(vNoUp="o","checked","")%>> 이미지 등록 안된것</label></span>
					</td>
				</tr>
				</table>	
			</td>
		</tr> 
		<tr>
			<td  bgcolor="#FFFFFF" style="padding:7 0 7 0;">
				<span style="padding-left:3px">상품코드 :</span> <textarea rows="3" cols="50" name="itemid" id="itemid"><%=replace(vItemID,",",chr(10))%></textarea>
				<span style="font-size:11px; color:gray;padding-left:5px;">(엔터로 복수입력가능)</span>
			</td>
		</tr>
		<tr>
			<td bgcolor="#D4FFFF"  style="padding:7 0 7 0;">
				<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td>
						<strong>&nbsp;Total : <%=FormatNumber(cUpImg.FTotalCount,0)%></strong>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
	<td bgcolor="#FFFFFF" width="10%" align="center">
		<table class="a">
		<tr>
			<td align="center"><input type="button" value="검 색" onClick="searchFrm('1');" style="width:60px;height:60px;"></td>
		</tr>
		<tr>
			<td align="center" style="padding-top:15px;"></td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<br>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td bgcolor="#FFFFFF">
		<select name="sortno" onChange="jsGoSortNo(this.value);">
			<option value="" <%=CHKIIF(vSortNo="","selected","")%>>-정렬-</option>
			<option value="1" <%=CHKIIF(vSortNo="1","selected","")%>>이미지등록 최신순</option>
			<option value="2" <%=CHKIIF(vSortNo="2","selected","")%>>이미지등록 오래된순</option>
			<option value="3" <%=CHKIIF(vSortNo="3","selected","")%>>상품코드 최신순</option>
			<option value="4" <%=CHKIIF(vSortNo="4","selected","")%>>상품코드 오래된순</option>
		</select>
	</td>
	<td align="right">
		<input type="button" value="새로운 이미지 올리기" style="height:30px;" onClick="jsGoCommonImgUp('');">
	</td>
</tr>
</table>
<br>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr align="center" bgcolor="#F3F3FF">
	<td width="70">IDX</td>
	<td width="100"></td>
	<td width="100">브랜드명<br>[브랜드ID]</td>
	<td width="100">상품코드</td>
	<td>상품명</td>
	<td width="130">등록된 이미지 수</td>
	<td width="160">등록일</td>
	<td width="70"></td>
</tr>
<%
If cUpImg.FResultCount = 0 Then
%>
	<tr>
		<td colspan="8" height="30" bgcolor="#FFFFFF" align="center">검색된 상품이 없습니다.</td>
	</tr>
<%
Else
	For i=0 To cUpImg.FResultCount-1
%>
	<tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer;">
		<td align="center"><%=cUpImg.FItemList(i).FIdx%></td>
		<td align="center"><img src="<%=cUpImg.FItemList(i).FListImage100%>"></td>
		<td align="center"><%=cUpImg.FItemList(i).FBrandName%><br>[<%=cUpImg.FItemList(i).FMakerID%>]</td>
		<td align="center"><%=cUpImg.FItemList(i).FOptInt%></td>
		<td> <%=cUpImg.FItemList(i).FItemName%></td>
		<td align="center"><%=cUpImg.FItemList(i).FRegImgCnt%></td>
		<td align="center"><%=cUpImg.FItemList(i).FRegdate%></td>
		<td align="center" onClick="jsGoCommonImgUp('<%=cUpImg.FItemList(i).FIdx%>');">수  정</td>
	</tr>
<%
	Next
%>
	<tr height="50" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if cUpImg.HasPreScroll then %>
			<a href="javascript:searchFrm('<%= cUpImg.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + cUpImg.StartScrollPage to cUpImg.FScrollCount + cUpImg.StartScrollPage - 1 %>
    			<% if i>cUpImg.FTotalpage then Exit for %>
    			<% if CStr(vCurrpage)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:searchFrm('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if cUpImg.HasNextScroll then %>
    			<a href="javascript:searchFrm('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
<%
End If
%>
</table>
<% Set cUpImg = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->