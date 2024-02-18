<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  업체리스트
' History : 2009.04.07 서동석 생성
'			2012.09.06 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
dim ogroup, frmname, page, rectconame, rectDesigner, rectsocno, groupid , ceoname ,i, isusing, vTmpGr, vGrArr, vItemTotalCount
	frmname     = request("frmname")
	page        = requestCheckVar(request("page"),9)
	rectconame  = requestCheckVar(request("rectconame"),32)
	rectDesigner = requestCheckVar(request("rectDesigner"),32)
	rectsocno   = requestCheckVar(request("rectsocno"),16)
	groupid   = requestCheckVar(request("groupid"),16)
	ceoname     = request("ceoname")
	isusing = requestCheckVar(request("isusing"),1)

if page="" then page=1

'### 전체 상품수 나타낼때 사용. 꼭 필요하답니다;	'/2017.04.26 강준구 추가(양정모이사님 지시)	'/2017.04.26 한용민 주석처리(양정모이사님이 다시 빼달라고 하심)
'vItemTotalCount = fnITemTotalCount()

set ogroup = new CPartnerGroup
	ogroup.FPageSize = 30
	ogroup.FCurrPage = page
	ogroup.FrectDesigner = rectDesigner
	ogroup.Frectconame = rectconame
	ogroup.FRectsocno = rectsocno
	ogroup.FRectGroupid = groupid
	ogroup.FRectceoname = ceoname
	ogroup.FRectIsusing = isusing
	
	if rectDesigner<>"" then
		ogroup.GetGroupInfoListByBrand
	else
		ogroup.GetGroupInfoList
	end if

	vTmpGr = ogroup.FGroupIdList
	If vTmpGr <> "" Then
		vTmpGr = Left(vTmpGr, Len(vTmpGr)-1)
		ogroup.FGroupIdList = vTmpGr
		vGrArr = ogroup.fnGroupInfoByItemCount
	End If
%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value=page;
    document.frm.submit();
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		회사명 : <input type="text" name="rectconame" class="text" value="<%= rectconame %>" size=10 maxlength=32>
		&nbsp;&nbsp;
		그룹코드 : <input type="text" name="groupid" class="text" value="<%= groupid %>" size=8 maxlength=6>
    	&nbsp;&nbsp;
    	사업자번호 : <input type="text" name="rectsocno" class="text" value="<%= rectsocno %>" size=15 maxlength=12>
    	&nbsp;&nbsp;
    	포함브랜드 : <input type="text" name="rectDesigner" class="text" value="<%= rectDesigner %>" Maxlength="32" size="16">
    	대표자명 : <input type="text" name="ceoname" class="text" value="<%= ceoname %>" Maxlength="8" size="8">
    	&nbsp;&nbsp;
    	업체검색 :
    	<select name="isusing">
    		<option value="" <%=CHKIIF(isusing="","selected","")%>>전체</option>
    		<option value="Y" <%=CHKIIF(isusing="Y","selected","")%>>사용중</option>
    		<option value="N" <%=CHKIIF(isusing="N","selected","")%>>종료</option>
    	</select>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>

<br>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ogroup.FtotalCount %> 건</b>
		&nbsp;
		페이지 : <b><%= page %> / <%= ogroup.FTotalpage %></b>
		<!--&nbsp;
		상품 총 수 : <b><%'=FormatNumber(vItemTotalCount,0)%></b> (판매여부상관없이 사용중인것)-->
	</td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
	<td width="60" >업체코드<br>(그룹코드)</td>
	<td width="130" >회사명</td>
	<td width="80" >사업자번호</td>
	<td width="60" >대표자</td>
	<td width="130" >전화번호<br>팩스번호</td>
	<td width="80" >담당자</td>
	<% if (FALSE) then %>
	<td>핸드폰번호<br>이메일주소</td>
    <% end if %>
	<td>진행브랜드</td>
	<td>상품수</td>
	<!-- <td >브랜드사용여부</td> -->
</tr>
<% if ogroup.FResultCount >0 then %>
<% for i=0 to ogroup.FResultCount -1 %>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= ogroup.FItemList(i).FGroupID %></td>
	<td><a href="javascript:PopUpcheInfoEdit('<%= ogroup.FItemList(i).FGroupID %>')"><%= ogroup.FItemList(i).Fcompany_name %></a></td>
	<td align="center"><%= socialnoReplace(ogroup.FItemList(i).Fcompany_no) %></td>
	<td align="center"><%= ogroup.FItemList(i).Fceoname %></td>
	<td>TEL : <%= ogroup.FItemList(i).Fcompany_tel %><br>FAX : <%= ogroup.FItemList(i).Fcompany_fax %></td>
	<td align="center"><%= ogroup.FItemList(i).Fmanager_name %></td>
	<% if (FALSE) then %>
	<td>H.P : <%= ogroup.FItemList(i).Fmanager_phone %><br>E-mail : <%= ogroup.FItemList(i).Fmanager_email %></td>
	<% end if %>
	<td <%=ChkIIF(ogroup.FItemList(i).getPartnerIdInfoStr="","bgcolor='#CCCCCC'","")%> ><%= ogroup.FItemList(i).getPartnerIdInfoStr %></td>
	<td align="center"><%=fnGroupListItemCntView(vGrArr,ogroup.FItemList(i).FGroupID)%> 개</td>
	<!-- <td></td> -->
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
	<% if ogroup.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ogroup.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + ogroup.StartScrollPage to ogroup.FScrollCount + ogroup.StartScrollPage - 1 %>
		<% if i>ogroup.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if ogroup.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan=10 align=center>[ 검색결과가 없습니다. ]</td>
</tr>
<% end if %>
</table>

<%
set ogroup = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->



