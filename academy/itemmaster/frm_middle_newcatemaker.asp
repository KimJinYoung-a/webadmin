<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYCategoryCls.asp"-->
<%
'###############################################
' PageName : frm_middle_newcatemaker.asp
' Discription : DIY샾 카테고리 변경 페이지
' History : 2010.09.16 허진원
'###############################################

dim cdl,cdm,cds
cdl = RequestCheckvar(request("cdl"),10)
cdm = RequestCheckvar(request("cdm"),10)
cds = RequestCheckvar(request("cds"),10)

dim oLcate
set oLcate = new CCatemanager
oLcate.GetNewCateMaster


dim oMcate
set oMcate = new CCatemanager
if (cdl<>"") then
	oMcate.GetNewCateMasterMid cdl
end if

dim oScate
set oScate = new CCatemanager
if (cdl<>"") and (cdm<>"") then
	oScate.GetNewCateMasterSmall cdl,cdm
end if

dim i,currposStr

if cdl<>"" then
	currposStr = oLcate.GetNewCateCurrentPos(cdl,cdm,cds)
end if
%>
<script language='javascript'>
function popNewCategory(cdl,cdm){
	var popwin = window.open('popNewCate.asp?cdl=' + cdl + '&cdm=' + cdm,'popnewcate','width=400,height=300,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function TnCategoryEdit(cdl,cdm,cds,odn,nm){
	var popwin = window.open('popEditCate.asp?cdl=' + cdl + '&cdm=' + cdm + '&cds=' + cds,'popeditcate','width=400,height=300,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function TnCategoryDel(cdl,cdm,cds,mode){
	var strMsg;
	if(mode=="mdel") {
		strMsg = "중분류 카테고리를 삭제하시겠습니까?\n\n※ 중분류 카테고리는 속해있는 소분류 카테고리가 없어야 삭제할 수 있습니다.\n 그리고 관련된 카테고리 속성은 함께 삭제됩니다.";
	} else if(mode=="sdel") {
		strMsg = "소분류 카테고리를 삭제하시겠습니까?\n\n※ 기본 카테고리로 등록된 상품이 없어야 삭제할 수 있습니다.\n그리고 추가 카테고리로 등록된 상품은 연결이 해제됩니다.";
	} else {
		return;
	}

	if (confirm(strMsg)){
		 var popwin = window.open('popDelCate.asp?cdl=' + cdl + '&cdm=' + cdm + '&cds=' + cds + '&mode=' + mode,'popdelcate','width=400,height=300,resizable=yes,scrollbars=yes');
		 popwin.focus();
	}
}
function MakeCateMenu(cdl,cdm){
	if (confirm("카테고리를 메인페이지 메뉴에 적용하시겠습니까?")){
	    var popwin = window.open('<%= wwwFingers %>/chtml/make_diyShopCate_menu2010.asp?cdl=' + cdl,'popnewcate','width=400,height=300,resizable=yes,scrollbars=yes');
		popwin.focus();
	}
}
function AvailCategory(){
<% if cds="" then %>
	return "";
<% else %>
	return "<%= cdl + cdm + cds + currposStr %>";
<% end if %>
}
</script>
<table border=0 cellspacing=0 cellpadding=0 class=a>
<tr>
	<td width="300">현재위치 : <%= currposStr %></td>
	<td><input type="button" value="대분류추가" onclick="popNewCategory('','')"></td>
	<td>
		<% if cdl<>"" then %>
		<input type="button" value="중분류추가" onclick="popNewCategory('<%= cdl %>','')">
		<% else %>

		<% end if %>
	</td>
	<td>
		<% if (cdl<>"") and (cdm<>"") then %>
		<input type="button" value="소분류추가" onclick="popNewCategory('<%= cdl %>','<%= cdm %>')">
		<% else %>

		<% end if %>
	</td>
	<td><input type="button" value="Menu적용<%= ChkIIF(cdl<>"","[" & cdl & "]","") %>" onclick="MakeCateMenu('<%= cdl %>')" <%= ChkIIF(cdl="","disabled","") %> ></td>
</tr>
</table>
<table border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#FFFFFF" class="a">사용안하는 카테고리 삭제,순서 정렬한후 우측상당 <font color="blue">MENU적용</font> 버튼을 눌러주세요.</td>
	</tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" >
	<tr>
		<td valign=top>
			<table border=1 cellspacing=1 cellpadding=0 class=a width=150>
			<% for i=0 to oLcate.FResultCount-1 %>
			<tr>
				<% if oLcate.FItemList(i).Fcdlarge=cdl then %>
				<td><b><a href="?cdl=<%= oLcate.FItemList(i).Fcdlarge %>">[<%= oLcate.FItemList(i).Fcdlarge %>]<%= oLcate.FItemList(i).Fnmlarge %></a></b></td>
				<% else %>
				<td><a href="?cdl=<%= oLcate.FItemList(i).Fcdlarge %>">[<%= oLcate.FItemList(i).Fcdlarge %>]<%= oLcate.FItemList(i).Fnmlarge %></a></td>
				<% end if %>
			</tr>
			<% next %>
			</table>
		</td>
		<td valign=top>
			<table border=1 cellspacing=1 cellpadding=1 class=a width=160>
			<% for i=0 to oMcate.FResultCount-1 %>
			<tr>
				<% if oMcate.FItemList(i).Fcdmid=cdm then %>
					<td><%= oMcate.FItemList(i).ForderNo %></td>
					<td><b><a href="?cdl=<%= oMcate.FItemList(i).Fcdlarge %>&cdm=<%= oMcate.FItemList(i).Fcdmid %>">[<%= oMcate.FItemList(i).Fcdmid %>]<%= oMcate.FItemList(i).Fnmlarge %></a>&nbsp;[<a href="javascript:TnCategoryEdit('<%= oMcate.FItemList(i).Fcdlarge %>','<%= oMcate.FItemList(i).Fcdmid %>','','<%= oMcate.FItemList(i).ForderNo %>','<%= oMcate.FItemList(i).Fnmlarge %>')">E</a>]&nbsp;[<a href="javascript:TnCategoryDel('<%= oMcate.FItemList(i).Fcdlarge %>','<%= oMcate.FItemList(i).Fcdmid %>','','mdel')">D</a>]</b></td>
				<% else %>
					<td><%= oMcate.FItemList(i).ForderNo %></td>
					<td><a href="?cdl=<%= oMcate.FItemList(i).Fcdlarge %>&cdm=<%= oMcate.FItemList(i).Fcdmid %>">[<%= oMcate.FItemList(i).Fcdmid %>]<%= oMcate.FItemList(i).Fnmlarge %></a>&nbsp;[<a href="javascript:TnCategoryEdit('<%= oMcate.FItemList(i).Fcdlarge %>','<%= oMcate.FItemList(i).Fcdmid %>','','<%= oMcate.FItemList(i).ForderNo %>','<%= oMcate.FItemList(i).Fnmlarge %>')">E</a>]&nbsp;[<a href="javascript:TnCategoryDel('<%= oMcate.FItemList(i).Fcdlarge %>','<%= oMcate.FItemList(i).Fcdmid %>','','mdel')">D</a>]</td>
				<% end if %>
			</tr>
			<% next %>
			</table>
		</td>
		<td valign=top>
			<table border=1 cellspacing=1 cellpadding=1 class=a width=150>
			<% for i=0 to oScate.FResultCount-1 %>
			<tr>
			<% if oScate.FItemList(i).Fcdsmall=cds then %>
				<td><%= oScate.FItemList(i).ForderNo %></td>
				<td><b><a href="?cdl=<%= oScate.FItemList(i).Fcdlarge %>&cdm=<%= oScate.FItemList(i).Fcdmid %>&cds=<%= oScate.FItemList(i).Fcdsmall %>">[<%= oScate.FItemList(i).Fcdsmall %>]<%= oScate.FItemList(i).Fnmlarge %></a></b>&nbsp;[<a href="javascript:TnCategoryEdit('<%= oScate.FItemList(i).Fcdlarge %>','<%= oScate.FItemList(i).Fcdmid %>','<%= oScate.FItemList(i).Fcdsmall %>','<%= oScate.FItemList(i).ForderNo %>','<%= oScate.FItemList(i).Fnmlarge %>')">E</a>]&nbsp;[<a href="javascript:TnCategoryDel('<%= oScate.FItemList(i).Fcdlarge %>','<%= oScate.FItemList(i).Fcdmid %>','<%= oScate.FItemList(i).Fcdsmall %>','sdel')">D</a>]</td>
			<% else %>
				<td><%= oScate.FItemList(i).ForderNo %></td>
				<td><a href="?cdl=<%= oScate.FItemList(i).Fcdlarge %>&cdm=<%= oScate.FItemList(i).Fcdmid %>&cds=<%= oScate.FItemList(i).Fcdsmall %>">[<%= oScate.FItemList(i).Fcdsmall %>]<%= oScate.FItemList(i).Fnmlarge %></a>&nbsp;[<a href="javascript:TnCategoryEdit('<%= oScate.FItemList(i).Fcdlarge %>','<%= oScate.FItemList(i).Fcdmid %>','<%= oScate.FItemList(i).Fcdsmall %>','<%= oScate.FItemList(i).ForderNo %>','<%= oScate.FItemList(i).Fnmlarge %>')">E</a>]&nbsp;[<a href="javascript:TnCategoryDel('<%= oScate.FItemList(i).Fcdlarge %>','<%= oScate.FItemList(i).Fcdmid %>','<%= oScate.FItemList(i).Fcdsmall %>','sdel')">D</a>]</td>
			<% end if %>
				<td width=20><%= oScate.FItemList(i).Fcatecnt %></td>
			</tr>
			<% next %>
			</table>
		</td>
		<td width=330>
		<iframe name=imatchitem src="imatchitem.asp?cdl=<%= cdl %>&cdm=<%= cdm %>&cds=<%= cds %>" width=330 height=600></iframe>
	</td>
</tr>
</table>

<%
set oLcate = Nothing
set oMcate = Nothing
set oScate = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->