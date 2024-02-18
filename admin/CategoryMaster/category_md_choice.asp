<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/category_md_choicecls.asp"-->
<%
dim page, cdl, cdm, isusing, vCateCode, vIdx
cdl = request("cdl")
cdm = request("cdm")
page = request("page")
isusing = request("isusing")
vCateCode = Request("catecode")
vIdx = request("idx")

if page="" then page=1
if isusing = "" Then isusing="Y"

dim omd
set omd = New CMDChoice
omd.FCurrPage = page
omd.FPageSize = 10
omd.FRectDisp1 = vCateCode
omd.FRectIsUsing = isusing
omd.GetMDChoiceThemeList

dim i
%>
<script language='javascript'>
<!--
function popItemWindow(idx){
	var popupitem = window.open("category_md_choice_itempop.asp?catecode=<%=vCateCode%>&idx="+idx+"", "popupitem", "width=1000,height=800,scrollbars=yes,resizable=yes");
	popupitem.focus();
}

function viewTheme(idx){
	frm.idx.value = idx;
	frm.page.value = <%=page%>;
	frm.submit();
}

function editTheme(idx){
	newtheme.location.href = "category_md_choice_newtheme.asp?disp1=<%=vCateCode%>&idx="+idx+"";
}

function RefreshCategoryEventBanner(){
	if(confirm("적용하시겠습니까?") == true) {
		 var mdchoice = window.open('<%=CHKIIF(application("Svr_Info")="Dev",wwwURL,"http://www1.10x10.co.kr")%>/chtml/dispcate/catemain_mdpick_make.asp?catecode=<%=vCateCode%>','mdchoice','');
		 mdchoice.focus();
	}
}

//-->
</script>
<form name="refreshFrm" method="post">
<input type="hidden" name="catecode" value="<%=vCateCode%>">
</form>
<!-- 상단 검색폼 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="idx" value="">
<table width="1000" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" height="40">
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">&nbsp;
		전시카테고리 : 
		<%
		Dim cDisp
		SET cDisp = New cDispCate
		cDisp.FCurrPage = 1
		cDisp.FPageSize = 2000
		cDisp.FRectDepth = 1
		'cDisp.FRectUseYN = "Y"
		cDisp.GetDispCateList()
		
		If cDisp.FResultCount > 0 Then
			Response.Write "<select name=""catecode"" class=""select"" onChange=""frm.submit();"">" & vbCrLf
			Response.Write "<option value="""">선택</option>" & vbCrLf
			For i=0 To cDisp.FResultCount-1
				Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """ " & CHKIIF(CStr(vCateCode)=CStr(cDisp.FItemList(i).FCateCode),"selected","") & ">" & cDisp.FItemList(i).FCateName & "</option>"
			Next
			Response.Write "</select>&nbsp;&nbsp;&nbsp;"
		End If
		Set cDisp = Nothing
		%>
		&nbsp;&nbsp;
		사용유무 :
		<select name="isusing" onchange="frm.submit();" class="select">
			<option value=""  <% if isusing="" then response.write "selected" %>>전체</option>
			<option value="Y" <% if isusing="Y" then response.write "selected" %>>사용</option>
			<option value="N" <% if isusing="N" then response.write "selected" %>>사용안함</option>
		</select>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검 색">
	</td>
</tr>
<%IF vCateCode <> "" THEN%>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="50" colspan="3">&nbsp;
		<a href="javascript:RefreshCategoryEventBanner()"><img src="/images/refreshcpage.gif" width="19" height="23" border="0" align="absmiddle"><b>Real 적용</b></a>
	</td>
</tr>
<%END IF%>
</table>
</form>
<% If vCateCode <> "" Then %>
<iframe src="category_md_choice_newtheme.asp?disp1=<%=vCateCode%>&idx=<%=vIdx%>" name="newtheme" id="newtheme" frameborder="0" width="900px" height="50px"></iframe>
<% End If %>
<style type="text/css">
	html, body, blockquote, caption, dd, div, dl, dt, fieldset, form, frame, iframe, input, legend, li, object, ol, p, pre, q, select, table, textarea, tr, td, ul, button {margin:0; padding:0; font-size:12px; font-family:verdana, tahoma, dotum, dotumche, '돋움', '돋움체', sans-serif; line-height:1.5; color:#555;}
	ol, ul {list-style:none;}
	.overHidden {overflow:hidden;}
	.ftLt {float:left;}
	.ftRt {float:right;}
	.tPad10 {padding-top:10px;}
	.bPad10 {padding-bottom:10px;}
	.ctgyMainAdminWrap {width:908px; /*padding-left:232px;*/ position:relative;}
	.mdpickWrap {border-top:5px solid #eee; border-bottom:1px solid #eee; overflow:hidden; _zoom:1;}
	.mdpickWrap .ftLt {padding:20px 15px; width:180px;}
	.mdpickWrap .ftLt li {padding:3px 0; text-align:right; font-size:11px;}
	.mdpickWrap .ftLt li a {text-decoration:none; color:#555;}
	.mdpickWrap .ftLt li.on {font-weight:bold;}
	.mdpickWrap .ftRt {border-left:1px solid #eee; width:680px;}
	.mdpickList {overflow:hidden; _zoom:1; width:660px; padding:0 20px;}
	.mdpickList li {float:left; padding:20px; width:120px; height:120px; display:table; text-align:center;}
	.mdpickList li div {background:#f5f5f5; width:120px; height:120px; display:table-cell; vertical-align:middle;}
	.mdpickList li div input[type=button] {padding:5px 15px;}
</style>
<%
If vIdx <> "" Then
	Dim vQuery, vTheme, vArr
	vQuery = "select * from db_sitemaster.dbo.tbl_category_MDChoice_theme where idx = '" & vIdx & "'"
	rsget.Open vQuery,dbget,1
	vTheme = db2html(rsget("subject"))
	rsget.close()
	
	vArr = 0
	vQuery = "select Top 8 i.itemid, (select icon2image from db_item.dbo.tbl_item where itemid = i.itemid) from db_sitemaster.dbo.tbl_category_MDChoice as i where i.theme_idx = '" & vIdx & "' and i.isusing = 'Y' order by i.sortNo asc, i.itemid desc "
	rsget.Open vQuery,dbget,1
	If not rsget.Eof Then
		vArr = rsget.getRows()
	End If
	rsget.close()
%>
<table bgcolor="#FFFFFF" border="1">
<tr>
	<td>
	<div class="ctgyMainAdminWrap">
		<p class="bPad10 tPad10"><img src="http://fiximage.10x10.co.kr/web2013/shopping/contit_mdpick.gif" alt="MD's PICK" /></p>
		<div class="mdpickWrap">
			<div class="ftLt">
				<ul>
					<li class="on"><%=vTheme%></li>
				<ul>
			</div>
			<div class="ftRt">
				<ul class="mdpickList">
				<% If isArray(vArr) THEN %>
					<% For i=0 To 7 %>
						<% If UBound(vArr,2) < i Then %>
							<li><div></div></li>
						<% Else %>
							<li><div><img src="http://webimage.10x10.co.kr/image/icon2/<%= GetImageSubFolderByItemid(vArr(0,i)) %>/<%= vArr(1,i) %>" width="120px" height="120px"></div></li>
						<% End If %>
					<% Next %>
				<% Else %>
					<li><div></div></li><li><div></div></li><li><div></div></li><li><div></div></li><li><div></div></li><li><div></div></li><li><div></div></li><li><div></div></li>
				<% End If %>
				</ul>
			</div>
		</div>
	</div>
	</td>
</tr>
</table>
<br>
<% End If %>
<!-- 검색 끝 -->
<% If vCateCode <> "" Then %>
<table width="1000" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">idx</td>
	<td align="center">테마</td>
	<td align="center">정렬순서</td>
	<td align="center">사용유무</td>
	<td align="center">등록일</td>
	<td align="center"></td>
</tr>
<% for i=0 to omd.FResultCount-1 %>
<tr bgcolor="#FFFFFF" height="20">
	<td align="center" style="cursor:pointer;" onClick="viewTheme('<%= omd.FItemList(i).Fidx %>');"><%= omd.FItemList(i).Fidx %></td>
	<td align="center" style="cursor:pointer;" onClick="viewTheme('<%= omd.FItemList(i).Fidx %>');"><%= omd.FItemList(i).Fsubject %></td>
	<td align="center" style="cursor:pointer;" onClick="viewTheme('<%= omd.FItemList(i).Fidx %>');"><%= omd.FItemList(i).FsortNo %></td>
	<td align="center" style="cursor:pointer;" onClick="viewTheme('<%= omd.FItemList(i).Fidx %>');"><%= omd.FItemList(i).Fisusing %></td>
	<td align="center" style="cursor:pointer;" onClick="viewTheme('<%= omd.FItemList(i).Fidx %>');"><%= omd.FItemList(i).Fregdate %></td>
	<td align="center">
		<input type="button" value="테마수정" onClick="editTheme('<%= omd.FItemList(i).Fidx %>');" class="button">&nbsp;
		<input type="button" value="상품관리" onClick="popItemWindow('<%= omd.FItemList(i).Fidx %>');" class="button">
	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" align="center">
	<% if omd.HasPreScroll then %>
		<a href="?page=<%= omd.StarScrollPage-1 %>&catecode=<%=vCateCode%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + omd.StarScrollPage to omd.FScrollCount + omd.StarScrollPage - 1 %>
		<% if i>omd.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&catecode=<%=vCateCode%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if omd.HasNextScroll then %>
		<a href="?page=<%= i %>&catecode=<%=vCateCode%>&isusing=<%=isusing%>&menupos=<%= menupos %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<% End If %>
<%
set omd = Nothing
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
