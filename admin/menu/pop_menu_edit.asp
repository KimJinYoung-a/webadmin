<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<%
	Dim page, SearchKey, SearchString, midd

	midd = requestCheckvar(Request("mid"),10)
	page = requestCheckvar(Request("page"),10)
	SearchKey = requestCheckvar(Request("SearchKey"),32)
	SearchString = Request("SearchString")
	if page="" then page=1

''신규메뉴 관리로 이동 2016/08/04 eastone
    response.redirect "/admin/menu/menu_edit.asp?mid="&midd
    dbget.close()
    response.end

	'// 내용 접수
	dim oMenu, lp
	Set oMenu = new CMenuList

	oMenu.FRectMid = midd
	
	oMenu.GetMenuCont
%>
<script language="javascript">
<!--
	// 권한 선택 팝업
	function popAuthSelect()
	{
		window.open("pop_Menu_auth.asp?mid=<%=midd%>", "popMenuAuth","width=360,height=200,scrollbars=no");
	}

	// 팝업에서 선택권한 추가
	function addAuthItem(psn,pnm,lsn,lnm)
	{
		var lenRow = tbl_auth.rows.length;

		// 기존에 값에 중복 파트 여부 검사
		if(lenRow>1)	{
			for(l=0;l<document.all.part_sn.length;l++)	{
				if(document.all.part_sn[l].value==psn) {
					alert("이미 권한이 지정된 부서입니다.\n기존 부서를 삭제하고 권한을 다시 지정해주세요.");
					return;
				}
			}
		}
		else {
			if(lenRow>0) {
				if(document.all.part_sn.value==psn) {
					alert("이미 권한이 지정된 부서입니다.\n기존 부서를 삭제하고 권한을 다시 지정해주세요.");
					return;
				}
			}
		}

		// 행추가
		var oRow = tbl_auth.insertRow(lenRow);
		oRow.onmouseover=function(){tbl_auth.clickedRowIndex=this.rowIndex};

		// 셀추가 (부서,등급,삭제버튼)
		var oCell1 = oRow.insertCell(0);		
		var oCell2 = oRow.insertCell(1);
		var oCell3 = oRow.insertCell(2);

		oCell1.innerHTML = pnm + "<input type='hidden' name='part_sn' value='" + psn + "'>";
		oCell2.innerHTML = lnm + "<input type='hidden' name='level_sn' value='" + lsn + "'>";
		oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delAuthItem()' align=absmiddle>";
	}

	// 선택권한 삭제
	function delAuthItem()
	{
		if(confirm("선택한 권한을 삭제하시겠습니까?"))
			tbl_auth.deleteRow(tbl_auth.clickedRowIndex);
	}

	// 폼검사 및 실행
	function submitForm()
	{
		var form = document.frm;

		if(!form.viewIdx.value||!IsDigit(form.viewIdx.value))
		{
			alert("표시순서를 정수로 입력해주십시오.");
			form.viewIdx.focus();
			return;
		}

		if(!form.menuname.value)
		{
			alert("메뉴명을 입력해주십시오.");
			form.menuname.focus();
			return;
		}
		if(!form.parentid.value)
		{
			alert("상위메뉴를 선택해주십시오.\n\n※상위메뉴가 없을경우 루트메뉴를 선택해주십시오.");
			form.parentid.focus();
			return;
		}

//		if(tbl_auth.rows.length<=0)
//		{
//			alert("메뉴에 접근할 수 있는 권한을 [추가]버튼을 눌러 지정하여주십시오.");
//			return;
//		}

/* 권한승계기능 없앰
		if(form.parentid.value=='0')
		{
			if(confirm("하위메뉴가 있을경우 권한을 승계하시겠습니까?\n\n[확인]:예, [취소]:아니오"))
				form.childYN.value="Y";
			else
				form.childYN.value="N";
		}
*/

        form.childYN.value="N";
        
		if(confirm("입력한 내용으로 저장하시겠습니까?"))
		{
			form.action="menu_process.asp";
			form.submit();
		}
		else
		{		
			return;
		}
	}
//-->
</script>
<script language="javascript" src="colorbox.js"></script>
<form name="frm" method="POST" action="" style="margin:0px;">
<input type="hidden" name="mode" value="popmodi">
<input type="hidden" name="mid" value="<%=midd%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="SearchKey" value="<%=SearchKey%>">
<input type="hidden" name="SearchString" value="<%=SearchString%>">
<input type="hidden" name="childYN" value="N">
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td colspan="2" align="center" bgcolor="#FFFFFF">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td bgcolor="#FFFFFF"><img src="/images/icon_star.gif" align="absmiddle"> <b>메뉴 상세보기/수정</b></td>
		</tr>
		</table>
	</td>
</tr>
<%
	if oMenu.FResultCount=0 then
%>
<tr>
	<td colspan="4" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 메뉴가 없습니다.</td>
</tr>
<%
	else
%>
<tr>
	<td width="100" bgcolor="#E6E6E6" align="center">일련번호</td>
	<td bgcolor="#FFFFFF"><b><%=oMenu.FitemList(1).Fmenu_id%></b></td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">표시순서</td>
	<td bgcolor="#FFFFFF"><input type="text" class='text' name="viewIdx" size="5" value="<%=oMenu.FitemList(1).Fmenu_viewIdx%>"></td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">메뉴명</td>
	<td bgcolor="#FFFFFF"><input type="text" class='text' name="menuname" size="40" value="<%=oMenu.FitemList(1).Fmenu_name%>"></td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">링크URL</td>
	<td bgcolor="#FFFFFF"><input type="text" class='text' name="linkurl" size="60" value="<%=oMenu.FitemList(1).Fmenu_linkurl%>">
	<input type="checkbox" name="useSslYN" value="Y" <% if (oMenu.FitemList(1).Fmenu_useSslYN = "Y") then %>checked<% end if %> > SSL 사용
		&nbsp;
		<input type="checkbox" name="criticinfo" value="1" <% if (oMenu.FitemList(1).Fmenu_criticinfo = "1") then %>checked<% end if %> > 개인정보등 민감정보포함
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">표시색상</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="prvColor" readonly style="background-color:'<%=oMenu.FitemList(1).Fmenu_color%>';width:21px;height:21px;border:1px solid #606060;cursor:pointer;" onClick="ShowColorBox(event.clientX, event.clientY+document.body.scrollTop)">
		<input type="text" class='text_ro' name="menucolor" size="7" maxlength="7" value="<%=oMenu.FitemList(1).Fmenu_color%>" readonly onClick="ShowColorBox(event.clientX, event.clientY+document.body.scrollTop)" style="cursor:pointer">
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">상위메뉴</td>
	<td bgcolor="#FFFFFF"><%=printRootMenuOption("parentid",oMenu.FitemList(1).Fmenu_parentid, "NoAction")%></td>
</tr>

<tr>
	<td bgcolor="#E6E6E6" align="center">지정권한</td>
	<td bgcolor="#FFFFFF">
		<table class=a>
		<tr>
			<td><%=getPartLevelInfo(oMenu.FitemList(1).Fmenu_id,"modi")%></td>
			<td valign="bottom"><input type="button" class='button' value="추가" onClick="popAuthSelect()"></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">사용여부</td>
	<td bgcolor="#FFFFFF">
		<select class="select" name="isUsing">
			<option value="Y">사용</option>
			<option value="N">삭제</option>
		</select>
		<script language="javascript">frm.isUsing.value='<%=oMenu.FitemList(1).Fmenu_isUsing%>';</script>
	</td>
</tr>
<tr>
    <td bgcolor="#E6E6E6" align="center">(기존권한)</td>
    <td bgcolor="#EEEEEE">
        <% DrawAuthBox "divcd",oMenu.FitemList(1).Fmenu_divcd %>
        (업체, 제휴사, 강사, 매장 /admin/ 폴더가 아닌곳.)
    </td>
</tr>
<%
	end if
%>
<tr height="25">
	<td colspan="2" align="center" bgcolor="#FFFFFF">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="center">
				<a href="javascript:submitForm();"><img src="/images/icon_confirm.gif" width="45" border="0" align="absmiddle"></a> &nbsp;
				<a href="javascript:self.close();"><img src="/images/icon_cancel.gif" width="45" border="0" align="absmiddle"></a>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<!-- 메인 내용 끝 -->
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->