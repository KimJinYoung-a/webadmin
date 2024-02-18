<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->
<!-- #include virtual="/lib/classes/etc/outmallConfirmCls.asp"-->
<%
Dim i
Dim oOutMall
SET oOutMall = new cOutmall
	oOutMall.FCurrPage			= 1
	oOutMall.FPageSize			= 100
	oOutMall.getAllOutmallList

	Dim strSql, arrList, vAction, intLoop, vMallGubun, vTotalCount
	Dim targetmall, arrtargetmall
	Dim targetmall2, tmpMallList
	vAction 		= Request("action")
	vMallGubun 		= NullFillWith(Request("mallgubun"),"")
	targetmall		= Request("targetmall")
	arrtargetmall	= Request("arrtargetmall")
	targetmall2		= Request("targetmall2")

	if (vMallGubun="lotte") then vMallGubun="lotteCom"   ''' 20130304 추가

	If vAction = "insert" OR vAction = "delete" Then
		Call Proc()
	End If

	strSql = ""
	strSql = strSql & " SELECT Count(j.idx) "
	strSql = strSql & " FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] as j "
	strSql = strSql & " JOIN db_item.dbo.tbl_Cate_large as l on j.cdl = l.code_large "
	strSql = strSql & " JOIN db_item.dbo.tbl_Cate_mid as m on j.cdl = m.code_large and j.cdm = m.code_mid "
	strSql = strSql & " Where 1=1  "
	strSql = strSql & " and j.mallgubun = '"& vMallGubun &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		vTotalCount = rsget(0)
	rsget.close

	strSql = ""
	strSql = strSql & " SELECT j.idx, j.mallgubun, j.cdl, j.cdm, l.code_nm, m.code_nm "
	strSql = strSql & " FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] as j "
	strSql = strSql & " JOIN db_item.dbo.tbl_Cate_large as l on j.cdl = l.code_large "
	strSql = strSql & " JOIN db_item.dbo.tbl_Cate_mid as m on j.cdl = m.code_large and j.cdm = m.code_mid "
	strSql = strSql & " Where 1=1  "
	strSql = strSql & " and j.mallgubun = '"& vMallGubun &"' "
	strSql = strSql & " order by j.cdl, j.cdm "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	IF not rsget.EOF THEN
		arrList = rsget.getRows()
	END IF
	rsget.close
%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function insert_id()
{
	var val = GetValueOfRadio('targetmall');
	var chk;
	var i, arrtargetmall;
	if(frm.cdl.value == "")
	{
		alert("대카테고리를 선택하세요.");
		frm.in_id.focus();
		return;
	}

	if(frm.cdm.value == "")
	{
		alert("중카테고리를 선택하세요.");
		frm.in_id.focus();
		return;
	}

	arrtargetmall = '';
	if (val == 'sel') {
		for (i = 0; ; i++) {
			chk = document.getElementById("chk" + i);
			if (chk == undefined) { break; }

			if (chk.checked == true) {
				arrtargetmall = arrtargetmall + ',' + chk.value;
			}
		}

		if (arrtargetmall == '') {
			alert('대상몰을 선택하세요.');
			return;
		}
	}
	frm.arrtargetmall.value = arrtargetmall;

	frm.action.value = "insert";
	frm.submit();
}
function delete_idx()
{
	frm.action.value = "delete";
	frm.submit();
}

window.onload = function() {
	window.resizeTo(750, 770);

	// 제외상품 대상몰
	jsShowHideTR();
	// 검색 대상몰
	jsShowHideTR2();
}

function GetValueOfRadio(name) {
	var radios = document.getElementsByName(name);
	for (var i = 0; i < radios.length; i++) {
		if (radios[i].checked == true) { return radios[i].value; }
	}
	return '';
}

function jsShowHideTR() {
	var row = $('.showhide');
	var frm = document.frm;
	var val = GetValueOfRadio('targetmall');
	if (val == 'sel') {
		row.show();
	} else {
		row.hide();
	}
}

function jsShowHideTR2() {
	var row = $('.showhide2');
	var frm = document.frm;
	var val = GetValueOfRadio('targetmall2');
	if (val == 'sel') {
		row.show();
	} else {
		row.hide();
	}
}
</script>
<center>Mall 구분 : <b><%=vMallGubun%></b></center>
<br>
<%
Dim reIdName, reIdName2
%>
<form name="frmsearch" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
			<td width="15%">대상몰 :</td>
			<td>
				<select name="mallgubun" class="select">
					<option value="">-선택-</option>
				<%
					For i = 0 To oOutMall.FResultCount - 1
						reIdName = oOutMall.FItemList(i).FMallid
						If reIdName = "gseshop" Then
							reIdName = "gsshop"
						End If
				%>
					<option <%= CHKIIF(lcase(vMallGubun) = lcase(reIdName), "selected", "") %> value="<%= reIdName %>"><%= reIdName %></option>
				<% Next %>
				</select>
			</td>
			<td rowspan="4" width="10%"><input type="button" value="검 색" style="width:50px;height:50px;" onClick="javascript:document.frmsearch.submit();"></td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<p />

<form name="frm" action="JaehyuMall_Not_In_Category.asp" methd="post" style="margin:0px;">
<input type="hidden" name="action" value="">
<input type="hidden" name="mallgubun" value="<%=vMallGubun%>">
<input type="hidden" name="arrtargetmall" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
			<td width="25%">
				대상몰 :
			</td>
			<td>
				<input type="radio" name="targetmall" value="<%=vMallGubun%>" onClick="jsShowHideTR()" <%= CHKIIF(targetmall="", "checked", "") %><%= CHKIIF(targetmall=vMallGubun, "checked", "") %>> 현재몰(<%=vMallGubun%>)
				<input type="radio" name="targetmall" value="all" onClick="jsShowHideTR()" <%= CHKIIF(targetmall="all", "checked", "") %>> 모든 Mall
				<input type="radio" name="targetmall" value="sel" onClick="jsShowHideTR()" <%= CHKIIF(targetmall="sel", "checked", "") %>> 일부 Mall
			</td>
		</tr>
		<tr class="showhide">
			<td></td>
			<td>
			<%
				For i = 0 To oOutMall.FResultCount - 1
					reIdName2 = oOutMall.FItemList(i).FMallid
					If reIdName2 = "gseshop" Then
						reIdName2 = "gsshop"
					End If
			%>
				<table width="100%" class="a">
					<tr>
						<td width="30"><input type="checkbox" id="chk<%= i %>" name="selMall" value="<%= reIdName2 %>" <%= CHKIIF(InStr(arrtargetmall, reIdName2) > 0, "checked", "") %>></td>
						<td><%= reIdName2 %></td>
					</tr>
				</table>
				<% Next %>
			</td>
		</tr>

		<tr>
			<td>제외 카테고리 :</td>
			<td width="90%">
				<!-- #include virtual="/common/module/categoryselectbox_cdm.asp"-->
				<input type="button" value="저 장" onClick="insert_id()">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>

<p />

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" class="a">
		<tr>
			<td width="80%">
			</td>
			<td width="20%" align="right">검색결과 : <b><%=vTotalCount%></b></td>
		</tr>
		</table>
	</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="10%">몰구분</td>
	<td width="20%">코드</td>
	<td width="50%">관리카테고리</td>
	<td width="20%">
	<% If vTotalCount <> 0 Then %>
		<input type="button" class="button" value="선택 카테고리 삭제" onClick="delete_idx()">
	<% End If %>
	</td>
</tr>
<%
	IF isArray(arrList) THEN
		For intLoop =0 To UBound(arrList,2)
%>
	<tr bgcolor="#FFFFFF" height="30">
		<td align="center"><%=arrList(1,intLoop)%></td>
		<td align="LEFT"><%=arrList(2,intLoop)%><%=arrList(3,intLoop)%></td>
		<td align="LEFT"><%=arrList(4,intLoop)%> > <%=arrList(5,intLoop)%></td>
		<td align="center"><input type="checkbox" name="del_idx" value="<%=arrList(0,intLoop)%>"></td>
	</tr>
<%
		Next
	Else
%>
	<tr bgcolor="#FFFFFF" height="80">
		<td colspan="4" align="center" class="page_link">[데이터가 없습니다.]</td>
	</tr>
<%
	End If
%>
</table>
</form>
<%
Function Proc()
	Dim strSql, vAction, vMallGubun, arrList, intLoop, j, k
	Dim targetmall, arrtargetmall
	Dim vIdx, vCdl, vCdm
	vAction = Request("action")
	vMallGubun = NullFillWith(Request("mallgubun"),"")
	targetmall = NullFillWith(Request("targetmall"),"")
	arrtargetmall = NullFillWith(Request("arrtargetmall"),"")
	vCdl = Request("cdl")
	vCdm = Request("cdm")

	If vAction = "insert" Then
		If targetmall = "all" Then
			strSql = "select 'interpark' union select 'lotteCom' union select 'lotteimall' union select 'cjmall' union select 'gsshop' union select 'auction1010' union select 'gmarket1010' union select '11st1010' union select 'ssg' union select 'coupang' union select 'hmall1010' union select 'nvstorefarm' union select 'WMP'"
			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				arrList = rsget.getRows()
			END IF
			rsget.close

			IF isArray(arrList) THEN
				For intLoop =0 To UBound(arrList,2)
					vMallGubun = arrList(0,intLoop)
					strSql = ""
					strSql = strSql & " If NOT EXISTS(SELECT * FROM [db_temp].dbo.tbl_JaehyuMall_Not_In_Category Where mallgubun = '" & vMallGubun & "' AND cdl = '" & vCdl & "' AND cdm = '" & vCdm & "') "
					strSql = strSql & "	BEGIN "
					strSql = strSql & "		INSERT INTO [db_temp].dbo.tbl_JaehyuMall_Not_In_Category (cdl, cdm, mallgubun, regdate) VALUES ('" & vCdl & "', '" & vCdm & "', '" & vMallGubun & "', getdate()) "
					strSql = strSql & "	END	"
					dbget.execute strSql
				Next
			End If
			vMallGubun = NullFillWith(Request("mallgubun"),"")
		ElseIf targetmall = "sel" Then
			arrtargetmall = Split(arrtargetmall, ",")

			For intLoop = 0 To UBound(arrtargetmall)
				vMallGubun = Trim(arrtargetmall(intLoop))
				if vMallGubun <> "" then
					strSql = ""
					strSql = strSql & " If NOT EXISTS(SELECT * FROM [db_temp].dbo.tbl_JaehyuMall_Not_In_Category Where mallgubun = '" & vMallGubun & "' AND cdl = '" & vCdl & "' AND cdm = '" & vCdm & "') "
					strSql = strSql & "	BEGIN "
					strSql = strSql & "		INSERT INTO [db_temp].dbo.tbl_JaehyuMall_Not_In_Category (cdl, cdm, mallgubun, regdate) VALUES ('" & vCdl & "', '" & vCdm & "', '" & vMallGubun & "', getdate()) "
					strSql = strSql & "	END	"
					dbget.execute strSql
				end if
			Next
			vMallGubun = NullFillWith(Request("mallgubun"),"")
			arrtargetmall = NullFillWith(Request("arrtargetmall"),"")

			Response.Write "<script>alert('처리되었습니다.');location.href='JaehyuMall_Not_In_Category.asp?mallgubun=" & vMallGubun & "&targetmall=" & targetmall & "&arrtargetmall=" & arrtargetmall & "';</script>"
			Response.End
		Else
			strSql = ""
			strSql = strSql & " If NOT EXISTS(SELECT * FROM [db_temp].dbo.tbl_JaehyuMall_Not_In_Category Where mallgubun = '" & vMallGubun & "' AND cdl = '" & vCdl & "' AND cdm = '" & vCdm & "') "
			strSql = strSql & "	BEGIN "
			strSql = strSql & "		INSERT INTO [db_temp].dbo.tbl_JaehyuMall_Not_In_Category (cdl, cdm, mallgubun, regdate) VALUES ('" & vCdl & "', '" & vCdm & "', '" & vMallGubun & "', getdate()) "
			strSql = strSql & "	END	"
			dbget.execute strSql
		End If
	ElseIf vAction = "delete" Then
		vIdx = Replace(Request("del_idx")," ","")
		vIdx = "'" & Replace(vIdx,",","','") & "'"
		strSql = "DELETE [db_temp].dbo.tbl_JaehyuMall_Not_In_Category WHERE mallgubun = '" & vMallGubun & "' AND idx IN (" & vIdx & ")"
		dbget.execute strSql
	End IF

	Response.Write "<script>alert('처리되었습니다.');location.href='JaehyuMall_Not_In_Category.asp?mallgubun=" & vMallGubun & "';</script>"
	Response.End
End Function
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
