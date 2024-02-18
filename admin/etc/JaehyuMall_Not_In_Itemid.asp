<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->
<!-- #include virtual="/lib/classes/etc/outmallConfirmCls.asp"-->
<%
Dim i, strSql, arrList, vAction, intLoop, vMallGubun, iDelCnt, vCurrPage, vTotalCount, vItemID, vItemName, vMakerID
Dim targetmall, arrtargetmall
Dim targetmall2, arrtargetmall2, tmpMallList, vBigo, vBigoText
Dim tmpItemID, arrTemp, arrItemid
vAction 		= Request("action")
vMallGubun 		= NullFillWith(Request("mallgubun"),"")
vCurrPage		= NullFillWith(Request("cp"),1)
vItemID			= Request("itemid")
vItemName		= Request("itemname")
vMakerID		= Request("makerid")
targetmall		= Request("targetmall")
arrtargetmall	= Request("arrtargetmall")
targetmall2		= Request("targetmall2")
arrtargetmall2	= Request("arrtargetmall2")
If Instr(arrtargetmall2, "gseshop") > 0 Then
	arrtargetmall2 = replace(arrtargetmall2, "gseshop", "gsshop")
End If

vBigo			= Request("bigo")
vBigoText		= Request("bigoText")

If vCurrPage = "" Then vCurrPage = 1
If vItemID<>"" then
	tmpItemID = vItemid
	tmpItemID = replace(tmpItemID,",",chr(10))
	tmpItemID = replace(tmpItemID,chr(13),"")
	arrTemp = Split(tmpItemID,chr(10))
	i = 0
	Do While i <= ubound(arrTemp)
		If Trim(arrTemp(i))<>"" then
			If Not(isNumeric(trim(arrTemp(i)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(i)) & ","
			End If
		End If
		i = i + 1
	Loop
	vItemID = left(arrItemid,len(arrItemid)-1)
end if

if (vMallGubun="lotte") then vMallGubun="lotteCom"   ''' 20130304 추가

If vAction = "insert" OR vAction = "delete" Then
	Call Proc()
End If

Dim oOutMall
SET oOutMall = new cOutmall
	oOutMall.FCurrPage			= 1
	oOutMall.FPageSize			= 100
	'oOutMall.getOutmallList		'스토어팜 제외 쿼리..이하는 포함 쿼리
	oOutMall.getNewOutmallList
%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function insert_id()
{
	var val = GetValueOfRadio('targetmall');
	var chk;
	var i, arrtargetmall;
	if(frm.in_id.value == "")
	{
		alert("ID를 입력하세요.");
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
function delete_id()
{
	frm.action.value = "delete";
	frm.submit();
}

function jsGoPage(iP){
	document.frmsearch.cp.value = iP;
	document.frmsearch.submit();
}

window.onload = function() {
	window.resizeTo(600, 770);

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

function jsSubmit() {
	var i, arrtargetmall2;
	var frm = document.frmsearch;
	var val = GetValueOfRadio('targetmall2');
	var chk;

	arrtargetmall2 = '';
	if (val == 'sel') {
		for (i = 0; ; i++) {
			chk = document.getElementById("vchk2" + i);
			if (chk == undefined) { break; }

			if (chk.checked == true) {
				arrtargetmall2 = arrtargetmall2 + ',' + chk.value;
			}
		}

		if (arrtargetmall2 == '') {
			alert('대상몰을 선택하세요.');
			return;
		}
	}
	frm.arrtargetmall2.value = arrtargetmall2;
	frm.submit();
}
</script>
<center>
Mall 구분 : <b><%=vMallGubun%></b>
</center>
<br>
<form name="frmsearch" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="mallgubun" value="<%=vMallGubun%>">
<input type="hidden" name="arrtargetmall2" value="<%= arrtargetmall2 %>">
<input type="hidden" name="cp" value="<%=vCurrPage%>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
			<td width="15%">대상몰 :</td>
			<td>
				<input type="radio" name="targetmall2" value="<%=vMallGubun%>" onClick="jsShowHideTR2()" <%= CHKIIF(targetmall2="", "checked", "") %><%= CHKIIF(targetmall2=vMallGubun, "checked", "") %>> 현재몰(<%=vMallGubun%>)
				<input type="radio" name="targetmall2" value="sel" onClick="jsShowHideTR2()" <%= CHKIIF(targetmall2="sel", "checked", "") %>> 일부 Mall
				(<input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmsearch.selMall);">전체선택)
			</td>
			<td rowspan="4" width="10%"><input type="button" value="검 색" style="width:50px;height:50px;" onClick="jsSubmit()"></td>
		</tr>
		<tr class="showhide2">
			<td></td>
			<td>
				<% For i = 0 To oOutMall.FResultCount - 1 %>
				<table width="100%" class="a">
					<tr>
						<td width="30"><input type="checkbox" id="vchk2<%= i %>" name="selMall" value="<%= oOutMall.FItemList(i).FMallid %>" <%= CHKIIF(InStr(arrtargetmall2, oOutMall.FItemList(i).FMallid) > 0, "checked", "") %>></td>
						<td><%= oOutMall.FItemList(i).FMallid %></td>
					</tr>
				</table>
				<% Next %>
			</td>
		</tr>
		<tr>
			<td width="15%">브랜드ID :</td>
			<td><input type="text" class="text" name="makerid" value="<%=vMakerID%>" size="20"></td>
		</tr>
		<tr>
			<td>상품ID :</td>
			<td><textarea class="textarea" name="itemid" rows="2" cols="16"><%=replace(vItemid,",",chr(10))%></textarea></td>
		</tr>
		<tr>
			<td>상품명 :</td>
			<td><input type="text" class="text" name="itemname" value="<%=vItemName%>" size="30"></td>
		</tr>
		<tr>
			<td>코맨트여부 :</td>
			<td>
				<Select name="bigo" class="select">
					<option value="">-전체-
					<option value="Y" <%= Chkiif(vBigo="Y", "selected", "") %> >Y
					<option value="N" <%= Chkiif(vBigo="N", "selected", "") %> >N
				</select>
				<input type="text" name="bigoText" value="<%=vBigoText%>">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<p />

<form name="frm" action="JaehyuMall_Not_In_itemid.asp" methd="post" style="margin:0px;">
<input type="hidden" name="action" value="">
<input type="hidden" name="mallgubun" value="<%=vMallGubun%>">
<input type="hidden" name="cp" value="<%=vCurrPage%>">
<input type="hidden" name="arrtargetmall" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
			<td width="15%">
				대상몰 :
			</td>
			<td>
				<input type="radio" name="targetmall" value="<%=vMallGubun%>" onClick="jsShowHideTR()" <%= CHKIIF(targetmall="", "checked", "") %><%= CHKIIF(targetmall=vMallGubun, "checked", "") %>> 현재몰(<%=vMallGubun%>)
				<input type="radio" name="targetmall" value="sel" onClick="jsShowHideTR()" <%= CHKIIF(targetmall="sel", "checked", "") %>> 일부 Mall
				(<input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frm.selMall);">전체선택)
			</td>
		</tr>
		<tr class="showhide">
			<td></td>
			<td>
				<% For i = 0 To oOutMall.FResultCount - 1 %>
				<table width="100%" class="a">
					<tr>
						<td width="30"><input type="checkbox" id="chk<%= i %>" name="selMall" value="<%= oOutMall.FItemList(i).FMallid %>" <%= CHKIIF(InStr(arrtargetmall, oOutMall.FItemList(i).FMallid) > 0, "checked", "") %>></td>
						<td><%= oOutMall.FItemList(i).FMallid %></td>
					</tr>
				</table>
				<% Next %>
			</td>
		</tr>
		<tr>
			<td>제외 상품ID :</td>
			<td width="90%">
				<textarea class="textarea" name="in_id" rows="2" cols="16"></textarea>
			</td>
		</tr>
		<tr>
			<td>코맨트 :</td>
			<td width="90%">
				<input type="text" class="text" name="bigo" size="40">
				<input type="button" class="button" value="저 장" onClick="insert_id()">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<%
	SET oOutMall = nothing

	SET oOutMall = new cOutmall
		oOutMall.FCurrPage			= vCurrPage
		oOutMall.FPageSize			= 100
		oOutMall.FRectMallGubun		= vMallGubun
		oOutMall.FRectTargetmall2	= targetmall2
		oOutMall.FRectTmpMallList	= arrtargetmall2
		oOutMall.FRectmakerid		= vMakerID
		oOutMall.FRectItemid		= vItemID
		oOutMall.FRectItemName		= vItemName
		oOutMall.FRectBigo			= vBigo
		oOutMall.FRectBigoText		= vBigoText
		oOutMall.getJaehyuNotinItemList
%>

<p />

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" class="a">
		<tr>
			<td width="80%">
			</td>
			<td width="20%" align="right">검색결과 : <b><%= FormatNumber(oOutMall.FTotalCount,0) %></b></td>
		</tr>
		</table>
	</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="10%">몰구분</td>
	<td width="10%">상품ID</td>
	<td width="50%">상품명</td>
	<td width="30%">
		<input type="button" value="선택 상품ID 삭제" onClick="delete_id()">
		<input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frm.del_id);">
	</td>
</tr>
<%
	IF oOutmall.FResultCount > 0 THEN
		For i=0 to oOutmall.FResultCount - 1
%>
	<tr bgcolor="#FFFFFF" height="30">
		<td align="center"><%= oOutmall.FItemList(i).FMallgubun %></td>
		<td align="center"><%= oOutmall.FItemList(i).FItemid %></td>
		<td>
		<%
			rw oOutmall.FItemList(i).FItemname
			If oOutmall.FItemList(i).FBigo <> "" Then
				response.write "<font color='blue'>코멘트 : " & oOutmall.FItemList(i).FBigo & "</font>"
			End If
		%>
		</td>
		<td align="center"><input type="checkbox" name="del_id" value="<%= oOutmall.FItemList(i).Fidx %>"></td>
	</tr>
<%
		Next
	Else
%>
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="4" align="center" class="page_link">[데이터가 없습니다.]</td>
	</tr>
<%
	End If
%>
</table>
</form>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oOutmall.HasPreScroll then %>
		<a href="javascript:jsGoPage('<%= oOutmall.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oOutmall.StartScrollPage to oOutmall.FScrollCount + oOutmall.StartScrollPage - 1 %>
    		<% if i>oOutmall.FTotalpage then Exit for %>
    		<% if CStr(vCurrPage)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:jsGoPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oOutmall.HasNextScroll then %>
    		<a href="javascript:jsGoPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<%
Function Proc()
	Dim strSql, vAction, vItemid, vMallGubun, vResult, vCurrPage, vIsAll, arrList, intLoop, j, k
	Dim targetmall, arrtargetmall, bigo
	Dim iA, tmpItemID, arrTemp, arrItemid
	vAction = Request("action")
	vMallGubun = NullFillWith(Request("mallgubun"),"")
	vCurrPage = NullFillWith(Request("cp"),1)
	vIsAll = NullFillWith(Request("isall"),"")
	targetmall = NullFillWith(Request("targetmall"),"")
	arrtargetmall = NullFillWith(Request("arrtargetmall"),"")
	bigo = NullFillWith(Trim(requestCheckVar(request("bigo"),300)),"")
	If bigo <> "" Then
		bigo = bigo & " (" & session("ssBctID") & ")"
	End If
	''response.write targetmall & "<br />"
	''response.write arrtargetmall & "<br />"
	''response.end
	If vAction = "insert" Then
		vItemid = Request("in_id")
		If vItemid<>"" then
			tmpItemID = vItemid
			tmpItemID = replace(tmpItemID,",",chr(10))
			tmpItemID = replace(tmpItemID,chr(13),"")
			arrTemp = Split(tmpItemID,chr(10))
			iA = 0
			Do While iA <= ubound(arrTemp)
				If Trim(arrTemp(iA))<>"" then
					If Not(isNumeric(trim(arrTemp(iA)))) then
						Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
						dbget.close()	:	response.End
					Else
						arrItemid = arrItemid & trim(arrTemp(iA)) & ","
					End If
				End If
				iA = iA + 1
			Loop
			arrItemid = left(arrItemid,len(arrItemid)-1)
		End If

		If targetmall = "all" Then
			strSql = "select 'interpark' union select 'lotteCom' union select 'lotteimall' union select 'cjmall' union select 'gsshop'  union select 'homeplus' union select 'auction1010' union select 'gmarket1010' union select '11st1010' union select 'ssg' union select 'halfclub' union select 'coupang' union select 'hmall1010' union select 'nvstorefarm' union select 'WMP' union select 'LFmall' union select 'lotteon' "

			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				arrList = rsget.getRows()
			END IF
			rsget.close

			IF isArray(arrList) THEN
				arrItemid = Split(arrItemid, ",")
				For intLoop =0 To UBound(arrList,2)
					vMallGubun = arrList(0,intLoop)
					for j = 0 to UBound(arrItemid)
						if Trim(arrItemid(j)) <> "" then
							vItemid = Trim(arrItemid(j))
							strSql = "	DECLARE @Temp CHAR(1) " & _
									 "	If NOT EXISTS(SELECT * FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = '" & vMallGubun & "' AND itemid = '" & vItemid & "') " & _
									 "		BEGIN " & _
							 		 "			INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_itemid(itemid,mallgubun,bigo) VALUES('" & vItemid & "','" & vMallGubun & "', '"& bigo &"') " & _
							 		 "		END	"
							dbget.execute strSql

							strSql = "	DECLARE @Temp CHAR(1) " & _
									 "	If NOT EXISTS(SELECT * FROM [db_outMall].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = '" & vMallGubun & "' AND itemid = '" & vItemid & "') " & _
									 "		BEGIN " & _
									 "			INSERT INTO [db_outMall].dbo.tbl_jaehyumall_not_in_itemid(itemid,mallgubun,bigo) VALUES('" & vItemid & "','" & vMallGubun & "', '"& bigo &"') " & _
									 "		END	"
							dbCTget.execute strSql
						end if
					Next
				Next
			End If
			vMallGubun = NullFillWith(Request("mallgubun"),"")
			vItemid = Request("in_id")
		ElseIf targetmall = "sel" Then
			arrtargetmall = Split(arrtargetmall, ",")
			arrItemid = Split(arrItemid, ",")

			For intLoop = 0 To UBound(arrtargetmall)
				'vMallGubun = Trim(arrtargetmall(intLoop))
				If Trim(arrtargetmall(intLoop)) = "gseshop" Then
					vMallGubun = "gsshop"
				Else
					vMallGubun = Trim(arrtargetmall(intLoop))
				End If

				if vMallGubun <> "" then
					for j = 0 to UBound(arrItemid)
						if Trim(arrItemid(j)) <> "" then
							vItemid = Trim(arrItemid(j))
							strSql = "	DECLARE @Temp CHAR(1) " & _
									 "	If NOT EXISTS(SELECT * FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = '" & vMallGubun & "' AND itemid = '" & vItemid & "') " & _
									 "		BEGIN " & _
							 		 "			INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_itemid(itemid,mallgubun,bigo) VALUES('" & vItemid & "','" & vMallGubun & "', '"& bigo &"') " & _
							 		 "		END	"
							dbget.execute strSql

							strSql = "	DECLARE @Temp CHAR(1) " & _
									 "	If NOT EXISTS(SELECT * FROM [db_outMall].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = '" & vMallGubun & "' AND itemid = '" & vItemid & "') " & _
									 "		BEGIN " & _
									 "			INSERT INTO [db_outMall].dbo.tbl_jaehyumall_not_in_itemid(itemid,mallgubun,bigo) VALUES('" & vItemid & "','" & vMallGubun & "', '"& bigo &"') " & _
									 "		END	"
							dbCTget.execute strSql
						end if
					Next
				end if
			Next
			vMallGubun = NullFillWith(Request("mallgubun"),"")
			vItemid = Request("in_id")
			arrtargetmall = NullFillWith(Request("arrtargetmall"),"")

			Response.Write "<script>alert('처리되었습니다.');location.href='JaehyuMall_Not_In_itemid.asp?mallgubun=" & vMallGubun & "&cp=" & vCurrPage & "&targetmall=" & targetmall & "&arrtargetmall=" & arrtargetmall & "';</script>"
			Response.End
		Else
			arrItemid = Split(arrItemid, ",")
			for j = 0 to UBound(arrItemid)
				if Trim(arrItemid(j)) <> "" then
					vItemid = Trim(arrItemid(j))
					strSql = "	DECLARE @Temp CHAR(1) " & _
							 "	If NOT EXISTS(SELECT * FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = '" & vMallGubun & "' AND itemid = '" & vItemid & "') " & _
							 "		BEGIN " & _
							 "			INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_itemid(itemid,mallgubun,bigo) VALUES('" & vItemid & "','" & vMallGubun & "', '"& bigo &"') " & _
							 "		END	"
					dbget.execute strSql
					''response.write strSql & "<br />"
					strSql = "	DECLARE @Temp CHAR(1) " & _
							 "	If NOT EXISTS(SELECT * FROM [db_outMall].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = '" & vMallGubun & "' AND itemid = '" & vItemid & "') " & _
							 "		BEGIN " & _
							 "			INSERT INTO [db_outMall].dbo.tbl_jaehyumall_not_in_itemid(itemid,mallgubun,bigo) VALUES('" & vItemid & "','" & vMallGubun & "', '"& bigo &"') " & _
							 "		END	"
					dbCTget.execute strSql
				end if
			Next
			vItemid = Request("in_id")
		End If
	ElseIf vAction = "delete" Then
		vItemid = Replace(Request("del_id")," ","")
		vItemid = "'" & Replace(vItemid,",","','") & "'"
		strSql = "DELETE FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE idx IN (" & vItemid & ")"
		dbget.execute strSql

		strSql = "DELETE FROM [db_outMall].dbo.tbl_jaehyumall_not_in_itemid WHERE idx IN (" & vItemid & ")"
		dbCTget.execute strSql
	End IF

	Response.Write "<script>alert('처리되었습니다.');location.href='JaehyuMall_Not_In_itemid.asp?mallgubun=" & vMallGubun & "&cp=" & vCurrPage & "';</script>"
	Response.End
End Function

SET oOutMall = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
