<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->
<!-- #include virtual="/lib/classes/etc/outmallConfirmCls.asp"-->
<%

dim i
dim oOutMall
SET oOutMall = new cOutmall
	oOutMall.FCurrPage			= 1
	oOutMall.FPageSize			= 100
	oOutMall.getOutmallList

	Dim strSql, arrList, vAction, intLoop, vMallGubun, iDelCnt, vPageSize, vCurrPage, vTotalCount, vItemID, vItemName, vMakerID
	Dim targetmall, arrtargetmall
	Dim targetmall2, arrtargetmall2, tmpMallList
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
	vPageSize 		= "15"

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
		arrItemid = left(arrItemid,len(arrItemid)-1)
	end if



	if (vMallGubun="lotte") then vMallGubun="lotteCom"   ''' 20130304 추가

	If vAction = "insert" OR vAction = "delete" Then
		Call Proc()
	End If

	iDelCnt =  ((vCurrPage - 1) * vPageSize )
	strSql = "SELECT Count(A.idx) FROM [db_temp].[dbo].[tbl_reg_avail_in_itemid] AS A "
	strSql = strSql & "		INNER JOIN [db_item].[dbo].[tbl_item] AS I ON A.itemid = I.itemid "
	strSql = strSql & "Where 1=1 "
	If targetmall2 = "sel"  Then
		tmpMallList = arrtargetmall2
		tmpMallList = Mid(tmpMallList, 2, 1000)
		tmpMallList = Replace(tmpMallList, ",", "','")
		strSql = strSql & " AND A.mallgubun in ('" & tmpMallList & "') "
	ElseIf targetmall2 = "all"  Then
		'//
	else
		strSql = strSql & " AND A.mallgubun = '" & vMallGubun & "' "
	End If
	If vMakerID <> ""  Then
		strSql = strSql & " AND I.makerid = '" & vMakerID & "' "
	End If
	If vItemID <> ""  Then
		strSql = strSql & " AND A.itemid in (" & arrItemid & ") "
	End If
	If vItemName <> ""  Then
		strSql = strSql & " AND I.itemname Like '%" & vItemName & "%' "
	End If
	rsget.Open strSql,dbget
	vTotalCount = rsget(0)
	rsget.close


	strSql = "SELECT Top 15 A.itemid, I.itemname, A.mallgubun FROM [db_temp].[dbo].[tbl_reg_avail_in_itemid] AS A "
	strSql = strSql & "		INNER JOIN [db_item].[dbo].[tbl_item] AS I ON A.itemid = I.itemid "
	strSql = strSql & "Where 1=1 "
	If targetmall2 = "sel"  Then
		strSql = strSql & " AND A.mallgubun in ('" & tmpMallList & "') "
	ElseIf targetmall2 = "all"  Then
		'//
	else
		strSql = strSql & " AND A.mallgubun = '" & vMallGubun & "' "
	End If
	If vMakerID <> ""  Then
		strSql = strSql & " AND I.makerid = '" & vMakerID & "' "
	End If
	If vItemID <> ""  Then
		strSql = strSql & " AND A.itemid in (" & arrItemid & ") "
	End If
	If vItemName <> ""  Then
		strSql = strSql & " AND I.itemname Like '%" & vItemName & "%' "
	End If
	strSql = strSql & "		AND A.idx NOT IN(SELECT TOP "&iDelCnt&" X.idx FROM [db_temp].[dbo].[tbl_reg_avail_in_itemid] AS X "
	strSql = strSql & "							INNER JOIN [db_item].[dbo].[tbl_item] AS Y ON X.itemid = Y.itemid "
	strSql = strSql & "							WHERE 1=1 "
	If targetmall2 = "sel"  Then
		strSql = strSql & " AND X.mallgubun in ('" & tmpMallList & "') "
	ElseIf targetmall2 = "all"  Then
		'//
	else
		strSql = strSql & " AND X.mallgubun = '" & vMallGubun & "' "
	End If
	If vItemID <> ""  Then
		strSql = strSql & " 						AND X.itemid in (" & arrItemid & ") "
	End If
	If vItemName <> ""  Then
		strSql = strSql & " 						AND Y.itemname Like '%" & vItemName & "%' "
	End If
	strSql = strSql & "						ORDER BY X.itemid DESC, X.mallgubun) "
	strSql = strSql & "ORDER BY A.itemid DESC, A.mallgubun"
	rsget.Open strSql,dbget, 1

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
	document.frmpage.cp.value = iP;
	document.frmpage.submit();
}

window.onload = function() {
	window.resizeTo(600, 770);

	// 예외상품 대상몰
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
			chk = document.getElementById("chk2" + i);
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
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
			<td width="15%">대상몰 :</td>
			<td>
				<input type="radio" name="targetmall2" value="<%=vMallGubun%>" onClick="jsShowHideTR2()" <%= CHKIIF(targetmall2="", "checked", "") %><%= CHKIIF(targetmall2=vMallGubun, "checked", "") %>> 현재몰(<%=vMallGubun%>)
				<input type="radio" name="targetmall2" value="all" onClick="jsShowHideTR2()" <%= CHKIIF(targetmall2="all", "checked", "") %> disabled> 모든 Mall
				<input type="radio" name="targetmall2" value="sel" onClick="jsShowHideTR2()" <%= CHKIIF(targetmall2="sel", "checked", "") %>> 일부 Mall
			</td>
			<td rowspan="4" width="10%"><input type="button" value="검 색" style="width:50px;height:50px;" onClick="jsSubmit()"></td>
		</tr>
		<tr class="showhide2">
			<td></td>
			<td>
				<% For i = 0 To oOutMall.FResultCount - 1 %>
				<table width="100%" class="a">
					<tr>
						<td width="30"><input type="checkbox" id="chk2<%= i %>" name="selMall" value="<%= oOutMall.FItemList(i).FMallid %>" <%= CHKIIF(InStr(arrtargetmall2, oOutMall.FItemList(i).FMallid) > 0, "checked", "") %>></td>
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
			<td><textarea class="textarea" name="itemid" rows="2" cols="16"><%= vItemID %></textarea></td>
		</tr>
		<tr>
			<td>상품명 :</td>
			<td><input type="text" class="text" name="itemname" value="<%=vItemName%>" size="30"></td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<p />

<form name="frm" action="reg_avail_itemid.asp" methd="post" style="margin:0px;">
<input type="hidden" name="action" value="">
<input type="hidden" name="mallgubun" value="<%=vMallGubun%>">
<input type="hidden" name="cp" value="<%=vCurrPage%>">
<input type="hidden" name="arrtargetmall" value="">
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
				<input type="radio" name="targetmall" value="all" onClick="jsShowHideTR()" <%= CHKIIF(targetmall="all", "checked", "") %>> 모든 Mall
				<input type="radio" name="targetmall" value="sel" onClick="jsShowHideTR()" <%= CHKIIF(targetmall="sel", "checked", "") %>> 일부 Mall
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
			<td>예외 상품ID :</td>
			<td width="90%">
				<textarea class="textarea" name="in_id" rows="2" cols="16"></textarea>
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
			<td width="20%" align="right">상품수 : <b><%=vTotalCount%></b></td>
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
	<td width="30%"><input type="button" value="선택 상품ID 삭제" onClick="delete_id()"></td>
</tr>
<%
	IF isArray(arrList) THEN
		For intLoop =0 To UBound(arrList,2)
%>
	<tr bgcolor="#FFFFFF" height="30">
		<td align="center"><%=arrList(2,intLoop)%></td>
		<td align="center"><%=arrList(0,intLoop)%></td>
		<td> <%=arrList(1,intLoop)%></td>
		<td align="center"><input type="checkbox" name="del_id" value="<%=arrList(0,intLoop)%>"></td>
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

<form name="frmpage" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="cp" value="<%=vCurrPage%>">
<input type="hidden" name="mallgubun" value="<%=vMallGubun%>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<%
Dim iStartPage, iEndPage, ix, iTotalPage
iStartPage = (Int((vCurrPage-1)/10)*10) + 1
iTotalPage 	=  int((vTotalCount-1)/vPageSize) +1

If (vCurrPage mod vPageSize) = 0 Then
	iEndPage = vCurrPage
Else
	iEndPage = iStartPage + (10-1)
End If
%>
<tr bgcolor="FFFFFF">
	<td height="30" align="center">
		<% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
		<% else %>[pre]<% end if %>
        <%
			for ix = iStartPage  to iEndPage
				if (ix > iTotalPage) then Exit for
				if Cint(ix) = Cint(vCurrPage) then
		%>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="red">[<%=ix%>]</font></a>
		<%		else %>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
		<%
				end if
			next
		%>
    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
		<% else %>[next]<% end if %>
	</td>
</tr>
</table>
</form>
<%
Function Proc()
	Dim strSql, vAction, vItemid, vMallGubun, vResult, vCurrPage, vIsAll, arrList, intLoop, j, k
	Dim targetmall, arrtargetmall
	Dim iA, tmpItemID, arrTemp, arrItemid
	vAction = Request("action")
	vMallGubun = NullFillWith(Request("mallgubun"),"")
	vCurrPage = NullFillWith(Request("cp"),1)
	vIsAll = NullFillWith(Request("isall"),"")
	targetmall = NullFillWith(Request("targetmall"),"")
	arrtargetmall = NullFillWith(Request("arrtargetmall"),"")

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
			strSql = "select 'interpark' union select 'lotteCom' union select 'lotteimall' union select 'cjmall' union select 'gsshop'  union select 'homeplus' union select 'auction1010' union select 'gmarket1010' union select '11st1010' union select 'ssg' union select 'halfclub' union select 'coupang' union select 'hmall1010' union select 'nvstorefarm' union select 'WMP'"

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
									 "	If NOT EXISTS(SELECT * FROM [db_temp].dbo.tbl_reg_avail_in_itemid Where mallgubun = '" & vMallGubun & "' AND itemid = '" & vItemid & "') " & _
									 "		BEGIN " & _
							 		 "			INSERT INTO [db_temp].dbo.tbl_reg_avail_in_itemid(itemid,mallgubun) VALUES('" & vItemid & "','" & vMallGubun & "') " & _
							 		 "		END	"
							dbget.execute strSql
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
				vMallGubun = Trim(arrtargetmall(intLoop))
				if vMallGubun <> "" then
					for j = 0 to UBound(arrItemid)
						if Trim(arrItemid(j)) <> "" then
							vItemid = Trim(arrItemid(j))
							strSql = "	DECLARE @Temp CHAR(1) " & _
									 "	If NOT EXISTS(SELECT * FROM [db_temp].dbo.tbl_reg_avail_in_itemid Where mallgubun = '" & vMallGubun & "' AND itemid = '" & vItemid & "') " & _
									 "		BEGIN " & _
							 		 "			INSERT INTO [db_temp].dbo.tbl_reg_avail_in_itemid(itemid,mallgubun) VALUES('" & vItemid & "','" & vMallGubun & "') " & _
							 		 "		END	"
							dbget.execute strSql
						end if
					Next
				end if
			Next
			vMallGubun = NullFillWith(Request("mallgubun"),"")
			vItemid = Request("in_id")
			arrtargetmall = NullFillWith(Request("arrtargetmall"),"")

			Response.Write "<script>alert('처리되었습니다.');location.href='reg_avail_itemid.asp?mallgubun=" & vMallGubun & "&cp=" & vCurrPage & "&targetmall=" & targetmall & "&arrtargetmall=" & arrtargetmall & "';</script>"
			Response.End
		Else
			arrItemid = Split(arrItemid, ",")
			for j = 0 to UBound(arrItemid)
				if Trim(arrItemid(j)) <> "" then
					vItemid = Trim(arrItemid(j))
					strSql = "	DECLARE @Temp CHAR(1) " & _
							 "	If NOT EXISTS(SELECT * FROM [db_temp].dbo.tbl_reg_avail_in_itemid Where mallgubun = '" & vMallGubun & "' AND itemid = '" & vItemid & "') " & _
							 "		BEGIN " & _
							 "			INSERT INTO [db_temp].dbo.tbl_reg_avail_in_itemid(itemid,mallgubun) VALUES('" & vItemid & "','" & vMallGubun & "') " & _
							 "		END	"
					dbget.execute strSql
				end if
			Next
			vItemid = Request("in_id")
		End If
	ElseIf vAction = "delete" Then
		vItemid = Replace(Request("del_id")," ","")
		vItemid = "'" & Replace(vItemid,",","','") & "'"
		strSql = "DELETE [db_temp].dbo.tbl_reg_avail_in_itemid WHERE mallgubun = '" & vMallGubun & "' AND itemid IN(" & vItemid & ")"
		dbget.execute strSql
	End IF

	Response.Write "<script>alert('처리되었습니다.');location.href='reg_avail_itemid.asp?mallgubun=" & vMallGubun & "&cp=" & vCurrPage & "';</script>"
	Response.End
End Function
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
