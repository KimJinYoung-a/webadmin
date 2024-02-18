<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/giftcard/giftcard_cls.asp"-->
<%
Dim userid, lp, strSql, okCnt, errStr, arrRows, i
okCnt = 0
userid = request("userid")
If userid <> "" then
	Dim useridCnt, iA2, arrTemp2, arruserid
	userid = replace(userid,",",chr(10))
	userid = replace(userid,chr(13),"")
	arrTemp2 = Split(userid,chr(10))
	iA2 = 0
	useridCnt = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arruserid = arruserid & trim(arrTemp2(iA2)) & ","
			useridCnt = useridCnt + 1
		End If
		iA2 = iA2 + 1
	Loop
	arruserid = left(arruserid,len(arruserid)-1)
End If
	strSql = ""
	strSql = strSql & " SELECT userid, username, usercell "
	strSql = strSql & " INTO #tmpUserTBL "
	strSql = strSql & " FROM [db_user].[dbo].tbl_user_n "
	strSql = strSql & " WHERE 1=2 "
	dbget.execute strSql

	For lp = 0 to useridCnt - 1
		strSql = ""
		strSql = strSql & " INSERT INTO #tmpUserTBL (userid, username, usercell) VALUES " & vbcrlf
		strSql = strSql & " ('"&Split(arruserid, ",")(lp)&"', '', '') "
		dbget.execute strSql
	Next

	strSql = ""
	strSql = strSql & " SELECT t.userid, isnull(n.username, '') as username, isnull(n.usercell, '') as usercell " & vbcrlf
	strSql = strSql & " FROM #tmpUserTBL as t "
	strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_n as n on t.userid = n.userid "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	IF Not (rsget.EOF OR rsget.BOF) THEN
		arrRows = rsget.getRows()
	END IF
	rsget.close
%>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td with="10%">번호</td>
    <td with="20%">ID</td>
	<td with="30%">이름</td>
	<td with="40%">휴대폰번호</td>
</tr>
<%
If isArray(arrRows) Then
	For i = 0 To Ubound(arrRows, 2)
%>
<tr align="center" height="25" bgcolor="#FFFFFF">
	<td><%= i+1 %></td>
    <td><%= arrRows(0, i) %></td>
	<td>
	<%
		If arrRows(1, i) <> "" Then
			response.write unescape(AstarUserName(arrRows(1, i)))
		Else
			response.write "<font color='RED'>이름 없음</font>"
		End If
	%>
	</td>
	<td>
	<%
		If arrRows(2, i) <> "" Then
			response.write AstarPhoneNumber(arrRows(2, i))
		Else
			response.write "<font color='RED'>폰 번호 없음</font>"
		End If
	%>
	</td>
</tr>
<%
	Next
End If
%>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->
