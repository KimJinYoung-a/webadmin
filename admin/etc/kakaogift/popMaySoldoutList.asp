<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim strSQL
Dim itemidarr, kakaoidarr

strSQL = " exec [db_etcmall].[dbo].[usp_Ten_OutMall_Kakaogift_MaysoldoutList] "
rsget.CursorLocation = adUseClient
rsget.CursorType=adOpenStatic
rsget.Locktype=adLockReadOnly
rsget.Open strSQL, dbget
If Not(rsget.EOF or rsget.BOF) Then
	Do Until rsget.EOF
		itemidarr = itemidarr & rsget("itemid") & ","
		kakaoidarr = kakaoidarr & rsget("kakaoGiftGoodNo") & ","
		rsget.MoveNext
	Loop
End If
rsget.Close

If Right(itemidarr,1) = "," Then
	itemidarr = Left(itemidarr, Len(itemidarr) - 1)
End If

If Right(kakaoidarr,1) = "," Then
	kakaoidarr = Left(kakaoidarr, Len(kakaoidarr) - 1)
End If
%>
<table width="100%" align="center">
<tr align="center">
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>텐바이텐 상품코드</td>
		</tr>
		<tr align="center" bgcolor="#FFFFFF">
			<td>
				<textarea cols="80" rows="5"><%=itemidarr%></textarea>
			</td>
		</tr>
		</table>
	</td>

	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>카카오기프트 상품코드</td>
		</tr>
		<tr align="center" bgcolor="#FFFFFF">
			<td>
				<textarea cols="80" rows="5"><%=kakaoidarr%></textarea>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>


<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
