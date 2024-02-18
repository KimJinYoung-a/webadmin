<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim cmdparam, arrItemid, vItemid, arrkeywords, i, vKeywords
Dim strSql, strSql2, subject, vQuery, strSql3, strSql4, midx
Dim mode, etc, prekeyword, nextkeyword, vRow
Dim tSQL, tArrRows, t, spk, k, tItemid, tItemid2, tmpItemid, tmpItemid2
cmdparam	= request("cmdparam")
arrItemid	= request("cksel")
arrkeywords	= request("arrkeywords")
mode		= request("mode")
etc			= request("etc")
prekeyword	= db2html(request("prekeyword"))
nextkeyword	= db2html(request("nextkeyword"))
If cmdparam = "chk" Then		'적용
	If Right(arrkeywords,4) = "*(^!" Then
		arrkeywords = Left(arrkeywords, Len(arrkeywords) - 4)
	End If
	vItemid		= split(arrItemid, ",")
	vKeywords	= split(arrkeywords, "*(^!")

	vQuery = " SELECT TOP 1 itemname FROM db_item.dbo.tbl_item WHERE itemid = '"&Trim(vItemid(0))&"' "
	rsget.Open vQuery,dbget,1
	If Not rsget.Eof Then
		subject = rsget("itemname")
	End If
	rsget.Close

	If UBound(vItemid) > 0 Then
		subject = subject & " 등 " & UBound(vItemid)+1 &"건"
	Else
		subject = subject & " " & UBound(vItemid)+1 &"건"
	End If

	strSql = ""
	strSql2 = ""
	For i = LBound(vItemid) to UBound(vItemid)
		strSql = strSql & " UPDATE db_item.dbo.tbl_item_contents SET  keywords = '"&Trim(vKeywords(i))&"'  WHERE itemid = '"&Trim(vItemid(i))&"';  " & vbcrlf
		strSql2 = strSql2 & " INSERT INTO watcher.[dbo].[kc_job_log] "
		'strSql2 = strSql2 & " SELECT "&Trim(vItemid(i))&",NULL,NULL,NULL,NULL,'U',getdate(),NULL,'P','db_item','dbo','TBL_ITEM','db_item','dbo','vw_item_DispCate', NULL; " & vbcrlf
		strSql2 = strSql2 & " SELECT 'itemDisp', "&Trim(vItemid(i))&", 'U', getdate(); " & vbcrlf
	Next
	dbget.execute strSql
	dbget.execute strSql2

	strSql3 = ""
	strSql3 = strSql3 & " INSERT INTO db_item.dbo.tbl_keyword_master (mode, subject, regid, regdate) VALUES ('U', '"&subject&"', '"&session("ssBctID")&"', getdate()) "
	dbget.execute strSql3
	
	vQuery = ""
	vQuery = vQuery & " SELECT TOP 1 idx FROM db_item.dbo.tbl_keyword_master ORDER BY idx DESC "
	rsget.Open vQuery,dbget,1
	If Not rsget.Eof Then
		midx = rsget("idx")
	End If
	rsget.Close

	strSql4 = ""
	strSql4 = strSql4 & " INSERT INTO db_item.dbo.tbl_keyword_detail "
	strSql4 = strSql4 & " SELECT '"&midx&"', itemid, keywords "
	strSql4 = strSql4 & " FROM db_item.dbo.tbl_item_contents "
	strSql4 = strSql4 & " WHERE itemid in ( "&Trim(arrItemid)&" ) "
	dbget.execute strSql4
	response.write "<script>alert('변경이 완료되었습니다');parent.location.reload();</script>"
ElseIf cmdparam = "allchk" Then		'일괄변경
	vItemid		= split(arrItemid, ",")
	vQuery = " SELECT TOP 1 itemname FROM db_item.dbo.tbl_item WHERE itemid = '"&Trim(vItemid(0))&"' "
	rsget.Open vQuery,dbget,1
	If Not rsget.Eof Then
		subject = rsget("itemname")
	End If
	rsget.Close

	If UBound(vItemid) > 0 Then
		subject = subject & " 등 " & UBound(vItemid)+1 &"건"
	Else
		subject = subject & " " & UBound(vItemid)+1 &"건"
	End If

	tSQL = ""
	tSQL = tSQL & " SELECT itemid, keywords FROM db_item.dbo.tbl_item_contents WHERE itemid in ("&Trim(arrItemid)&") ORDER BY itemid DESC "
	rsget.Open tSQL,dbget,1
	If Not rsget.Eof Then
		tArrRows = rsget.getRows()
	End If
	rsget.Close

	If mode = "I" Then
		For t = 0 to Ubound(tArrRows, 2)
			spk = Split(tArrRows(1,t), ",")
			For k = 0 to Ubound(spk)
				If Trim(spk(k)) = Trim(nextkeyword) Then
					tmpItemid = tmpItemid & tArrRows(0,t) & ","
				End If
			Next
		Next

		If Right(tmpItemid,1) = "," Then
			tmpItemid = Left(tmpItemid, Len(tmpItemid) - 1)
		End If

		If tmpItemid <> "" Then
			tItemid = Split(tmpItemid, ",")

			tSQL = ""
			tSQL = tSQL & " SELECT itemid FROM db_item.dbo.tbl_item WHERE itemid NOT in ("&Trim(tmpItemid)&") AND itemid in ("&Trim(arrItemid)&") ORDER BY itemid DESC "
			rsget.Open tSQL,dbget,1
			If Not rsget.Eof Then
				tArrRows = rsget.getRows()
			End If
			rsget.Close

			For t = 0 to Ubound(tArrRows, 2)
				tmpItemid2 = tmpItemid2 & tArrRows(0,t) & ","
			Next
			If Right(tmpItemid2,1) = "," Then
				tmpItemid2 = Left(tmpItemid2, Len(tmpItemid2) - 1)
			End If
			tItemid2 = Split(tmpItemid2, ",")
		Else
			tItemid2 = vItemid
		End If

		strSql = ""
		strSql2 = ""
		For i = LBound(tItemid2) to UBound(tItemid2)
			strSql = strSql & " UPDATE db_item.dbo.tbl_item_contents SET  keywords = keywords + ',"& Trim(nextkeyword)&"'  WHERE itemid = '"&Trim(tItemid2(i))&"';  " & vbcrlf
			strSql2 = strSql2 & " INSERT INTO watcher.[dbo].[kc_job_log] "
'			strSql2 = strSql2 & " SELECT "&Trim(tItemid2(i))&",NULL,NULL,NULL,NULL,'U',getdate(),NULL,'P','db_item','dbo','TBL_ITEM','db_item','dbo','vw_item_DispCate', NULL; " & vbcrlf
			strSql2 = strSql2 & " SELECT 'itemDisp', "&Trim(tItemid2(i))&", 'U', getdate();  " & vbcrlf
		Next
		dbget.execute strSql
		dbget.execute strSql2

		strSql3 = ""
		strSql3 = strSql3 & " INSERT INTO db_item.dbo.tbl_keyword_master (mode, subject, nextkeyword, etc, regid, regdate)  VALUES ('"&mode&"', '"&subject&"', '"&Trim(nextkeyword)&"', '"&etc&"', '"&session("ssBctID")&"', getdate()) "
		dbget.execute strSql3

		vQuery = ""
		vQuery = vQuery & " SELECT TOP 1 idx FROM db_item.dbo.tbl_keyword_master ORDER BY idx DESC "
		rsget.Open vQuery,dbget,1
		If Not rsget.Eof Then
			midx = rsget("idx")
		End If
		rsget.Close

		strSql4 = ""
		strSql4 = ""
		strSql4 = strSql4 & " INSERT INTO db_item.dbo.tbl_keyword_detail "
		strSql4 = strSql4 & " SELECT '"&midx&"', itemid, keywords "
		strSql4 = strSql4 & " FROM db_item.dbo.tbl_item_contents "
		strSql4 = strSql4 & " WHERE itemid in ( "&Trim(arrItemid)&" ) "
		dbget.execute strSql4
		response.write "<script>alert('변경이 완료되었습니다');opener.location.reload();self.close();</script>"
	ElseIf mode = "U" Then
		strSql = ""
		strSql2 = ""
		For i = LBound(vItemid) to UBound(vItemid)
			strSql = strSql & " UPDATE db_item.dbo.tbl_item_contents SET  keywords = replace(keywords, ',"& Trim(prekeyword)&"', ',"& Trim(nextkeyword)&"')  WHERE itemid = '"&Trim(vItemid(i))&"' and charindex(',"& Trim(prekeyword)&",', keywords)> 0;  " & vbcrlf
			strSql = strSql & " UPDATE db_item.dbo.tbl_item_contents SET  keywords = replace(keywords, '"& Trim(prekeyword)&",', '"& Trim(nextkeyword)&",')  WHERE itemid = '"&Trim(vItemid(i))&"' and charindex('"& Trim(prekeyword)&",', keywords)> 0;  " & vbcrlf
			strSql2 = strSql2 & " INSERT INTO watcher.[dbo].[kc_job_log] "
			strSql2 = strSql2 & " SELECT 'itemDisp', "&Trim(vItemid(i))&", 'U', getdate(); " & vbcrlf
		Next
		dbget.execute strSql
		dbget.execute strSql2

		strSql3 = ""
		strSql3 = strSql3 & " INSERT INTO db_item.dbo.tbl_keyword_master (mode, subject, prekeyword, nextkeyword, etc, regid, regdate)  VALUES ('"&mode&"', '"&subject&"', '"&Trim(prekeyword)&"', '"&Trim(nextkeyword)&"', '"&etc&"', '"&session("ssBctID")&"', getdate()) "
		dbget.execute strSql3

		vQuery = ""
		vQuery = vQuery & " SELECT TOP 1 idx FROM db_item.dbo.tbl_keyword_master ORDER BY idx DESC "
		rsget.Open vQuery,dbget,1
		If Not rsget.Eof Then
			midx = rsget("idx")
		End If
		rsget.Close

		strSql4 = ""
		strSql4 = ""
		strSql4 = strSql4 & " INSERT INTO db_item.dbo.tbl_keyword_detail "
		strSql4 = strSql4 & " SELECT '"&midx&"', itemid, keywords "
		strSql4 = strSql4 & " FROM db_item.dbo.tbl_item_contents "
		strSql4 = strSql4 & " WHERE itemid in ( "&Trim(arrItemid)&" ) "
		dbget.execute strSql4
		response.write "<script>alert('변경이 완료되었습니다');opener.location.reload();self.close();</script>"
	ElseIf mode = "D" Then
		For t = 0 to Ubound(tArrRows, 2)
			spk = Split(tArrRows(1,t), ",")
			For k = 0 to Ubound(spk)
				If Trim(spk(k)) = Trim(nextkeyword) Then
					tmpItemid = tmpItemid & tArrRows(0,t) & ","
				End If
			Next
		Next

		If Right(tmpItemid,1) = "," Then
			tmpItemid = Left(tmpItemid, Len(tmpItemid) - 1)
		End If

		If tmpItemid <> "" Then
			tItemid = Split(tmpItemid, ",")
		Else
			tItemid = vItemid
		End If

		strSql = ""
		strSql2 = ""
		For i = LBound(tItemid) to UBound(tItemid)
			strSql = strSql & " UPDATE db_item.dbo.tbl_item_contents SET  keywords = replace(keywords, '"& Trim(nextkeyword)&",', '')  WHERE itemid = '"&Trim(tItemid(i))&"';  " & vbcrlf
			strSql = strSql & " UPDATE db_item.dbo.tbl_item_contents SET  keywords = replace(keywords, ',"& Trim(nextkeyword)&"', '')  WHERE itemid = '"&Trim(tItemid(i))&"';  " & vbcrlf
			strSql2 = strSql2 & " INSERT INTO watcher.[dbo].[kc_job_log] "
			strSql2 = strSql2 & " SELECT 'itemDisp', "&Trim(tItemid(i))&", 'U', getdate(); " & vbcrlf
		Next
		dbget.execute strSql
		dbget.execute strSql2

		strSql3 = ""
		strSql3 = strSql3 & " INSERT INTO db_item.dbo.tbl_keyword_master (mode, subject, nextkeyword, etc, regid, regdate)  VALUES ('"&mode&"', '"&subject&"', '"&Trim(nextkeyword)&"', '"&etc&"', '"&session("ssBctID")&"', getdate()) "
		dbget.execute strSql3

		vQuery = ""
		vQuery = vQuery & " SELECT TOP 1 idx FROM db_item.dbo.tbl_keyword_master ORDER BY idx DESC "
		rsget.Open vQuery,dbget,1
		If Not rsget.Eof Then
			midx = rsget("idx")
		End If
		rsget.Close

		strSql4 = ""
		strSql4 = ""
		strSql4 = strSql4 & " INSERT INTO db_item.dbo.tbl_keyword_detail "
		strSql4 = strSql4 & " SELECT '"&midx&"', itemid, keywords "
		strSql4 = strSql4 & " FROM db_item.dbo.tbl_item_contents "
		strSql4 = strSql4 & " WHERE itemid in ( "&Trim(arrItemid)&" ) "
		dbget.execute strSql4
		response.write "<script>alert('변경이 완료되었습니다');opener.location.reload();self.close();</script>"
	End If
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
