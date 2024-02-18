<%@ language=vbscript %>
<% option explicit %>
<%
Response.CharSet = "euc-kr"
%>
<%
'####################################################
' Description : 비트윈
' History : 2014.10.02 원승현 생성
'			2015.08.12 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/between/betweenItemcls.asp"-->

<%
Dim cDisp, vQuery, vAction, vDepth, vCateCode, vItemID, vSortNo, vIsDefault, vTemp
	vAction		= Request("action")
	vCateCode	= Request("catecode")
	vDepth		= (Len(vCateCode)/3)
	vItemID		= Request("itemid")
	vSortNo		= Request("sortno")
	vIsDefault	= Request("isdefault")
	
If vItemID = "" Then
	dbCTget.close() : Response.End
End IF

If vCateCode = "" Then
	dbCTget.close() : Response.End
End IF

If vAction = "" Then
	vAction = "insert"
End IF

If vSortNo = "" Then
	vSortNo = 9999
End If
	
vQuery = ""
If vAction = "update" OR vAction = "delete" Then
	'vQuery = "SELECT count(catecode) FROM db_outmall.dbo.tbl_between_cate_item WHERE itemid = '" & vItemID & "'"
	
	'//v2 리뉴얼 버전 개편으로 인한 임시방편
	vQuery = "	SELECT count(ci.catecode) as cnt" & vbCrLf
	vQuery = vQuery & "	FROM db_outmall.dbo.tbl_between_cate_item ci" & vbCrLf
	vQuery = vQuery & "	join db_outmall.dbo.tbl_between_cate c" & vbCrLf
	vQuery = vQuery & "		on ci.catecode = c.catecode" & vbCrLf
	vQuery = vQuery & "		AND c.dispyn = 'Y'" & vbCrLf
	vQuery = vQuery & "	WHERE ci.itemid = '" & vItemID & "'" & vbCrLf
		
	'response.write vQuery & "<br>"
	rsCTget.Open vQuery,dbCTget,1
	If not rsCTget.EOF Then
		vTemp = rsCTget(0)
	End If
	rsCTget.close()
End IF

If vAction = "update" Then
	If vTemp = 1 Then
		vIsDefault = "y"	'### 무조건 한개는 기본이어야함. 총갯수가 1개이므로 n으로 변경 불가.
	ElseIf vTemp > 1 Then
		'vQuery = "SELECT catecode FROM db_outmall.dbo.tbl_between_cate_item WHERE itemid = '" & vItemID & "' AND isDefault = 'y'"

		'//v2 리뉴얼 버전 개편으로 인한 임시방편
		vQuery = "	SELECT ci.catecode" & vbCrLf
		vQuery = vQuery & "	FROM db_outmall.dbo.tbl_between_cate_item ci" & vbCrLf
		vQuery = vQuery & "	join db_outmall.dbo.tbl_between_cate c" & vbCrLf
		vQuery = vQuery & "		on ci.catecode = c.catecode" & vbCrLf
		vQuery = vQuery & "		AND ci.isDefault = 'y'" & vbCrLf
		vQuery = vQuery & "		AND c.dispyn = 'Y'" & vbCrLf
		vQuery = vQuery & "	WHERE ci.itemid = '" & vItemID & "'" & vbCrLf
	
		'response.write vQuery & "<br>"
		rsCTget.Open vQuery,dbCTget,1
		If not rsCTget.EOF Then
			vTemp = rsCTget(0)
		end if
		rsCTget.close()

		If CStr(vTemp) = CStr(vCateCode) AND vIsDefault = "n" Then
			'vQuery = "SELECT TOP 1 catecode FROM db_outmall.dbo.tbl_between_cate_item where itemid = '" & vItemID & "' AND isDefault = 'n' ORDER BY depth ASC, sortno ASC"

			'//v2 리뉴얼 버전 개편으로 인한 임시방편
			vQuery = "	SELECT TOP 1  ci.catecode" & vbCrLf
			vQuery = vQuery & "	FROM db_outmall.dbo.tbl_between_cate_item ci" & vbCrLf
			vQuery = vQuery & "	join db_outmall.dbo.tbl_between_cate c" & vbCrLf
			vQuery = vQuery & "		on ci.catecode = c.catecode" & vbCrLf
			vQuery = vQuery & "		AND ci.isDefault = 'n'" & vbCrLf
			vQuery = vQuery & "		AND c.dispyn = 'Y'" & vbCrLf
			vQuery = vQuery & "	WHERE ci.itemid = '" & vItemID & "'" & vbCrLf
			vQuery = vQuery & "	ORDER BY ci.depth ASC, ci.sortno ASC" & vbCrLf

			'response.write vQuery & "<br>"
			rsCTget.Open vQuery,dbCTget,1
			If not rsCTget.EOF Then
				vTemp = rsCTget(0)
			end if
			rsCTget.close()

			'vQuery = "UPDATE db_outmall.dbo.tbl_between_cate_item SET isDefault = 'y' WHERE catecode = '" & vTemp & "' AND itemid = '" & vItemID & "' " & vbCrLf

			'//v2 리뉴얼 버전 개편으로 인한 임시방편
			vQuery = "UPDATE ci SET ci.isDefault = 'y' FROM" & vbCrLf
			vQuery = vQuery & "	db_outmall.dbo.tbl_between_cate_item ci" & vbCrLf
			vQuery = vQuery & "	join db_outmall.dbo.tbl_between_cate c" & vbCrLf
			vQuery = vQuery & "		on ci.catecode = c.catecode" & vbCrLf
			vQuery = vQuery & "		AND c.dispyn = 'Y'" & vbCrLf
			vQuery = vQuery & "	WHERE ci.itemid = '" & vItemID & "'" & vbCrLf
			vQuery = vQuery & "	AND ci.catecode = '" & vTemp & "'" & vbCrLf

			'response.write vQuery & "<br>"
			dbCTget.execute vQuery
		End If
		If CStr(vTemp) <> CStr(vCateCode) AND vIsDefault = "y" Then		'### 이미 y가 있는데 다른카테고리를 y로 지정할경우 일단 같은 itemid 모두 n으로 변경.
			'vQuery = "UPDATE db_outmall.dbo.tbl_between_cate_item SET isDefault = 'n' WHERE itemid = '" & vItemID & "'"

			'//v2 리뉴얼 버전 개편으로 인한 임시방편
			vQuery = "UPDATE ci SET ci.isDefault = 'n' FROM" & vbCrLf
			vQuery = vQuery & "	db_outmall.dbo.tbl_between_cate_item ci" & vbCrLf
			vQuery = vQuery & "	join db_outmall.dbo.tbl_between_cate c" & vbCrLf
			vQuery = vQuery & "		on ci.catecode = c.catecode" & vbCrLf
			vQuery = vQuery & "		AND c.dispyn = 'Y'" & vbCrLf
			vQuery = vQuery & "	WHERE ci.itemid = '" & vItemID & "'" & vbCrLf

			'response.write vQuery & "<br>"
			dbCTget.execute vQuery
		End If
	End If

'	vQuery = "IF EXISTS(SELECT catecode FROM db_outmall.dbo.tbl_between_cate_item WHERE catecode = '" & vCateCode & "' AND itemid = '" & vItemID & "') " & vbCrLf
'	vQuery = vQuery & "	BEGIN " & vbCrLf
'	vQuery = vQuery & "		UPDATE db_outmall.dbo.tbl_between_cate_item SET " & vbCrLf
'	vQuery = vQuery & "			sortNo = '" & vSortNo & "', " & vbCrLf
'	vQuery = vQuery & "			isDefault = '" & vIsDefault & "' " & vbCrLf
'	vQuery = vQuery & "		WHERE catecode = '" & vCateCode & "' AND itemid = '" & vItemID & "' " & vbCrLf
'	vQuery = vQuery & "	END " & vbCrLf

	'//v2 리뉴얼 버전 개편으로 인한 임시방편
	vQuery = "IF EXISTS(" & vbCrLf
	vQuery = vQuery & "		select ci.catecode" & vbCrLf
	vQuery = vQuery & "		FROM db_outmall.dbo.tbl_between_cate_item ci" & vbCrLf
	vQuery = vQuery & "		join db_outmall.dbo.tbl_between_cate c" & vbCrLf
	vQuery = vQuery & "			on ci.catecode = c.catecode" & vbCrLf
	vQuery = vQuery & "			AND c.dispyn = 'Y'" & vbCrLf
	vQuery = vQuery & "		WHERE ci.itemid = '" & vItemID & "'" & vbCrLf
	vQuery = vQuery & "		AND ci.catecode = '" & vCateCode & "'" & vbCrLf
	vQuery = vQuery & "	) " & vbCrLf
	vQuery = vQuery & "	BEGIN " & vbCrLf
	vQuery = vQuery & "		UPDATE ci SET ci.sortNo = '" & vSortNo & "', ci.isDefault = '" & vIsDefault & "' FROM" & vbCrLf
	vQuery = vQuery & "		db_outmall.dbo.tbl_between_cate_item ci" & vbCrLf
	vQuery = vQuery & "		join db_outmall.dbo.tbl_between_cate c" & vbCrLf
	vQuery = vQuery & "			on ci.catecode = c.catecode" & vbCrLf
	vQuery = vQuery & "			AND c.dispyn = 'Y'" & vbCrLf
	vQuery = vQuery & "		WHERE ci.itemid = '" & vItemID & "'" & vbCrLf
	vQuery = vQuery & "		AND ci.catecode = '" & vCateCode & "'" & vbCrLf
	vQuery = vQuery & "	END " & vbCrLf

	'response.write vQuery & "<br>"
	dbCTget.execute vQuery

	response.write "<script type='text/javascript'>parent.location.reload();</script>"

ElseIf vAction = "delete" Then
	'isDefault = 'y' 인것을 지우려할 경우 ORDER BY depth ASC, sortno ASC 로 top 1 catecode를 기본으로 지정.
	If vTemp > 1 Then
		'vQuery = "SELECT catecode FROM db_outmall.dbo.tbl_between_cate_item WHERE itemid = '" & vItemID & "' AND isDefault = 'y'"

		'//v2 리뉴얼 버전 개편으로 인한 임시방편
		vQuery = "SELECT ci.catecode" & vbCrLf
		vQuery = vQuery & "	FROM db_outmall.dbo.tbl_between_cate_item ci" & vbCrLf
		vQuery = vQuery & "	join db_outmall.dbo.tbl_between_cate c" & vbCrLf
		vQuery = vQuery & "		on ci.catecode = c.catecode" & vbCrLf
		vQuery = vQuery & "		AND ci.isDefault = 'y'" & vbCrLf
		vQuery = vQuery & "		AND c.dispyn = 'Y'" & vbCrLf
		vQuery = vQuery & "	WHERE ci.itemid = '" & vItemID & "'" & vbCrLf

		'response.write vQuery & "<br>"
		rsCTget.Open vQuery,dbCTget,1
		If not rsCTget.EOF Then
		vTemp = rsCTget(0)
		end if
		rsCTget.close()
		If CStr(vTemp) = CStr(vCateCode) Then
			'vQuery = "SELECT TOP 1 catecode FROM db_outmall.dbo.tbl_between_cate_item where itemid = '" & vItemID & "' AND isDefault = 'n' ORDER BY depth ASC, sortno ASC"

			'//v2 리뉴얼 버전 개편으로 인한 임시방편
			vQuery = "SELECT ci.catecode" & vbCrLf
			vQuery = vQuery & "	FROM db_outmall.dbo.tbl_between_cate_item ci" & vbCrLf
			vQuery = vQuery & "	join db_outmall.dbo.tbl_between_cate c" & vbCrLf
			vQuery = vQuery & "		on ci.catecode = c.catecode" & vbCrLf
			vQuery = vQuery & "		AND ci.isDefault = 'n'" & vbCrLf
			vQuery = vQuery & "		AND c.dispyn = 'Y'" & vbCrLf
			vQuery = vQuery & "	WHERE ci.itemid = '" & vItemID & "'" & vbCrLf
			vQuery = vQuery & "	ORDER BY ci.depth ASC, ci.sortno ASC" & vbCrLf

			'response.write vQuery & "<br>"
			rsCTget.Open vQuery,dbCTget,1
			If not rsCTget.EOF Then
				vTemp = rsCTget(0)
			end if
			rsCTget.close()

			'vQuery = "UPDATE db_outmall.dbo.tbl_between_cate_item SET isDefault = 'y' WHERE catecode = '" & vTemp & "' AND itemid = '" & vItemID & "' " & vbCrLf

			'//v2 리뉴얼 버전 개편으로 인한 임시방편
			vQuery = "UPDATE ci SET ci.isDefault = 'y' FROM" & vbCrLf
			vQuery = vQuery & "	db_outmall.dbo.tbl_between_cate_item ci" & vbCrLf
			vQuery = vQuery & "	join db_outmall.dbo.tbl_between_cate c" & vbCrLf
			vQuery = vQuery & "		on ci.catecode = c.catecode" & vbCrLf
			vQuery = vQuery & "		AND c.dispyn = 'Y'" & vbCrLf
			vQuery = vQuery & "	WHERE ci.catecode = '" & vTemp & "'" & vbCrLf
			vQuery = vQuery & "	AND ci.itemid = '" & vItemID & "'" & vbCrLf

			'response.write vQuery & "<br>"
			dbCTget.execute vQuery
		End If
	End If
	'#################################################################################################################################################################
	
	'vQuery = "DELETE FROM db_outmall.dbo.tbl_between_cate_item WHERE catecode = '" & vCateCode & "' AND itemid = '" & vItemID & "' " & vbCrLf

	'//v2 리뉴얼 버전 개편으로 인한 임시방편
	vQuery = "DELETE ci FROM" & vbCrLf
	vQuery = vQuery & "	db_outmall.dbo.tbl_between_cate_item ci" & vbCrLf
	vQuery = vQuery & "	join db_outmall.dbo.tbl_between_cate c" & vbCrLf
	vQuery = vQuery & "		on ci.catecode = c.catecode" & vbCrLf
	vQuery = vQuery & "		AND c.dispyn = 'Y'" & vbCrLf
	vQuery = vQuery & "	WHERE ci.catecode = '" & vCateCode & "'" & vbCrLf
	vQuery = vQuery & "	AND ci.itemid = '" & vItemID & "'" & vbCrLf

	'response.write vQuery & "<br>"
	dbCTget.execute vQuery

	response.write "<script type='text/javascript'>parent.location.reload();</script>"

Else
'		vQuery = "IF NOT EXISTS(SELECT catecode FROM db_outmall.dbo.tbl_between_cate_item WHERE catecode = '" & vCateCode & "' AND itemid = '" & vItemID & "') " & vbCrLf
'		vQuery = vQuery & "	BEGIN " & vbCrLf
'		vQuery = vQuery & "		IF NOT EXISTS(SELECT catecode FROM db_outmall.dbo.tbl_between_cate_item WHERE itemid = '" & vItemID & "' AND isDefault = 'y') " & vbCrLf
'		vQuery = vQuery & "		BEGIN " & vbCrLf
'		vQuery = vQuery & "			INSERT INTO db_outmall.dbo.tbl_between_cate_item(catecode, itemid, depth, sortNo, isDefault, regdate) " & vbCrLf
'		vQuery = vQuery & "			VALUES('" & vCateCode & "', '" & vItemID & "', '" & vDepth & "', '" & vSortNo & "', 'y', getdate()) " & vbCrLf
'		vQuery = vQuery & "		END " & vbCrLf
'		vQuery = vQuery & "		ELSE " & vbCrLf
'		vQuery = vQuery & "		BEGIN " & vbCrLf
'		vQuery = vQuery & "			INSERT INTO db_outmall.dbo.tbl_between_cate_item(catecode, itemid, depth, sortNo, isDefault, regdate) " & vbCrLf
'		vQuery = vQuery & "			VALUES('" & vCateCode & "', '" & vItemID & "', '" & vDepth & "', '" & vSortNo & "', 'n', getdate())" & vbCrLf
'		vQuery = vQuery & "		END " & vbCrLf
'		vQuery = vQuery & "	END " & vbCrLf

	'//v2 리뉴얼 버전 개편으로 인한 임시방편
	vQuery = " IF EXISTS(" & vbCrLf
	vQuery = vQuery & "		SELECT ci.catecode " & vbCrLf
	vQuery = vQuery & "		FROM db_outmall.dbo.tbl_between_cate_item ci" & vbCrLf
	vQuery = vQuery & "		join db_outmall.dbo.tbl_between_cate c" & vbCrLf
	vQuery = vQuery & "			on ci.catecode = c.catecode" & vbCrLf
	vQuery = vQuery & "			AND ci.isDefault = 'y'" & vbCrLf
	vQuery = vQuery & "			AND c.dispyn = 'Y'" & vbCrLf
	vQuery = vQuery & "		WHERE ci.itemid = '" & vItemID & "'" & vbCrLf
	vQuery = vQuery & "	)" & vbCrLf
	vQuery = vQuery & "		BEGIN " & vbCrLf
	vQuery = vQuery & "		INSERT INTO db_outmall.dbo.tbl_between_cate_item(catecode, itemid, depth, sortNo, isDefault, regdate) " & vbCrLf
	vQuery = vQuery & "		VALUES('" & vCateCode & "', '" & vItemID & "', '" & vDepth & "', '" & vSortNo & "', 'n', getdate()) " & vbCrLf
	vQuery = vQuery & "		END " & vbCrLf
	vQuery = vQuery & "	ELSE " & vbCrLf
	vQuery = vQuery & "		BEGIN " & vbCrLf
	vQuery = vQuery & "		INSERT INTO db_outmall.dbo.tbl_between_cate_item(catecode, itemid, depth, sortNo, isDefault, regdate) " & vbCrLf
	vQuery = vQuery & "		VALUES('" & vCateCode & "', '" & vItemID & "', '" & vDepth & "', '" & vSortNo & "', 'y', getdate()) " & vbCrLf
	vQuery = vQuery & "		END "

	'response.write vQuery & "<br>"
	dbCTget.execute vQuery
	'response.write "<script type='text/javascript'>parent.location.reload();</script>"
End If


If vAction = "insert" Then
	SET cDisp = New cDispCate
	cDisp.FCurrPage = 1
	cDisp.FPageSize = 1
	cDisp.FRectDepth = vDepth
	cDisp.FRectItemID = vItemID
	cDisp.GetDispCateItemList()
	
	If cDisp.FResultCount > 0 Then
		Response.Write fnCateCodeNameSplit(cDisp.FItemList(0).FCateName, vItemID)
	End If
	
	SET cDisp = Nothing
End If
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->