<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	dim strSql, mode, sACd, sSortNo, sIsUsing, i
	Dim attribCd,attribDiv,attribDivName,attribName,attribNameAdd,attribUsing,attribSortNo
	dim referer, strRtn, newDiv
	DIM strSql2
	DIM mobile_image1, mobile_image2, mobile_image3, mobile_image4, mobile_image5, mobile_image6
	DIM pc_image1, pc_image2, pc_image3, pc_image4, pc_image5, pc_image6
	referer = request.ServerVariables("HTTP_REFERER")

	mode = request.form("mode")
	strSql = ""

	attribCd		= request.form("attribCd")
	attribDiv		= request.form("attribDiv")
	attribDivName	= request.form("attribDivName")
	attribName		= request.form("attribName")
	attribNameAdd	= request.form("attribNameAdd")
	attribUsing		= request.form("attribUsing")
	attribSortNo	= request.form("attribSortNo")
	newDiv			= request.form("newDiv")
	mobile_image1	= request.form("mobile_image1")
	mobile_image2	= request.form("mobile_image2")
	mobile_image3	= request.form("mobile_image3")
	mobile_image4	= request.form("mobile_image4")
	mobile_image5	= request.form("mobile_image5")
	mobile_image6	= request.form("mobile_image6")
	pc_image1	= request.form("pc_image1")
    pc_image2	= request.form("pc_image2")
    pc_image3	= request.form("pc_image3")
    pc_image4	= request.form("pc_image4")
    pc_image5	= request.form("pc_image5")
    pc_image6	= request.form("pc_image6")

	'// 처리 모드 분기
	Select Case mode
		Case "attrArr"
			'상품속성 상태 일괄 저장

			for i=1 to request.form("chkCd").count
				sACd = request.form("chkCd")(i)
				if sACd<>"" then
					sSortNo = request.form("sort"&sACd)
					sIsUsing = request.form("use"&sACd)
					if sSortNo="" then sSortNo="0"
					if sIsUsing="" then sIsUsing="N"

					strSql = strSql & "Update db_item.dbo.tbl_itemAttribute Set "
					strSql = strSql & " attribSortNo='" & sSortNo & "'"
					strSql = strSql & " ,attribUsing='" & sIsUsing & "'"		'사이트 메인: 사용여부 > 선노출로 변경
					strSql = strSql & " Where attribCd='" & sACd & "';" & vbCrLf
				end if
			next

			strRtn = "location.replace('" + referer + "');"

		Case "attrNew"
			'상품속성 신규등록

			if newDiv="Y" then
				'구분코드 생성
				strSql = "Select Max(attribDiv) as maxDiv From db_item.dbo.tbl_itemAttribute"
				rsget.Open strSql, dbget, 1
				if isNull(rsget("maxDiv")) then
					attribDiv = "100"
				else
					attribDiv = Num2Str(rsget("maxDiv")+1,3,"0","R")
				end if
				rsget.Close
			end if

			'속성코드 생성
			strSql = "Select Max(attribCd) as maxCd From db_item.dbo.tbl_itemAttribute Where attribDiv='" & attribDiv & "'"
			rsget.Open strSql, dbget, 1
			if isNull(rsget("maxCd")) then
				attribCd = attribDiv & "101"
			else
				attribCd = attribDiv & Num2Str(cInt(right(rsget("maxCd"),3))+1,3,"0","R")
			end if
			rsget.Close

			if Not(attribName="") then
				strSql = "Insert into db_item.dbo.tbl_itemAttribute (attribCd,attribDiv,attribDivName,attribName,attribNameAdd,attribUsing,attribSortNo) values "
				strSql = strSql & "('" & attribCd & "'"
				strSql = strSql & ",'" & attribDiv & "'"
				strSql = strSql & ",'" & attribDivName & "'"
				strSql = strSql & ",'" & attribName & "'"
				strSql = strSql & ",'" & attribNameAdd & "'"
				strSql = strSql & ",'" & attribUsing & "'"
				strSql = strSql & ",'" & attribSortNo & "')"
			end if

			strRtn = "opener.history.go(0);location.replace('" + referer + "');"

		Case "attrModi"
			'상품속성 수정
			if Not(attribCd="" or attribName="") then
				strSql = "Update db_item.dbo.tbl_itemAttribute"
				strSql = strSql & " Set attribName='" & attribName & "'"
				strSql = strSql & " ,attribNameAdd='" & attribNameAdd & "'"
				strSql = strSql & " ,attribUsing='" & attribUsing & "'"
				strSql = strSql & " ,attribSortNo='" & attribSortNo & "'"
				strSql = strSql & " Where attribCd='" & attribCd & "'"

				strSql2 = "MERGE INTO db_item.dbo.tbl_itemAttribute_detail AS id"
                strSql2 = strSql2 & " USING (SELECT attribCd = '" & attribCd & "') as sub1"
                strSql2 = strSql2 & " ON (sub1.attribCd = id.attribCd)"
                strSql2 = strSql2 & " WHEN MATCHED THEN"
                strSql2 = strSql2 & " UPDATE SET mobile_image1 = '" & mobile_image1 & "'"
                strSql2 = strSql2 & " , mobile_image2 = '" & mobile_image2 & "'"
                strSql2 = strSql2 & " , mobile_image3 = '" & mobile_image3 & "'"
                strSql2 = strSql2 & " , mobile_image4 = '" & mobile_image4 & "'"
                strSql2 = strSql2 & " , mobile_image5 = '" & mobile_image5 & "'"
                strSql2 = strSql2 & " , mobile_image6 = '" & mobile_image6 & "'"
                strSql2 = strSql2 & " , pc_image1 = '" & pc_image1 & "'"
                strSql2 = strSql2 & " , pc_image2 = '" & pc_image2 & "'"
                strSql2 = strSql2 & " , pc_image3 = '" & pc_image3 & "'"
                strSql2 = strSql2 & " , pc_image4 = '" & pc_image4 & "'"
                strSql2 = strSql2 & " , pc_image5 = '" & pc_image5 & "'"
                strSql2 = strSql2 & " , pc_image6 = '" & pc_image6 & "'"
                strSql2 = strSql2 & " WHEN NOT MATCHED THEN"
                strSql2 = strSql2 & " INSERT(attribCd, mobile_image1, mobile_image2, mobile_image3, mobile_image4, mobile_image5, mobile_image6, pc_image1, pc_image2, pc_image3, pc_image4, pc_image5, pc_image6) VALUES"
                strSql2 = strSql2 & " ('"& attribCd &"', '" &mobile_image1& "', '" &mobile_image2& "', '" &mobile_image3& "', '" &mobile_image4& "', '" &mobile_image5& "', '" &mobile_image6& "', '" &pc_image1& "', '" &pc_image2& "', '" &pc_image3& "', '" &pc_image4& "', '" &pc_image5& "', '" &pc_image6& "');"
			end if

			strRtn = "opener.history.go(0);location.replace('" + referer + "');"

	end Select

	if strSql<>"" then
		dbget.Execute strSql

		IF strSql2 <> "" THEN
		    dbget.Execute strSql2
        END IF
	else
		Call Alert_return("저장할 내용이 없습니다.")
		dbget.Close: Response.End
	end if

	response.write "<script>alert('저장되었습니다.');</script>"
	response.write "<script>" & strRtn & "</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->