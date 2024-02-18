<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : 클리어런스 세일 어드민 등록 처리페이지
'	History		: 2016.01.14 유태욱 생성
'#############################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
dim itemid, i, j, mode, dispcate1, dispcate1arr
dim idxarr, isusingarr, tmpIsusing, tmpdispcate1, cnt
dim sqlstr, arrList, overlapitemid
	mode = requestCheckvar(Request("mode"),15)
	itemid = requestCheckvar(request("itemid"),255)
	menupos = requestCheckvar(request("menupos"),10)
	isusingarr = Request("isusingarr")
	dispcate1arr = Request("dispcate1arr")
	idxarr = Request("idxarr")

Select Case mode
	Case "sortisusingedit"

		'선택이미지 파악
		idxarr = split(idxarr,",")
		cnt = ubound(idxarr)
		isusingarr	=  split(isusingarr,",")
		dispcate1arr	=  split(dispcate1arr,",")

		For i = 0 to cnt
			tmpIsusing = isusingarr(i)
			tmpdispcate1 = dispcate1arr(i)
			
			sqlStr = ""
			sqlStr = sqlStr & " UPDATE db_sitemaster.dbo.tbl_clearance_sale_item SET " & VBCRLF
			sqlStr = sqlStr & " isusing = '"&tmpIsusing&"', dispcate1 = '"&tmpdispcate1&"'" & VBCRLF
			sqlStr = sqlStr & " WHERE idx =" & idxarr(i)
			
			'response.write sqlStr & "<Br>"
			dbget.execute sqlStr
		Next

	Case "iteminsert"
		if itemid<>"" then
			dim iA ,arrTemp, arrItemid
				itemid = replace(itemid,chr(13),"")
				arrTemp = Split(itemid,chr(10))
				iA = 0

				do while iA <= ubound(arrTemp)
					if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
						arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
					end if
					iA = iA + 1
				loop

			if len(arrItemid)>0 then
				itemid = left(arrItemid,len(arrItemid)-1)
			else
				if Not(isNumeric(itemid)) then
					itemid = ""
				end if
			end if
		end if

		arrItemid = Split(itemid,",")	''숫자로된 코드만 배열로 다시 담음

		''이미 등록된 상품 체크
		if ubound(arrItemid) >= 0 then
			for i = 0 to ubound(arrItemid)
				sqlstr= "select itemid  " &_
					" FROM db_sitemaster.dbo.tbl_clearance_sale_item " &_
					" where itemid=" & arrItemid(i)
					'response.write sqlstr
				rsget.Open sqlStr,dbget
				IF Not rsget.EOF THEN
					overlapitemid = rsget(0)
				END IF
				rsget.Close

				if overlapitemid <> "" then
					Response.Write "<script language=javascript>alert('"&overlapitemid&" 상품은 이미 등록된 상품 입니다.');history.back();</script>"
					dbget.close()	:	response.End
				end if
			next
		end if

		if ubound(arrItemid) >= 0 and ubound(arrItemid) < 10 then
			for i = 0 to ubound(arrItemid)

				''상품의 전시카테고리값 가져옴
				sqlstr= "select dispcate1  " &_
					" FROM db_item.dbo.tbl_item " &_
					" where itemid=" & arrItemid(i)
					'response.write sqlstr
				rsget.Open sqlStr,dbget
				IF Not rsget.EOF THEN
					dispcate1 = rsget(0)
				END IF
				rsget.Close

				''클리어런스 DB에 상품코드 저장
				sqlstr = "insert into db_sitemaster.dbo.tbl_clearance_sale_item (itemid, dispcate1)"
				sqlstr = sqlstr & " values ("&arrItemid(i)&", '"&dispcate1&"')"
				'response.write sqlstr
				'response.end
				dbget.execute sqlstr

				''클리어런스 상품 전시카테고리 코드 백업
				sqlstr= "select catecode, itemid  " &_
					" FROM db_item.[dbo].[tbl_display_cate_item] " &_
					" where itemid=" & arrItemid(i)
					'response.write sqlstr
				rsget.Open sqlStr,dbget
				IF Not rsget.EOF THEN
					arrList = rsget.getRows()
				END IF
				rsget.Close

				if isArray(arrList) then
					if ubound(arrList,2) >= 0 then
						for j = 0 to ubound(arrList,2)
							sqlstr = "insert into db_sitemaster.dbo.tbl_clearance_sale_catecode (catecode,itemid)"
							sqlstr = sqlstr & " values ("&arrList(0,j)&","&arrList(1,j)&")"
	'						response.write sqlstr
							dbget.execute sqlstr
						next

						''클리어런스 상품 전시카테고리에서 삭제 (>> 2018-07-05 삭제 안함)
	'					sqlstr = "delete from db_item.[dbo].[tbl_display_cate_item]"
	'					sqlstr = sqlstr & " where itemid="&arrItemid(i)&""
	'					dbget.execute sqlstr

						''클리어런스 상품 itemdb_dispcate1 삭제 (>> 2018-07-05 삭제 안함)
	'					sqlstr = "update db_item.dbo.tbl_item set"
	'					sqlstr = sqlstr & " dispcate1 = NULL"
	'					sqlstr = sqlstr & " where itemid="&arrItemid(i)&""
	'					dbget.execute sqlstr

					end if
				end if
			next
		elseif ubound(arrItemid) > 9 then
			Response.Write "<script language=javascript>alert('한번에 최대 10개까지만 등록할 수 있습니다.');history.back();</script>"
			dbget.close()	:	response.End
		else
			response.write "<script>"
			response.write "	alert('저장할 상품이 없습니다.');"
			response.write "	location.href='/admin/clearancesale/index.asp?menupos="&menupos&"';"
			response.write "</script>"
			dbget.close()	:	response.End
		end if

	Case "itemDelete"
			sqlStr = ""
			sqlStr = sqlStr & " Delete From db_sitemaster.dbo.tbl_clearance_sale_item " & VBCRLF
			sqlStr = sqlStr & " WHERE idx in (" & idxarr & ")"
			'response.write sqlStr & "<Br>"
			dbget.execute sqlStr

	Case Else
		Response.Write "<script language='javascript'>alert('구분자가 없습니다.'); history.back(-1);</script>"
		dbget.close()	:	response.End	
end Select
%>
<script language = "javascript">
<% if mode="sortisusingedit" then %>
	alert("저장되었습니다.\n\n실제 적용까지 3~10분 소요됩니다.");
<% else %>
	alert("저장되었습니다.");
<% end if %>
	location.href="/admin/clearancesale/index.asp?menupos=<%=menupos%>";
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->