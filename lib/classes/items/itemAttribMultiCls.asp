<%
'####################################################
' Description : 상품 속성 일괄수정
' History : 2019/09/05 by eastone
'####################################################

function drawSubshopcateCheckBox(compname,selcatearr)
    Dim strSql, ArrRows, ret, pshopname
    ''subshopdiv	subshopName	catecode	catename	useyn
    ''100	다이어리샵	101102101101	심플	Y

    ArrRows = session("subcatearr")
    if NOT isArray(ArrRows) then
        strSql = " exec [db_item].[dbo].[sp_Ten_category_attrib_subshop_catelist] "
        rsget.CursorLocation = adUseClient
        rsget.Open strSql,dbget,adOpenForwardOnly,adLockReadOnly
        If not rsget.EOF Then
            ArrRows = rsget.getRows()
            session("subcatearr") = ArrRows
        End If
        rsget.Close
    end if
    if isArray(ArrRows) then
        ret = ""
        pshopname = ""
        For i=0 To UBound(ArrRows,2)
            if (pshopname<>ArrRows(1,i)) then ret = ret&" : <strong>"&ArrRows(1,i)&"</strong> : "

            if InStr(selcatearr,ArrRows(0,i)&"-"&ArrRows(2,i))>0 then
                ret = ret&" <input type='checkbox' name='"&compname&"' value='"&ArrRows(0,i)&"-"&ArrRows(2,i)&"' checked>"&ArrRows(3,i)
            else
                ret = ret&" <input type='checkbox' name='"&compname&"' value='"&ArrRows(0,i)&"-"&ArrRows(2,i)&"'>"&ArrRows(3,i)
            end if
    
            pshopname=ArrRows(1,i)
		Next	
    end if 
    ret = ret&"</select>"

    drawSubshopcateCheckBox = ret
End function

Class CAttribGubunItem
    public FattribDiv       ''301
    public FattribDivName   ''다이어리 구분
    public FattribCd        ''301001
    public FattribName      ''다이어리

    public FisChecked       '' 1 /0

    Private Sub Class_Initialize()
        FisChecked = 0
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

			
			

Class CAttribItemItem
    public Fitemid
    public Fitemname
    public Fmakerid
    public Fcatecode
    public Fsmallimage
    
    public FAttrListSize
    public FAttrValList()
    

    
    public Sub SetAttrListSize(iattrSize)
        FAttrListSize = iattrSize
        redim preserve FAttrValList(FAttrListSize)
    End Sub

	Private Sub Class_Initialize()
        redim  FAttrValList(0)
		FAttrListSize  = 0
	End Sub

	Private Sub Class_Terminate()
	End Sub

end Class


'===============================================
'// 상품속성 클래스
'===============================================
Class CAttribMulti
    public FOneItem
    public FAttrGbnList()
    public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
    public FAttrResultCount

    public FRectcateCodeList
	' public FRectAttribCd
	' public FRectAttribDiv
    ' public FRectattribUsing
    ' public FRectDispCate
    ' public FRectItemid
	' public FRectItemName
	' public FRectMakerid
	' public FRectIncludeOption

    '# 상품속성 목록
	public Sub GetAttribMultiItemList()
		dim strSQL,  i, j
		strSQL = "exec db_item.dbo.[sp_Ten_category_attrib_list] "&CHKIIF(FRectcateCodeList="","","'"&FRectcateCodeList&"'")&","&FPageSize&","&FCurrpage
        rsget.CursorLocation = adUseClient
        rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
    
        FAttrResultCount = rsget.RecordCount
        if (FAttrResultCount<1) then FAttrResultCount=0

        redim preserve FAttrGbnList(FAttrResultCount)
        i = 0
        if Not(rsget.EOF or rsget.BOF) then
            Do until rsget.eof
                set FAttrGbnList(i) = new CAttribGubunItem
                FAttrGbnList(i).FattribDiv       = rsget("attribDiv")      ''301
                FAttrGbnList(i).FattribDivName   = rsget("attribDivName")  ''다이어리 구분
                FAttrGbnList(i).FattribCd        = rsget("attribCd")       ''301001
                FAttrGbnList(i).FattribName      = rsget("attribName")     ''다이어리

                i=i+1
                rsget.moveNext
			loop
        end if

        Dim objRs
        SET objRs = rsget.NextRecordset
        if not (objRs is Nothing) then
            if Not(objRs.EOF or objRs.BOF) then
                FTotalCount = objRs(0)
            end if
        end if

        if (FTotalCount<1) then
            FResultCount = 0
            rsget.close()
            SET objRs = Nothing
            Exit Sub
        end if

        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

        SET objRs = rsget.NextRecordset
        if not (objRs is Nothing) then
            FResultCount = objRs.RecordCount
            if (FResultCount<1) then FResultCount=0
            redim preserve FItemList(FResultCount)
            i = 0
            if Not(objRs.EOF or objRs.BOF) then
                Do until objRs.eof
                    set FItemList(i) = new CAttribItemItem
                    FItemList(i).SetAttrListSize(FAttrResultCount)
                    FItemList(i).Fitemid        = objRs("itemid") 
                    FItemList(i).Fitemname      = objRs("itemname") 
                    FItemList(i).Fmakerid       = objRs("makerid") 
                    FItemList(i).Fcatecode      = objRs("catecode") 
                    FItemList(i).Fsmallimage    = objRs("smallimage")

                    FItemList(i).Fsmallimage    = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fsmallimage

                    For J=0 to FAttrResultCount-1
                        set FItemList(i).FAttrValList(j) = New CAttribGubunItem
                        FItemList(i).FAttrValList(j).FattribDiv = FAttrGbnList(j).FattribDiv
                        FItemList(i).FAttrValList(j).FattribDivName = FAttrGbnList(j).FattribDivName
                        FItemList(i).FAttrValList(j).FattribCd = FAttrGbnList(j).FattribCd
                        FItemList(i).FAttrValList(j).FattribName = FAttrGbnList(j).FattribName
                        FItemList(i).FAttrValList(j).FisChecked = objRs(FAttrGbnList(j).FattribCd)
    
                    Next

                    i=i+1
                    objRs.moveNext
                loop
            end if
        end if
        rsget.close
        SET objRs = Nothing

	End Sub

    ' '# 상품속성 정보
	' public Sub GetOneAttrib()
	' 	dim sqlStr

	' 	'내용 접수
    '     sqlStr = "Select top 1 * "
    '     sqlStr = sqlStr & "From db_item.dbo.tbl_itemAttribute "
    '     sqlStr = sqlStr & "Where attribCd='" & attribCd & "'"
	' 	rsget.Open sqlStr, dbget, 1

	' 	FResultCount = rsget.RecordCount

	' 	if Not(rsget.EOF or rsget.BOF) then
	' 		set FOneItem = new CAttribItem

    '         FOneItem.FattribCd			= rsget("attribCd")
    '         FOneItem.FattribDiv			= rsget("attribDiv")
    '         FOneItem.FattribDivName		= rsget("attribDivName")
    '         FOneItem.FattribName		= rsget("attribName")
    '         FOneItem.FattribNameAdd		= rsget("attribNameAdd")
    '         FOneItem.FattribUsing		= rsget("attribUsing")
    '         FOneItem.FattribSortNo		= rsget("attribSortNo")
	' 	end if
	' 	rsget.close
	' End Sub


    Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()
    End Sub

    public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class


' '===============================================
' '// 기타 함수
' '===============================================
' '// 상품속성 선택상자 출력
' function getAttribDivSelectbox(frmNm,selVal,selDisp,addStr)
' 	dim sqlStr, i, strRst

' 	strRst = "<select name='" & frmNm & "' " & addStr & " class='select'>"
' 	strRst = strRst & "<option value="""">::선택::</option>"

' 	sqlStr = "Select attribDiv, attribDivName "
' 	sqlStr = sqlStr & "From db_item.dbo.tbl_itemAttribute "
' 	sqlStr = sqlStr & "Where attribUsing='Y' "

' 	if selDisp<>"" then		'전시
' 	sqlStr = sqlStr & "	and attribDiv in ( "
' 	sqlStr = sqlStr & "		Select distinct attribDiv "
' 	sqlStr = sqlStr & "		from db_item.dbo.tbl_itemAttrib_dispCate "
' 	sqlStr = sqlStr & "		where catecode like '" & selDisp & "%' "
' 	sqlStr = sqlStr & "	) "
' 	end if

' 	sqlStr = sqlStr & "group by attribDiv, attribDivName "
' 	sqlStr = sqlStr & "order by attribDiv"
' 	rsget.Open sqlStr, dbget, 1

' 	if Not(rsget.EOF or rsget.BOF) then
' 		Do Until rsget.EOF
' 			strRst = strRst & "<option value=""" & rsget("attribDiv") & """" & chkIIF(cStr(rsget("attribDiv"))=cStr(selVal),"selected","") & ">" & rsget("attribDivName") & "</option>"
' 			rsget.MoveNext
' 		Loop
' 	end if

' 	rsget.Close

' 	strRst = strRst & "</select>"

' 	getAttribDivSelectbox = strRst
' end function

' '// 전시카테고리 선택상자 출력 (1Depth)
' function getDispCateSelectbox(frmNm,selVal,addStr)
' 	dim sqlStr, i, strRst

' 	strRst = "<select name='" & frmNm & "' " & addStr & " class='select'>"
' 	strRst = strRst & "<option value="""">::선택::</option>"

' 	sqlStr = "select catecode, catename "
' 	sqlStr = sqlStr & "from db_item.dbo.tbl_display_cate "
' 	sqlStr = sqlStr & "where depth='1' "
' 	sqlStr = sqlStr & "	and useyn='Y' "
' 	sqlStr = sqlStr & "order by sortNo, catecode "
' 	rsget.Open sqlStr, dbget, 1

' 	if Not(rsget.EOF or rsget.BOF) then
' 		Do Until rsget.EOF
' 			strRst = strRst & "<option value=""" & rsget("catecode") & """" & chkIIF(cStr(rsget("catecode"))=cStr(selVal),"selected","") & ">" & rsget("catename") & "</option>"
' 			rsget.MoveNext
' 		Loop
' 	end if

' 	rsget.Close

' 	strRst = strRst & "</select>"

' 	getDispCateSelectbox = strRst
' end function

' '// 카테고리 Histoty 출력
' function getDispCateHistory(code)
' 	dim strHistory, strLink, SQL, i, j
' 	j = (len(code)/3)
    
' 	'히스토리 기본
' 	strHistory = ""

' 	'// 카테고리 이름 접수
' 	SQL = "SELECT ([db_item].[dbo].[getCateCodeFullDepthName]('" & code & "'))"
' 	rsget.Open SQL, dbget, 1

' 	if NOT(rsget.EOF or rsget.BOF) then
' 		for i = 1 to j
' 			if i>1 then strHistory = strHistory & "&nbsp;&gt;&nbsp;"
' 			if i = j then
' 				strHistory = strHistory & "<strong>" & Split(db2html(rsget(0)),"^^")(i-1) & "</strong>"
' 			else
' 				strHistory = strHistory & Split(db2html(rsget(0)),"^^")(i-1)
' 			end if
' 		next
' 	end if
	
' 	rsget.Close

' 	getDispCateHistory=strHistory
' end Function
%>