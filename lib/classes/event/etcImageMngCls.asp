<%
'####################################################
' Description :  이미지 관리 클래스
' History : 2016.07.28 서동석 생성
'			2016.08.12 한용민 수정
'####################################################
%>
<%
function getImgEtcFolderTitleByFolderIdx(folderidx)
    Dim sqlStr
    sqlStr = " SELECT folderidx,foldertitle"&vbCRLF
    sqlStr = sqlStr & "  from db_event.[dbo].[tbl_etcImage_master] "
    sqlStr = sqlStr & " where folderidx="&folderidx
    
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    IF not rsget.eof THEN
        getImgEtcFolderTitleByFolderIdx = db2html(rsget("foldertitle"))
    End IF
    rsget.close
    
end function

''대 폴더 구분 select box
Sub sbDrawEtcImgGbn(ByVal selName,ByVal sIDValue,ByVal sScript)
    Dim sqlStr, arrList, intLoop
    sqlStr = " SELECT folderidx,foldertitle"
    sqlStr = sqlStr & "  from db_event.[dbo].[tbl_etcImage_master] "
    sqlStr = sqlStr & " WHERE isusing='Y'  order by sortkey, folderidx desc"
    
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    IF not rsget.eof THEN
        arrList = rsget.getRows()
    End IF
    rsget.close
    
%>
	<select name="<%=selName%>" <%=sScript%> class="Select">
	<option value="">선택</option>
<%
   If isArray(arrList) THEN
   		For intLoop = 0 To UBound(arrList,2)
%>
	<option value="<%=arrList(0,intLoop)%>" <%if Cstr(arrList(0,intLoop)) = Cstr(sIDValue) then %>selected<%end if%>><%=arrList(1,intLoop)%></option>
<%
   		Next
   End IF
%>
	</select>
<%
end Sub

Class CEtcImageItem
    public FfolderIdx
    public FfolderTitle
    public FrealPath
    public Fsortkey
    public Fisusing
    public FetcimgIdx
    public Fsubfolder
    public Fimagename
    public Freguserid
    public Flastuserid
    public Fregdate
    public Flastupdate
    public Fdeldt
    
    public function isDeletedItem()
        isDeletedItem = Not isNULL(Fdeldt)
    end function

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

Class CEtcImageManage
    public FItemList()
	public FOneItem
	public FCurrPage
	public FPageSize
	public FTotalCount
	public FResultCount
	public FTotalPage
	public FScrollCount
	public FPageCount
	
	public FRectFolderidx
	public FRectSubFolder
	public FRectDelYN
	public FRectetcimgIdx

	'//admin/eventmanage/etcimage/popImageReg.asp
    public Sub getEtcImage_one()
        dim sqlStr, sqladd

		if FRectetcimgIdx <> "" then
			sqladd = sqladd & " and d.etcimgIdx = "& FRectetcimgIdx &"" & VBCRLF
		end if

        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr&" m.folderIdx, m.folderTitle, m.realPath, m.sortkey, m.isusing " & VBCRLF
		sqlStr = sqlStr&" , d.etcimgIdx, d.subfolder, d.imagename, d.reguserid, d.lastuserid, d.regdate, d.lastupdate, d.deldt " & VBCRLF
		sqlStr = sqlStr&" FROM db_event.[dbo].[tbl_etcImage_master] m " & VBCRLF
		sqlStr = sqlStr&"   Join db_event.[dbo].[tbl_etcImage_detail] d " & VBCRLF
		sqlStr = sqlStr&"   on m.folderIdx=d.folderIdx" & VBCRLF
		sqlStr = sqlStr&" WHERE 1=1 " & sqladd

        'response.write sqlStr & "<br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CEtcImageItem
        if Not rsget.Eof then
	        FOneItem.FfolderIdx     = rsget("folderIdx")
            FOneItem.FfolderTitle   = db2html(rsget("folderTitle"))
            FOneItem.FrealPath      = rsget("realPath")
            FOneItem.Fsortkey       = rsget("sortkey")
            FOneItem.Fisusing       = rsget("isusing")
            FOneItem.FetcimgIdx     = rsget("etcimgIdx")
            FOneItem.Fsubfolder     = rsget("subfolder")
            FOneItem.Fimagename     = rsget("imagename")
            FOneItem.Freguserid     = rsget("reguserid")
            FOneItem.Flastuserid    = rsget("lastuserid")
            FOneItem.Fregdate       = rsget("regdate")
            FOneItem.Flastupdate       = rsget("lastupdate")
            FOneItem.Fdeldt         = rsget("deldt")
        end if
        rsget.Close
    end Sub

	public sub getEtcImageList()
		dim sqlStr,i, vSubQuery

		If FRectDelYN = "Y" Then
			vSubQuery = vSubQuery & " AND d.deldt is Not NULL"& VBCRLF
		elseif FRectDelYN = "N" Then
		    vSubQuery = vSubQuery & " AND d.deldt is NULL"& VBCRLF
		End IF

		If FRectFolderidx <> "" Then
			vSubQuery = vSubQuery & " AND d.folderidx="&FRectFolderidx& VBCRLF
		End IF
		
		If FRectSubFolder <> "" Then
			vSubQuery = vSubQuery & " AND d.subfolder='"&FRectSubFolder&"'"& VBCRLF
		End IF

		'총 갯수 구하기
		sqlStr = "SELECT COUNT(*) as cnt" & VBCRLF
		sqlStr = sqlStr &" FROM db_event.[dbo].[tbl_etcImage_master] m " & VBCRLF
		sqlStr = sqlStr &"   Join db_event.[dbo].[tbl_etcImage_detail] d " & VBCRLF
		sqlStr = sqlStr &"   on m.folderIdx=d.folderIdx" & VBCRLF
		sqlStr = sqlStr &" WHERE 1=1 " & vSubQuery

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & " " & VBCRLF
		sqlStr = sqlStr &" m.folderIdx, m.folderTitle, m.realPath, m.sortkey, m.isusing " & VBCRLF
		sqlStr = sqlStr &" , d.etcimgIdx, d.subfolder, d.imagename, d.reguserid, d.lastuserid, d.regdate, d.lastupdate, d.deldt " & VBCRLF
		sqlStr = sqlStr &" FROM db_event.[dbo].[tbl_etcImage_master] m " & VBCRLF
		sqlStr = sqlStr &"   Join db_event.[dbo].[tbl_etcImage_detail] d " & VBCRLF
		sqlStr = sqlStr &"   on m.folderIdx=d.folderIdx" & VBCRLF
		sqlStr = sqlStr &" WHERE 1=1 " & vSubQuery
		sqlStr = sqlStr &" ORDER BY d.etcimgidx DESC "

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CEtcImageItem

		        FItemList(i).FfolderIdx     = rsget("folderIdx")
                FItemList(i).FfolderTitle   = db2html(rsget("folderTitle"))
                FItemList(i).FrealPath      = rsget("realPath")
                FItemList(i).Fsortkey       = rsget("sortkey")
                FItemList(i).Fisusing       = rsget("isusing")
    
                FItemList(i).FetcimgIdx     = rsget("etcimgIdx")
                FItemList(i).Fsubfolder     = rsget("subfolder")
                FItemList(i).Fimagename     = rsget("imagename")
                FItemList(i).Freguserid     = rsget("reguserid")
                FItemList(i).Flastuserid    = rsget("lastuserid")
                FItemList(i).Fregdate       = rsget("regdate")
                FItemList(i).Flastupdate       = rsget("lastupdate")
                FItemList(i).Fdeldt         = rsget("deldt")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

    public Sub getEtcImage_masterone()
        dim sqlStr, sqladd

		if FRectfolderIdx <> "" then
			sqladd = sqladd & " and m.folderIdx = "& FRectfolderIdx &"" & VBCRLF
		end if

        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr&" m.folderIdx, m.folderTitle, m.realPath, m.sortkey, m.isusing " & VBCRLF
		sqlStr = sqlStr&" FROM db_event.[dbo].[tbl_etcImage_master] m " & VBCRLF
		sqlStr = sqlStr&" WHERE 1=1 " & sqladd

        'response.write sqlStr & "<br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new CEtcImageItem
        if Not rsget.Eof then
	        FOneItem.FfolderIdx     = rsget("folderIdx")
            FOneItem.FfolderTitle   = db2html(rsget("folderTitle"))
            FOneItem.FrealPath      = rsget("realPath")
            FOneItem.Fsortkey       = rsget("sortkey")
            FOneItem.Fisusing       = rsget("isusing")
        end if
        rsget.Close
    end Sub

	public sub getEtcImagemasterList()
		dim sqlStr,i, vSubQuery

		'총 갯수 구하기
		sqlStr = "SELECT COUNT(*) as cnt" & VBCRLF
		sqlStr = sqlStr &" FROM db_event.[dbo].[tbl_etcImage_master] m " & VBCRLF

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		sqlStr = "SELECT Top " & (FPageSize * FCurrPage) & " " & VBCRLF
		sqlStr = sqlStr &" m.folderIdx, m.folderTitle, m.realPath, m.sortkey, m.isusing " & VBCRLF
		sqlStr = sqlStr &" FROM db_event.[dbo].[tbl_etcImage_master] m " & VBCRLF
		sqlStr = sqlStr &" ORDER BY m.folderIdx DESC "

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CEtcImageItem

		        FItemList(i).FfolderIdx     = rsget("folderIdx")
                FItemList(i).FfolderTitle   = db2html(rsget("folderTitle"))
                FItemList(i).FrealPath      = rsget("realPath")
                FItemList(i).Fsortkey       = rsget("sortkey")
                FItemList(i).Fisusing       = rsget("isusing")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

    Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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
End Class
%>