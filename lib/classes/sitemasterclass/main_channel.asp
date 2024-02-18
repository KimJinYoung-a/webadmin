<%
'#######################################################
'	History	:  2009.04.18 한용민 생성
'	Description : 메인페이지 감성채널 클래스
'#######################################################
%>
<%
'/// MD 빅찬스/베스트 브랜드 클래스 ///
Class CSpecialItem

	public Fidx
	public Fcdl
	public Fcdm
	public Fitemid
	public Fisusing
	public Fcode_nm
	public Fcdm_nm
	public FitemName
	public FImageSmall
	public Fgubun
	public FsellYn
	public FsailYn
	public ForgPrice
	public FsailPrice

	public fimglink
	public FImage
	public Fregdate
	public FsortNo

	public Ftitleimgurl

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function GetMdGubun()
		if FGubun = "01" then
			GetMdGubun = "MD1"
		elseif FGubun = "02" then
			GetMdGubun = "MD2"
		elseif FGubun = "03" then
			GetMdGubun = "MD3"
		elseif FGubun = "04" then
			GetMdGubun = "MD4"
		end if
	end Function

end Class

Class CMDSRecommend
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectCDL
	public FRectCDM
	public FRectStyleSerail
	public FRectGubun
	public FRectIsUsing
	public FRectIdx

	Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub
	
	'// 베스트 브랜드 목록
	public Function GetBestBrandList()
		dim sqlStr, addSQL, i

		'추가 조건 쿼리
		if FRectCDL<>"" then
			addSQL = " and b.cdl='" + FRectCDL + "'"
		end if
		if FRectIdx<>"" then
			addSQL = " and b.idx=" & FRectIdx
		end if
		if FRectisusing<>"" then
			addSQL = addSQL & " and isusing='" & FRectisusing & "'"
		end if

		'목록 카운트
		sqlStr = "select count(idx), CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") from [db_sitemaster].dbo.tbl_main_channel as b "
		sqlStr = sqlStr + " where idx<>0" & addSQL

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FtotalPage	= rsget(1)
		rsget.Close

		'목록 내용 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " b.idx, b.cdl, b.imglink, b.imgfile, b.regdate, b.isusing, b.sortNo ,c.code_nm " + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].dbo.tbl_main_channel b " + vbcrlf
		sqlStr = sqlStr + " left join db_item.dbo.tbl_Cate_mid c" + vbcrlf
		sqlStr = sqlStr + " on b.cdl = c.code_mid" + vbcrlf
		sqlStr = sqlStr + " where " + vbcrlf
		sqlStr = sqlStr + " c.code_large = 110  and isusing='Y'" & addSQL 
		sqlStr = sqlStr + " order by b.sortno asc , b.idx desc" + vbcrlf
	
		'response.write sqlStr &"<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CSpecialItem

				FItemList(i).FIdx		= rsget("idx")
				FItemList(i).FCdL		= rsget("cdl")
				FItemList(i).fimglink	= db2html(rsget("imglink"))
				FitemList(i).FImage		= db2html(rsget("imgfile"))
				FItemList(i).Fcode_nm	= rsget("code_nm")
				FItemList(i).Fisusing	= rsget("isusing")
				FItemList(i).Fregdate	= rsget("regdate")
				FItemList(i).FsortNo	= rsget("sortNo")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	'// 페이지 관련 함수
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

function DrawSelectBoxmainchannel(boxname, stats)		
	dim userquery, tem_str

	response.write "<select name='" & boxname & "' style='width:140px; height:22px' class='input_01'>"	
	response.write "<option value=''>감성채널카테고리</option>"
	if stats = "010" then									
	response.write "<option value='010' selected>camera</option>"
	else 
	response.write "<option value='010'>camera</option>"	
	end if
	if stats = "020" then									
	response.write "<option value='020' selected>travel</option>"
	else 
	response.write "<option value='020'>travel</option>"	
	end if	
	if stats = "030" then									
	response.write "<option value='030' selected>music</option>"
	else 
	response.write "<option value='030'>music</option>"	
	end if
	if stats = "040" then									
	response.write "<option value='040' selected>book</option>"
	else 
	response.write "<option value='040'>book</option>"	
	end if	
	if stats = "050" then									
	response.write "<option value='050' selected>diy</option>"
	else 
	response.write "<option value='050'>diy</option>"	
	end if
	if stats = "060" then									
	response.write "<option value='060' selected>flower</option>"
	else 
	response.write "<option value='060'>flower</option>"	
	end if	
	if stats = "070" then									
	response.write "<option value='070' selected>taste</option>"
	else 
	response.write "<option value='070'>taste</option>"	
	end if
	if stats = "080" then									
	response.write "<option value='080' selected>beauty</option>"
	else 
	response.write "<option value='080'>beauty</option>"	
	end if	
	response.write "</select>"
End function
%>
