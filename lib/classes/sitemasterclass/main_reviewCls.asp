<%
Class Ceventuser

	public fuserid
	public fjuminno
	public fusername
	public fusermail
	public fuserphone
	public fusercell
	public fzipcode
	public faddress1
	public fuseraddr
	public fLevel
	public fevtcom_txt
	public fWcnt
	public fWdate

	public fitemoption
	public fitemid
	public fcontents
	public FsortNo
	public fitemname
	public fcate_large
	public fcate_mid
	public fmakerid
	public FFile1
	public FFile2
	public Fregdate
	public FImageIcon1
	public FImageIcon2
	public FRectItemName
	public FRectKeyword

   Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end class
'##################################################################
class Ceventuserlist

	public flist

	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage

	public fSearchKey1
	public fSearchKey2
	public FRectMakerid
	public FRectCDL
	public FRectCDM
	public FRectCDS
	public FRectStartDt
	public FRectEndDt
	public FRectPhotoMode
	public FRectPoint

	public FRectSellYN
	public FRectKeyword

	public function frectseach()
		if fSearchKey2 <> "" then
			frectseach = "& fSearchKey2 &"
		else
			frectseach = 0
		end if
	end function


	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	Private Sub Class_Initialize()
		redim  flist(0)

		FCurrPage =1
		FPageSize = 15
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public function GetImageFolerName(byval i)
		'GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
		GetImageFolerName = GetImageSubFolderByItemid(flist(i).FItemID)
	end function

	Sub drawSelectBoxSell(selectBoxName,selectedId)
   dim tmp_str,query1
   %>
   <select class="select" name="<%=selectBoxName%>">
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >판매</option>
   <option value="S" <% if selectedId="S" then response.write "selected" %> >일시품절</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >품절</option>
   <option value="YS" <% if selectedId="YS" then response.write "selected" %> >판매+일시품절</option>
   </select>
   <%
	End sub

	public sub Feventuserlist99
	dim sql , i , AddSQL
		sql = ""

		if fSearchKey1<>"" then
				AddSQL = AddSQL & " and a.userid='" & fSearchKey1 & "' "
		end if
		if fSearchKey2<>"" then
				AddSQL = AddSQL & " and a.itemid=" & fSearchKey2 & " "
		end if
		IF (FRectMakerid<>"") then
		    AddSQL = AddSQL & " and b.makerid='"&FRectMakerid&"'"
		end if
		if FRectCDL<>"" then
			AddSQL = AddSQL & " and b.cate_large='" & FRectCDL & "' "
		end if
		if FRectCDM<>"" then
			AddSQL = AddSQL & " and b.cate_mid='" & FRectCDM & "' "
		end if
		if FRectCDS<>"" then
			AddSQL = AddSQL & " and b.cate_small='" & FRectCDS & "' "
		end if
		if Not(FRectStartDt="" or FRectEndDt="") then
			AddSQL = AddSQL & " and a.regdate between '" & FRectStartDt & "' and DateAdd(day,1,'" & FRectEndDt & "') "
		end if
		IF FRectPhotoMode="on" then
			AddSQL = AddSQL & " and a.File1 is not null "
		End IF
        if (FRectSellYN <> "") then
            AddSQL = AddSQL & " and B.sellyn='" + FRectSellYN + "'"
        end if
        if (FRectKeyword <> "") then
            AddSQL = AddSQL & " and B.itemname like '%" + html2db(FRectKeyword) + "%'"
        end if

		sql = sql & "select count(*),CEILING(CAST(Count(*) AS FLOAT)/"& FPageSize &") from [db_board].[dbo].[tbl_Item_Evaluate] as a "
		sql = sql & "join db_item.dbo.tbl_item as b "
		sql = sql & "on a.itemid = b.itemid "
		sql = sql & "where 1=1 " & AddSQL & " and a.isusing='Y' and a.totalpoint='4'"
		'response.write sql
		rsget.Open SQL,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		sql = ""
		sql = sql & "select top "& CStr(FPageSize*FCurrPage) & " a.userid, a.itemid, a.itemoption, a.contents, b.itemname, b.cate_large, b.cate_mid, b.cate_small, b.makerid, a.file1, a.file2, b.sellyn ,a.totalpoint, a.regdate "
		sql = sql & "from [db_board].[dbo].[tbl_Item_Evaluate] as a "
		sql = sql & "join db_item.dbo.tbl_item as b "
		sql = sql & "on a.itemid = b.itemid "
		sql = sql & "where 1=1 " & AddSQL & " and a.isusing='Y' and a.totalpoint='4' order by idx desc"
		'response.write sql
		rsget.pagesize = FPageSize
		rsget.open sql,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		if FResultCount<1 then FResultCount=0
		redim preserve flist(FResultCount)
		i = 0
	if not rsget.eof then
		rsget.absolutepage = FCurrPage
		do until rsget.eof
			set flist(i) = new Ceventuser

			flist(i).fuserid = rsget("userid")
			flist(i).fitemid = rsget("itemid")
			flist(i).fitemoption = rsget("itemoption")
			flist(i).fcontents = rsget("contents")
			flist(i).fitemname = rsget("itemname")
			flist(i).fcate_large = rsget("cate_large")
			flist(i).fcate_mid = rsget("cate_mid")
			flist(i).fmakerid = rsget("makerid")
			flist(i).FFile1 = rsget("File1")
			flist(i).FFile2 = rsget("File2")
			flist(i).Fregdate = rsget("regdate")

			IF Not(rsget("File1")="" or isNull(rsget("File1"))) Then
					flist(i).FImageIcon1		= "http://imgstatic.10x10.co.kr/goodsimage/" + GetImageFolerName(i) + "/" + rsget("File1")
			End IF
			IF Not(rsget("File2")="" or isNull(rsget("File2"))) Then
				flist(i).FImageIcon2		= "http://imgstatic.10x10.co.kr/goodsimage/" + GetImageFolerName(i) + "/" + rsget("File2")
			End IF

			rsget.movenext
			i = i+1
			loop
		end if
	rsget.close

	end sub

end class

'##################################################################
%>