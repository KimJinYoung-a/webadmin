<%

Class CFingersChoiceItem

	public Fidx
	Public FMenuId
	public Flec_idx
	public Fisusing
	public FsortNo
	public Flec_title
	public FImageSmall

	public FRegYn
	public FDispYn
	public FLImitCount
	public FLimitsold
	public FRegStartDate
	public FRegEndDate
	public FLecStartDate

	public function IsSoldOut()
		IsSoldOut = (FRegYn="N") or ((FDispYn="N") and (FLImitCount-FLimitsold<1))
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

Class CFingersChoice
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	Public FRectMenuId
	Public FRectIsUsing
	public FRectStyleSerail

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

	public function GetImageFolerName(byval i)
		'GetImageFolerName = "0" + CStr(Clng(FItemList(i).Flec_idx\10000))
		GetImageFolerName = GetImageSubFolderByItemid(FItemList(i).Flec_idx)
	end function

	public Function GetFingersChoiceList()
		dim sqlStr,i

		sqlStr = "select count(c.idx) as cnt " + vbcrlf
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_fingersChoice c," + vbcrlf
		sqlStr = sqlStr + " [db_academy].dbo.tbl_lec_item i" + vbcrlf
		sqlStr = sqlStr + " where c.lec_idx = i.idx" + vbcrlf
		if FRectMenuId<>"" then
			sqlStr = sqlStr + " and c.MenuId = '" + FRectMenuId + "'" + vbcrlf
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " c.idx, c.MenuId, c.lec_idx, c.isusing, i.lec_title, i.smallimg, c.sortNo " + vbcrlf
		sqlStr = sqlStr + " ,i.reg_yn, i.disp_yn, i.reg_startday, i.reg_endday, i.lec_startday1, i.limit_count, i.limit_sold " + vbcrlf
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_fingersChoice c," + vbcrlf
		sqlStr = sqlStr + " [db_academy].dbo.tbl_lec_item i" + vbcrlf
		sqlStr = sqlStr + " where c.lec_idx = i.idx " + vbcrlf

		if FRectMenuId<>"" then
			sqlStr = sqlStr + " and c.MenuId = '" + FRectMenuId + "'" + vbcrlf
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if
		sqlStr = sqlStr + " order by c.sortNo, c.idx desc"
'response.write sqlStr
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CFingersChoiceItem

				FItemList(i).Fidx		= rsACADEMYget("idx")
				FItemList(i).FMenuId		= rsACADEMYget("MenuId")
				FItemList(i).Flec_idx	= rsACADEMYget("lec_idx")
				FItemList(i).Fisusing	= rsACADEMYget("isusing")
				FItemList(i).Flec_title	= db2html(rsACADEMYget("lec_title"))
				FItemList(i).FImageSmall = imgFingers & "/lectureitem/small/" + GetImageSubFolderByItemid(FItemList(i).Flec_idx) + "/" + rsACADEMYget("smallimg")
				FItemList(i).FsortNo	= rsACADEMYget("sortNo")

				FItemList(i).FRegYn		= rsACADEMYget("Reg_Yn")
				FItemList(i).FDispYn		= rsACADEMYget("Disp_Yn")
				FItemList(i).FLImitCount		= rsACADEMYget("LImit_Count")
				FItemList(i).Flimitsold		= rsACADEMYget("limit_sold")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
	end function

	public Function GetFingersnewChoiceList()	'�Ź���
		dim sqlStr,i

		sqlStr = "select count(c.idx) as cnt " + vbcrlf
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_fingersChoice c," + vbcrlf
		sqlStr = sqlStr + " [db_academy].dbo.tbl_lec_item i" + vbcrlf
		sqlStr = sqlStr + " where c.lec_idx = i.idx" + vbcrlf
		if FRectMenuId<>"" then
			sqlStr = sqlStr + " and c.MenuId = '" + FRectMenuId + "'" + vbcrlf
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " c.idx, c.MenuId, c.lec_idx, c.isusing, i.lec_title, i.smallimg, c.sortNo " + vbcrlf
		sqlStr = sqlStr + " ,i.reg_yn, i.disp_yn, i.reg_startday, i.reg_endday, i.lec_startday1, i.limit_count, i.limit_sold " + vbcrlf
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_fingersChoice c," + vbcrlf
		sqlStr = sqlStr + " [db_academy].dbo.tbl_lec_item i" + vbcrlf
		sqlStr = sqlStr + " where c.lec_idx = i.idx " + vbcrlf

		if FRectMenuId<>"" then
			sqlStr = sqlStr + " and c.MenuId = '" + FRectMenuId + "'" + vbcrlf
		end if
		if FRectIsUsing<>"" then
			sqlStr = sqlStr + " and c.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if
		sqlStr = sqlStr + " order by c.idx desc, c.sortNo asc"

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CFingersChoiceItem

				FItemList(i).Fidx			= rsACADEMYget("idx")
				FItemList(i).FMenuId		= rsACADEMYget("MenuId")
				FItemList(i).Flec_idx		= rsACADEMYget("lec_idx")
				FItemList(i).Fisusing		= rsACADEMYget("isusing")
				FItemList(i).Flec_title		= db2html(rsACADEMYget("lec_title"))
				FItemList(i).FImageSmall 	= imgFingers & "/lectureitem/small/" + GetImageSubFolderByItemid(FItemList(i).Flec_idx) + "/" + rsACADEMYget("smallimg")
				FItemList(i).FsortNo		= rsACADEMYget("sortNo")
				FItemList(i).FRegYn			= rsACADEMYget("Reg_Yn")
				FItemList(i).FDispYn		= rsACADEMYget("Disp_Yn")
				FItemList(i).FLImitCount	= rsACADEMYget("LImit_Count")
				FItemList(i).Flimitsold		= rsACADEMYget("limit_sold")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
	end function

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

Function getLecMenuName(mnid)
	Select Case mnid
		Case "10"
			getLecMenuName = "������ü"
		Case "20"
			getLecMenuName = "������ Ŭ����"
		Case "30"
			getLecMenuName = "����� Ŭ����"
		Case "40"
			getLecMenuName = "��Ʃ��� ��ũ��"
		Case "50"
			getLecMenuName = "�����ӹ� Ŭ����"
	End Select
End Function

Function getLecMenunewName(mnid)
	Select Case mnid
		Case "1"	getLecMenunewName = "������ü"
		Case "10"	getLecMenunewName = "������"
		Case "20"	getLecMenunewName = "��ű�"
		Case "30"	getLecMenunewName = "�ٹ̱�"
		Case "40"	getLecMenunewName = "������"
		Case "50"	getLecMenunewName = "�׸���"
		Case "60"	getLecMenunewName = "����"
		Case "110"	getLecMenunewName = "������ Ŭ����"
		Case "120"	getLecMenunewName = "��Ŭ�� Ŭ����"
		Case "220"	getLecMenunewName = "��Ʃ���"
	End Select
End Function
%>
