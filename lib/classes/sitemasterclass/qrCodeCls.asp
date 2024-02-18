<%
'#######################################################
'	History	:  2011.01.11 ������ ����
'			   2022.07.04 �ѿ�� ����(isms�������������, �ҽ�ǥ��ȭ)
'	Description : QR�ڵ� ����
'#######################################################

Class CQRCodeItem

	public FqrSn
	Public FqrTitle
	public FqrDiv
	public FcountYn
	public FqrContent
	public FqrImage
	public FisUsing
	public Fregdate
	public FqrHitCount

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class CQRCode
	public FItemList()

	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectQRSn
	public FRectQRDiv
	public FRectCntYn
	Public FRectIsUsing
	Public FRectkeyWd

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public Sub GetQRCode()
		dim sqlStr,addSql, i

		'�߰� ����
		if FRectQRSn<>"" then
			addSql = addSql + " and qrSn = '" + FRectQRSn + "'" + vbcrlf
		end if
		if FRectQRDiv<>"" then
			addSql = addSql + " and qrDiv = '" + FRectQRDiv + "'" + vbcrlf
		end if
		if FRectCntYn<>"" then
			addSql = addSql + " and countYn = '" + FRectCntYn + "'" + vbcrlf
		end if
		if Not(FRectIsUsing="A" or FRectIsUsing="") then
			addSql = addSql + " and isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if
		if FRectkeyWd<>"" then
			addSql = addSql + " and qrTitle like '%" + FRectkeyWd + "%'" + vbcrlf
		end if

		'�Ѽ� ����
		sqlStr = "select count(qrSn), CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") " + vbcrlf
		sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_QRCodeList with (nolock)" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit Sub
		end if

		'���� ����
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " qrSn, qrTitle, qrDiv, countYn, qrContent, qrImage, isUsing, regdate, qrHitCount " + vbcrlf
		sqlStr = sqlStr + " from db_sitemaster.dbo.tbl_QRCodeList with (nolock)" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf
		sqlStr = sqlStr + " order by qrSn desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CQRCodeItem

				FItemList(i).FqrSn			= rsget("qrSn")
				FItemList(i).FqrTitle		= db2html(rsget("qrTitle"))
				FItemList(i).FqrDiv			= rsget("qrDiv")
				FItemList(i).FcountYn		= rsget("countYn")
				FItemList(i).FqrContent		= db2html(rsget("qrContent"))
				FItemList(i).FqrImage		= staticImgUrl & "/mobile/QRCode/" & rsget("qrImage")
				FItemList(i).FisUsing		= rsget("isUsing")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).FqrHitCount	= rsget("qrHitCount")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end Sub

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

Sub DrawSelectBoxQRDiv(boxname,stats)
%>
<select name='<%=boxname%>' class='select'>
	<option value="">�����ϼ���</option>
	<option value=1 <% if stats = "1" then response.write " selected" %>>URL</option>
	<option value=2 disabled <% if stats = "2" then response.write " selected" %>>�ؽ�Ʈ</option>
	<option value=3 disabled <% if stats = "3" then response.write " selected" %>>�̹���</option>
	<option value=4 disabled <% if stats = "4" then response.write " selected" %>>������</option>
	<option value=5 disabled <% if stats = "5" then response.write " selected" %>>APP URL</option>
</select>
<%
end Sub
%>
