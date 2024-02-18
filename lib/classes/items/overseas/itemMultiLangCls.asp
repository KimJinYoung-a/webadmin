<%
'####################################################
' Description : 상품 다국어정보 클래스
' History : 2013.07.11 허진원 생성
'####################################################

'//해외언어 선택상자
Sub drawSelectboxMultiLangCountrycd(selBoxName, selVal, chplg)
	dim strRst

	strRst = "<select name='" & selBoxName & "' " & chplg & ">" & vbCrLf
	strRst = strRst & "	<option value='' " & chkIIF(selVal=""," selected","") & ">:: 선택 ::</option>" & vbCrLf
	strRst = strRst & "	<option value='KR' " & chkIIF(selVal="KR"," selected","") & ">한국어</option>" & vbCrLf
	strRst = strRst & "	<option value='EN' " & chkIIF(selVal="EN"," selected","") & ">영어</option>" & vbCrLf
	strRst = strRst & "	<option value='CN' " & chkIIF(selVal="CN"," selected","") & ">중국어</option>" & vbCrLf
	strRst = strRst & "	<option value='JP' " & chkIIF(selVal="JP"," selected","") & ">일본어</option>" & vbCrLf
	strRst = strRst & "	<option value='ITSWEB' " & chkIIF(selVal="ITSWEB"," selected","") & ">아이띵소 해외</option>" & vbCrLf
	strRst = strRst & "</select>"

	Response.Write strRst
end Sub

function getCountryCdName(ctrCd)
	Select Case ctrCd
		Case "KR"
			getCountryCdName = "한국어"
		Case "EN"
			getCountryCdName = "English"
		Case "CN"
			getCountryCdName = "汉文"
		Case "JP"
			getCountryCdName = "にほんご"
		Case "ITSWEB"
			getCountryCdName = "아이띵소 해외"
		Case Else
			getCountryCdName = "English"
	End Select
end function

Class CMultiLangItem
	public Foptisusing
	public Foptsellyn
	public Foptlimityn
	public Foptlimitno
	public Foptlimitsold
	public Foptionname
	public Foptiontypename
	public FLastUpdate
    public Fitemid
    public Fmakerid
    public Fitemdiv
    public Fitemgubun
    public Fitemname
    public Fregdate
    public Fisusing
    public Fisextusing
    public Fmwdiv
    public Foptioncnt
    public Fitemcopy
    public FExistMultiLang
	public fuseyn
	public fcountrycd

	public Fitemname_kr
	public Fitemcopy_kr
	public Fitemsource_kr
	public Fitemsize_kr
	public Fsourcearea_kr
	public Fmakername_kr
	public Fkeywords_kr
	public Foptionname_kr
	public Foptiontypename_kr

    ''tbl_item_Contents
    public Fkeywords
    public Fsourcearea
    public Fmakername
    public Fitemsource
    public Fitemsize
	public FitemWeight
    public Fusinghtml
    public Fitemcontent
    public Fdesignercomment

    public Fitemoption

    ''tbl_item_multiSite_regItem
    public fmultiplerate
	public FchkMultiLang

    Private Sub Class_Initialize()
        Foptioncnt = 0
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class


Class CMultiLang
    public FOneItem
	public FItemList()
	public FTotalCount
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectItemId
	public FRectCountryCd

	'// 상품 다국어 정보 접수
	public Sub GetMultiLangItemInfo()
		dim sqlstr,i

		sqlstr = "SELECT i.itemname as inm_kr, c.designercomment as icp_kr, c.sourcearea as isa_kr, c.itemsource as isc_kr, c.itemsize as isz_kr, c.makername as imk_kr, c.keywords as ikw_kr "
		sqlstr = sqlstr & " ,m.* "
		sqlstr = sqlstr & " FROM db_item.dbo.tbl_item as i "
		sqlstr = sqlstr & " 	join db_item.dbo.tbl_item_contents as c "
		sqlstr = sqlstr & "			on i.itemid=c.itemid "
		sqlstr = sqlstr & " 	left join db_item.dbo.tbl_item_multiLang AS m "
		sqlstr = sqlstr & "			on i.itemid=m.itemid AND m.countryCd = '" & FRectCountryCd & "' "
		sqlstr = sqlstr & " WHERE i.itemid=" & FRectItemID

		'response.write sqlstr & "<br>"
		rsget.Open sqlstr,dbget,1

		ftotalcount = rsget.recordcount

		set FOneItem = new CMultiLangItem
		if Not rsget.Eof then
			FOneItem.FchkMultiLang	= Not(isNull(rsget("itemname")))		'// 다국어 정보 등록 여부

			FOneItem.Fitemname 		= db2html(rsget("itemname"))
			FOneItem.Fitemcopy 		= db2html(rsget("itemcopy"))
			FOneItem.Fitemcontent 	= db2html(rsget("itemcontent"))
			FOneItem.Fitemsource	= db2html(rsget("itemsource"))
			FOneItem.Fitemsize		= db2html(rsget("itemsize"))
			FOneItem.Fsourcearea	= db2html(rsget("sourcearea"))
			FOneItem.Fmakername		= db2html(rsget("makername"))
			FOneItem.fuseyn		    = chkIIF(Not(rsget("useyn")="" or isNull(rsget("useyn"))),rsget("useyn"),"Y")
			FOneItem.fkeywords 		= db2html(rsget("keywords"))

			FOneItem.Fitemname_kr 	= db2html(rsget("inm_kr"))
			FOneItem.Fitemcopy_kr 	= db2html(rsget("icp_kr"))
			FOneItem.Fitemsource_kr	= db2html(rsget("isc_kr"))
			FOneItem.Fitemsize_kr	= db2html(rsget("isz_kr"))
			FOneItem.Fsourcearea_kr	= db2html(rsget("isa_kr"))
			FOneItem.Fmakername_kr	= db2html(rsget("imk_kr"))
			FOneItem.fkeywords_kr 	= db2html(rsget("ikw_kr"))
		end if
		rsget.Close

	end Sub


	'// 상품옵션 다국어 정보 접수
	public Sub GetItemOptionMultiLang
		dim sqlstr,i

		sqlstr = " select io.itemoption, io.optionTypeName as ot_kr, io.optionname as on_kr " & vbCrLf
		sqlstr = sqlstr + " 	,isNull(mo.countryCd,'" & frectCountryCd & "') as countryCd, isNull(mo.isUsing,'Y') as isUsing , mo.optionTypeName, mo.optionName " & vbCrLf
		sqlstr = sqlstr + " from db_item.dbo.tbl_item_option as io " & vbCrLf
		sqlstr = sqlstr + " 	left join db_item.dbo.tbl_item_multiLang_option as mo " & vbCrLf
		sqlstr = sqlstr + " 		on io.itemid=mo.itemid " & vbCrLf
		sqlstr = sqlstr + " 			and io.itemoption=mo.itemoption " & vbCrLf
		sqlstr = sqlstr + " 			and mo.CountryCd='" & frectCountryCd & "' " & vbCrLf
		sqlstr = sqlstr + " where io.itemid=" & FRectItemID & vbCrLf
		sqlstr = sqlstr + " order by io.itemoption "

		'response.write sqlstr &"<br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CMultiLangItem

				FItemList(i).Fitemoption	= rsget("itemoption")
				FItemList(i).Foptisusing	= rsget("isusing")
				FItemList(i).FchkMultiLang	= Not(isNull(rsget("optionname")))		'// 다국어 정보 등록 여부

				FItemList(i).Foptionname	= db2html(rsget("optionname"))
				FItemList(i).Foptiontypename = db2html(rsget("optiontypename"))
				FItemList(i).Foptionname_kr	= db2html(rsget("on_kr"))
				FItemList(i).Foptiontypename_kr = db2html(rsget("ot_kr"))

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
	end Sub

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

%>