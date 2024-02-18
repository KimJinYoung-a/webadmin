<%
Class cPartnerInfoReqItem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub


	public Ftidx
	public Fgubun
	public Freguserid
	public Fgroupid
	public Fgroupid_old
	public Fcompany_name
	public Fcompany_no
	public Fcompany_name_old
	public Fcompany_no_old
	public Fstatus
	public Fregdate
	public Fusername
	public Fceoname
	public Fjungsan_gubun
	public Fcompany_zipcode
	public Fcompany_address
	public Fcompany_address2
	public Fcompany_uptae
	public Fcompany_upjong
	public Fcompany_tel
	public Fcompany_fax
	public Freturn_zipcode
	public Freturn_address
	public Freturn_address2
	public Fjungsan_bank
	public Fjungsan_acctno
	public Fjungsan_acctname
	public Fjungsan_date
	public Fjungsan_date_off
	public Fmanager_name
	public Fmanager_phone
	public Fmanager_hp
	public Fmanager_email
	public Fjungsan_name
	public Fjungsan_phone
	public Fjungsan_hp
	public Fjungsan_email
	public Flastupdate
	public Fconfirmuserid
	public FComment
	public FBrandList
    public FdecCompNo ''암호화 해제한 사업자(주민)번호
    
    public function getDecCompNo()
        if isNULL(FdecCompNo) then
            ''getDecCompNo = ""
			if (Fcompany_no<>"") and (LEN(TRIM(replace(Fcompany_no,"-","")))=10) then
                getDecCompNo = Fcompany_no
            else
                getDecCompNo = ""
            end if
        else
            getDecCompNo = FdecCompNo
        end if
    end function
    
	public function getBrandList()
		if Right(FBrandList,1)="," then
			getBrandList = Left(FBrandList,Len(FBrandList)-1)
		else
			getBrandList = FBrandList
		end if
	end function

end class

class cPartnerInfoReq
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	public Ftidx
	public Fgroupid
	public Freqgubun
	public Freqname
	public Freqcompany
	public Freqgcode
	public Freqgcodegubun
	public FreqcompanyNo
	public Freqstatus


	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub


	'####### 변경 신청서 리스트 #######
	public sub fRequestlist()
		Dim sqlStr, addSql, i

		If Freqgubun <> "" Then
			addSql = addSql & " AND I.gubun = '" & Freqgubun & "' "
		End IF

		If Freqname <> "" Then
			addSql = addSql & " AND U.username Like '%" & Freqname & "%' "
		End IF

		If Freqcompany <> "" Then
			addSql = addSql & " AND I.company_name Like '%" & Freqcompany & "%' "
		End IF

		If Freqgcodegubun = "1" Then
			If Freqgcode <> "" Then
				addSql = addSql & " AND I.groupid = '" & Freqgcode & "' "
			End IF
		ElseIf Freqgcodegubun = "2" Then
			addSql = addSql & " AND I.groupid = '' "
		End IF

		If FreqcompanyNo <> "" Then
			addSql = addSql & " AND I.company_no = '" & FreqcompanyNo & "' "
		End If

		If Freqstatus <> "" Then
			addSql = addSql & " AND I.status = '" & Freqstatus & "' "
		End If


		'총 갯수 구하기
		sqlStr = "SELECT" & vbcrlf
		sqlStr = sqlStr & " COUNT(I.tidx) as cnt, CEILING(CAST(Count(I.tidx) AS FLOAT)/" & FPageSize & ") as totPg" & vbcrlf
		sqlStr = sqlStr & " FROM [db_partner].[dbo].[tbl_partner_temp_info] AS I" & vbcrlf
		sqlStr = sqlStr & " 	LEFT JOIN [db_partner].[dbo].[tbl_user_tenbyten] AS U ON I.reguserid = U.userid and U.userid<>''  " & vbcrlf
		sqlStr = sqlStr & " WHERE 1=1 " & addSql

		rsget.Open sqlStr,dbget,1
			FTotalCount	= rsget("cnt")
			FTotalPage	= rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		'데이터 리스트
		sqlStr = "SELECT TOP " & Cstr(FPageSize * FCurrPage) & vbcrlf
		sqlStr = sqlStr & " I.tidx, I.gubun, I.reguserid, U.username, I.groupid, I.groupid_old, I.company_name, I.company_no, I.status, I.regdate, I.lastupdate, G.company_name AS company_name_old, G.company_no AS company_no_old " & vbcrlf
		sqlStr = sqlStr & " FROM [db_partner].[dbo].[tbl_partner_temp_info] AS I " & vbcrlf
		sqlStr = sqlStr & " 	LEFT JOIN [db_partner].[dbo].[tbl_partner_temp_makerid] AS M ON I.tidx = M.tidx " & vbcrlf
		sqlStr = sqlStr & " 	LEFT JOIN [db_partner].[dbo].[tbl_user_tenbyten] AS U ON I.reguserid = U.userid and U.userid<>'' " & vbcrlf
		sqlStr = sqlStr & " 	LEFT JOIN [db_partner].[dbo].[tbl_partner_group] AS G ON I.groupid_old = G.groupid " & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & addSql
		sqlStr = sqlStr & " GROUP BY I.tidx, I.gubun, I.reguserid, U.username, I.groupid, I.groupid_old, I.company_name, I.company_no, I.status, I.regdate, I.lastupdate, G.company_name, G.company_no " & vbcrlf
		sqlStr = sqlStr & " ORDER BY I.tidx DESC" & vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cPartnerInfoReqItem

				FItemList(i).Ftidx				= rsget("tidx")
				FItemList(i).Fgubun				= rsget("gubun")
				FItemList(i).Freguserid			= rsget("reguserid")
				FItemList(i).Fusername			= rsget("username")
				FItemList(i).Fgroupid			= rsget("groupid")
				FItemList(i).Fgroupid_old		= rsget("groupid_old")
				FItemList(i).Fcompany_name		= db2html(rsget("company_name"))
				FItemList(i).Fcompany_no		= rsget("company_no")
				FItemList(i).Fcompany_name_old	= db2html(rsget("company_name_old"))
				FItemList(i).Fcompany_no_old	= rsget("company_no_old")
				FItemList(i).Fstatus			= rsget("status")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Flastupdate		= rsget("lastupdate")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub


	'####### 변경 신청서 상세 내용 #######
	public Sub fRequestDetail()
		Dim sqlStr, addSql, i
		sqlStr = "SELECT " & vbcrlf
		sqlStr = sqlStr & "I.company_no, I.groupid, isNull(I.groupid_old,'') AS groupid_old, I.company_name, I.ceoname, I.jungsan_gubun, I.company_zipcode, I.company_address, I.company_address2, I.company_uptae, " & vbcrlf
		sqlStr = sqlStr & "I.company_upjong, I.company_tel, I.company_fax, I.return_zipcode, I.return_address, I.return_address2, I.jungsan_bank, I.jungsan_acctno, I.jungsan_acctname, I.jungsan_date, " & vbcrlf
		sqlStr = sqlStr & "I.jungsan_date_off, I.manager_name, I.manager_phone, I.manager_hp, I.manager_email, I.regdate, I.lastupdate, I.confirmuserid, I.reguserid, U.username, I.comment, I.status, " & vbcrlf
		sqlStr = sqlStr & "I.jungsan_name, I.jungsan_phone, I.jungsan_hp, I.jungsan_email,"
		''sqlStr = sqlStr & "[db_partner].[dbo].[uf_DecSOCNoPH1](I.encCompNo) as decCompNo " & vbCrLf
		sqlStr = sqlStr & "db_cs.[dbo].[uf_DecCompanyNoAES256](e.encCompNo64) as decCompNo64 " & vbCrLf
		sqlStr = sqlStr & "FROM [db_partner].[dbo].[tbl_partner_temp_info] AS I " & vbcrlf
		sqlStr = sqlStr & " 	LEFT JOIN [db_partner].[dbo].[tbl_user_tenbyten] AS U ON I.confirmuserid = U.userid AND I.confirmuserid <> '' and U.userid<>''  " & vbcrlf
		sqlStr = sqlStr & " 	Left join [db_partner].[dbo].[tbl_partner_temp_info_adddata] e on I.tidx=e.tidx " & vbcrlf
		sqlStr = sqlStr & "WHERE I.tidx = '" & Ftidx & "'" & vbcrlf

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		set FOneItem = new cPartnerInfoReqItem
		if Not rsget.Eof then
			FOneItem.Fcompany_no		= rsget("company_no")
			FOneItem.Fgroupid			= db2html(rsget("groupid"))
			FOneItem.Fgroupid_old		= db2html(rsget("groupid_old"))
			FOneItem.Fcompany_name		= db2html(rsget("company_name"))
			FOneItem.Fceoname			= db2html(rsget("ceoname"))
			FOneItem.Fjungsan_gubun		= rsget("jungsan_gubun")
			FOneItem.Fcompany_zipcode	= rsget("company_zipcode")
			FOneItem.Fcompany_address	= db2html(rsget("company_address"))
			FOneItem.Fcompany_address2	= db2html(rsget("company_address2"))
			FOneItem.Fcompany_uptae		= db2html(rsget("company_uptae"))
			FOneItem.Fcompany_upjong	= db2html(rsget("company_upjong"))
			FOneItem.Fcompany_tel		= rsget("company_tel")
			FOneItem.Fcompany_fax		= rsget("company_fax")
			FOneItem.Freturn_zipcode	= rsget("return_zipcode")					'업체사무실주소로전용
			FOneItem.Freturn_address	= db2html(rsget("return_address"))
			FOneItem.Freturn_address2	= db2html(rsget("return_address2"))
			FOneItem.Fjungsan_bank		= rsget("jungsan_bank")
			FOneItem.Fjungsan_acctno	= rsget("jungsan_acctno")
			FOneItem.Fjungsan_acctname	= db2html(rsget("jungsan_acctname"))
			FOneItem.Fjungsan_date		= rsget("jungsan_date")
			FOneItem.Fjungsan_date_off	= rsget("jungsan_date_off")
			FOneItem.Fmanager_name		= db2html(rsget("manager_name"))
			FOneItem.Fmanager_phone		= rsget("manager_phone")
			FOneItem.Fmanager_hp		= rsget("manager_hp")
			FOneItem.Fmanager_email		= rsget("manager_email")
			FOneItem.Fjungsan_name		= rsget("jungsan_name")
			FOneItem.Fjungsan_phone		= rsget("jungsan_phone")
			FOneItem.Fjungsan_hp		= rsget("jungsan_hp")
			FOneItem.Fjungsan_email		= rsget("jungsan_email")
			FOneItem.Fregdate			= rsget("regdate")
			FOneItem.Flastupdate		= rsget("lastupdate")
			FOneItem.Fconfirmuserid		= rsget("confirmuserid")
			FOneItem.Freguserid			= rsget("reguserid")
			FOneItem.Fusername			= rsget("username")
			FOneItem.FComment			= db2html(rsget("comment"))
			FOneItem.Fstatus			= rsget("status")
			
			FOneItem.FdecCompNo         = rsget("decCompNo64")
		end if
		rsget.close

        Dim bufStr
		sqlStr = "SELECT makerid FROM [db_partner].[dbo].[tbl_partner_temp_makerid] WHERE tidx = '" & Ftidx & "'"
		rsget.Open sqlStr,dbget,1
			do until rsget.eof
				bufStr = rsget("makerid")

				FOneItem.FBrandList = FOneItem.FBrandList + bufStr + ","
				rsget.movenext
			loop
		rsget.close
	end sub


	'####### 사업자번호 변경 된 경우 변경 전 그룹코드가져와야함 #######
	public Sub fTIdxGroupID_OLD()
		Dim sqlStr, i

		sqlStr = "SELECT groupid, groupid_old FROM [db_partner].[dbo].[tbl_partner_temp_info] AS I WHERE tidx = '" & Ftidx & "' "
		rsget.Open sqlStr,dbget,1

		set FOneItem = new cPartnerInfoReqItem
			FOneItem.Fgroupid		= rsget("groupid")
			FOneItem.Fgroupid_old 	= rsget("groupid_old")
		rsget.Close
	end sub


	'####### 협조문첨부파일리스트 #######
	public Function fnGetFileList
		Dim strSql
		strSql = "	SELECT file_idx, file_name, real_name " & _
				"		FROM [db_partner].[dbo].tbl_partner_temp_file " & _
				"	WHERE tidx = '" & Ftidx & "' " & _
				"	ORDER BY file_idx ASC "
		rsget.Open strSql,dbget,1
		'response.write strSql
		IF not rsget.EOF THEN
			fnGetFileList = rsget.getRows()
		END IF
		rsget.close
	End Function


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


Function RequestStateName(value)
	SELECT CASE value
		CASE "0" : RequestStateName = "삭제"
		CASE "1" : RequestStateName = "신청"
		CASE "2" : RequestStateName = "작업중"
		CASE "3" : RequestStateName = "변경완료"
		CASE "5" : RequestStateName = "등록완료"
	END SELECT
End Function


Function RequestDocumentName(value)
	SELECT CASE value
		CASE "newcompreg" : RequestDocumentName = "사업자등록(신규)"
		CASE "companyreginfo" : RequestDocumentName = "사업자등록정보"
		CASE "bankinfo" : RequestDocumentName = "결제계좌정보"
		CASE "jungsandate" : RequestDocumentName = "정산일정보"
	END SELECT
End Function
%>
