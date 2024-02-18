<%
'###########################################################
' Description :  통합 회원 카드 클래스
' History : 2009.07.08 강준구 생성
'			2010.06.04 한용민 수정
'###########################################################

Dim g_MenuPos	' ?? 왜있는거. 용만
	g_MenuPos = "1115"

class cTotalPoint_item
	public fCardNo
	public fPoint
	public FUseYN
	public FGrade
	public fJumin1
	public fJumin2_Enc		
	public forderno
	public fshopid
	public fregdate	
	public fpointuserno
	public fitemgubun
	public fitemid
	public fitemoption
	public fitemname
	public fitemoptionname
	public fsellprice
	public frealsellprice
	public fsuplyprice
	public fitemno
	public fOnLineUSerID
	public fcancelyn
	public fUserName
	public fUserSeq
	public fshopname
	public fregshopid
	public fusercount
	public fmemberY
	public fmemberN
	public ftotalsum
	public frealsum
	public fAuthIdx
	public fUserHp
	public fSmsYN
	public fKakaoTalkYN
	public fIsUsing
	public fLastUpdate
	public fipkumdiv

	Public function shopIpkumDivName()
		if Fipkumdiv="1" then
			shopIpkumDivName="배송지입력전"
		elseif Fipkumdiv="2" then
			shopIpkumDivName="배송지입력완료"
		elseif Fipkumdiv="5" then
			shopIpkumDivName="업체통보"
		elseif Fipkumdiv="6" then
			shopIpkumDivName="배송준비"
		elseif Fipkumdiv="7" then
			shopIpkumDivName="일부출고"
		elseif Fipkumdiv="8" then
			shopIpkumDivName="출고완료"
		end if
	end Function

end class

Class TotalPoint
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FUserSeq
	public FUserName
	public FUserID
	public FJumin1
	public FJumin2_Enc
	public FTotalCardNo
	public FTotalPoint
	public FCardNo
	public FCardGubun
	public FPoint
	public FGrade
	public FSexFlag
	public FTelNo
	public FHpNo
	public FZipCode
	public FAddress
	public FAddressDetail
	public FEmail
	public FEmailYN
	public FSMSYN
	public FUserStatus
	public FLastUpdate
	public FRegdate
	public FShopName
	public FUseYN	
	public FTotCnt
	public FCPage
	public FPSize
	public FRectShopID
	public frectmemberyn
	public frectorderno
	public FRectStartDay
	public FRectEndDay
	public FRectInc3pl
	public frectdatefg
	public FRectOldData
	public FRectuserhp
	public FRectBeaSongYN

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
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

	'//admin/totalpoint/shopuser_sum.asp
	public sub fshopuser_sum()
		dim sqlStr,i , strSubSql

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            strSubSql = strSubSql & " and isNULL(p.tplcompanyid,'')<>''"
	        end if
	    else
	        strSubSql = strSubSql & " and isNULL(p.tplcompanyid,'')=''"
	    end if
		'//회원가입여부
		If frectmemberyn = "Y" Then
			strSubSql = strSubSql & " and isnull(sc.userseq,0)<>0"
		elseIf frectmemberyn = "N" Then
			strSubSql = strSubSql & " and isnull(sc.userseq,0)=0"			
		End If
		
		If frectshopid <> "" Then
			strSubSql = strSubSql & " AND sc.regShopid = '" & frectshopid & "'"
		End If

		if FRectStartDay<>"" then
			strSubSql = strSubSql + " and sc.regdate>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			strSubSql = strSubSql + " and sc.regdate<'" + CStr(FRectEndDay) + "'"
		end if
		
		sqlStr = "select top 500"
		sqlStr = sqlStr & " sc.regshopid, su.shopname ,count(*) as usercount"
		sqlStr = sqlStr & " ,isnull(sum(case"
		sqlStr = sqlStr & " 	when isnull(sc.userseq,0)<>0 then 1 end),0) as memberY"
		sqlStr = sqlStr & " ,isnull(sum(case"
		sqlStr = sqlStr & " 	when isnull(sc.userseq,0)=0 then 1 end),0) as memberN"
		sqlStr = sqlStr & " from db_shop.dbo.tbl_total_shop_card sc"
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shop_user su"
		sqlStr = sqlStr & " 	on sc.regshopid=su.userid and su.isusing = 'Y'"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	    sqlStr = sqlStr & "       on sc.regshopid=p.id "
		sqlStr = sqlStr & " where sc.UseYn='Y'" & strSubSql
		sqlStr = sqlStr & " group by sc.regshopid, su.shopname ,su.shopname"
		sqlStr = sqlStr & " order by usercount desc"
				
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordcount

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cTotalPoint_item

				FItemList(i).fregshopid = rsget("regshopid")
				FItemList(i).fshopname = rsget("shopname")
				FItemList(i).fusercount = rsget("usercount")
				FItemList(i).fmemberY = rsget("memberY")
				FItemList(i).fmemberN = rsget("memberN")
																					
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'/admin/totalpoint/customer_sell_history.asp
	public sub fsell_history_master()
		dim sqlStr, i , sqlsearch

        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
	            sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')<>''" & vbcrlf
	        end if
	    else
	        sqlsearch = sqlsearch & " and isNULL(p.tplcompanyid,'')=''" & vbcrlf
	    end if
		If FUserName <> "" Then
			sqlsearch = sqlsearch & " AND u.UserName = '" & FUserName & "' " & vbcrlf
		End If		
		If FCardNo <> "" Then
			sqlsearch = sqlsearch & " AND m.pointuserno = '" & FCardNo & "' " & vbcrlf
		End If
		If FUserID <> "" Then
			sqlsearch = sqlsearch & " AND u.OnlineUserID = '" & FUserID & "' " & vbcrlf
		End If
		If frectshopid <> "" Then
			sqlsearch = sqlsearch & " AND m.shopid = '" & frectshopid & "' " & vbcrlf
		End If
		If FRectorderno <> "" Then
			sqlsearch = sqlsearch & " AND m.orderno = '" & FRectorderno & "' " & vbcrlf
		End If
		If FRectuserhp <> "" Then
			sqlsearch = sqlsearch & " AND replace(sc.userhp,'-','') = '" & replace(FRectuserhp,"-","") & "' " & vbcrlf
		End If

        '//주문일 기준
		if frectdatefg = "jumun" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch & " 	and m.shopregdate>='" + CStr(FRectStartDay) + "'" & vbcrlf
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch & " 	and m.shopregdate<'" + CStr(FRectEndDay) + "'" & vbcrlf
			end if

		'//매출일 기준
		elseif frectdatefg = "maechul" then
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch & " 	and m.IXyyyymmdd>='" + CStr(FRectStartDay) + "'" & vbcrlf
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch & " 	and m.IXyyyymmdd<'" + CStr(FRectEndDay) + "'" & vbcrlf
			end if
		else
			if FRectStartDay<>"" then
				sqlsearch = sqlsearch & " 	and m.shopregdate>='" + CStr(FRectStartDay) + "'" & vbcrlf
			end if
			if FRectEndDay<>"" then
				sqlsearch = sqlsearch & " 	and m.shopregdate<'" + CStr(FRectEndDay) + "'" & vbcrlf
			end if
		end if

		If FRectBeaSongYN <> "" Then
			If FRectBeaSongYN = "Y" Then
				sqlsearch = sqlsearch & " and bm.orderno is not null" & vbcrlf
			elseif FRectBeaSongYN = "N" Then
				sqlsearch = sqlsearch & " and bm.orderno is null" & vbcrlf
			end If
		end If

		'총 갯수 구하기
		sqlStr = "select count(*) as cnt" & vbcrlf

		if FRectOldData="on" then
			sqlStr = sqlStr & " 	from [db_shoplog].[dbo].tbl_old_shopjumun_master m" & vbcrlf
		else
	        sqlStr = sqlStr & "		from [db_shop].[dbo].tbl_shopjumun_master m" & vbcrlf
		end if

		sqlStr = sqlStr & " left Join [db_shop].[dbo].[tbl_total_shop_card] AS B " & vbcrlf
		sqlStr = sqlStr & " 	ON m.pointuserno = B.cardno" & vbcrlf
		sqlStr = sqlStr & " left JOIN [db_shop].[dbo].tbl_total_shop_user AS u" & vbcrlf
		sqlStr = sqlStr & " 	ON b.UserSeq = u.UserSeq" & vbcrlf
		sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_user s" & vbcrlf
		sqlStr = sqlStr & " 	on m.shopid = s.userid" & vbcrlf
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p" & vbcrlf
	    sqlStr = sqlStr & " 	on m.shopid=p.id " & vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shopjumun_sms_cert sc" & vbcrlf
	    sqlStr = sqlStr & " 	on m.orderno=sc.orderno" & vbcrlf
	    sqlStr = sqlStr & " 	and sc.isusing='Y'" & vbcrlf
		sqlStr = sqlStr & " left join db_shop.[dbo].[tbl_shopbeasong_order_master] bm" & vbcrlf
		sqlStr = sqlStr & " 	on m.orderno=bm.orderno" & vbcrlf
		sqlStr = sqlStr & " 	and bm.cancelyn='N'" & vbcrlf

		sqlStr = sqlStr & " where m.idx <>0 " & sqlsearch				

		'response.write sqlStr &"<br>"			
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " m.orderno , m.shopid , m.regdate , m.pointuserno ,m.totalsum, m.realsum" & vbcrlf
		sqlStr = sqlStr & " ,s.shopname ,u.OnLineUSerID , m.cancelyn , u.UserName ,u.UserSeq" & vbcrlf
		sqlStr = sqlStr & " , sc.AuthIdx, sc.UserHp, sc.SmsYN, sc.KakaoTalkYN" & vbcrlf
		sqlStr = sqlStr & " , sc.IsUsing, sc.LastUpdate, bm.ipkumdiv" & vbcrlf

		if FRectOldData="on" then
			sqlStr = sqlStr & " 	from [db_shoplog].[dbo].tbl_old_shopjumun_master m" & vbcrlf
		else
	        sqlStr = sqlStr & "		from [db_shop].[dbo].tbl_shopjumun_master m" & vbcrlf
		end if

		sqlStr = sqlStr & " left Join [db_shop].[dbo].[tbl_total_shop_card] AS B " & vbcrlf
		sqlStr = sqlStr & " 	ON m.pointuserno = B.cardno" & vbcrlf
		sqlStr = sqlStr & " left JOIN [db_shop].[dbo].tbl_total_shop_user AS u" & vbcrlf
		sqlStr = sqlStr & " 	ON b.UserSeq = u.UserSeq" & vbcrlf
		sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_user s" & vbcrlf
		sqlStr = sqlStr & " 	on m.shopid = s.userid" & vbcrlf
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p" & vbcrlf
	    sqlStr = sqlStr & " 	on m.shopid=p.id " & vbcrlf
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shopjumun_sms_cert sc" & vbcrlf
	    sqlStr = sqlStr & " 	on m.orderno=sc.orderno" & vbcrlf
	    sqlStr = sqlStr & " 	and sc.isusing='Y'" & vbcrlf
		sqlStr = sqlStr & " left join db_shop.[dbo].[tbl_shopbeasong_order_master] bm" & vbcrlf
		sqlStr = sqlStr & " 	on m.orderno=bm.orderno" & vbcrlf
		sqlStr = sqlStr & " 	and bm.cancelyn='N'" & vbcrlf

		sqlStr = sqlStr & " where m.idx <>0 " & sqlsearch
		sqlStr = sqlStr & " order by m.orderno Desc"

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
				set FItemList(i) = new cTotalPoint_item

				FItemList(i).fipkumdiv = rsget("ipkumdiv")
				FItemList(i).fshopname = db2html(rsget("shopname"))
				FItemList(i).fUserSeq = db2html(rsget("UserSeq"))
				FItemList(i).fUserName = db2html(rsget("UserName"))
				FItemList(i).fcancelyn = rsget("cancelyn")
				FItemList(i).fOnLineUSerID = rsget("OnLineUSerID")
				FItemList(i).forderno = rsget("orderno")
				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fpointuserno = rsget("pointuserno")
				FItemList(i).ftotalsum = rsget("totalsum")
				FItemList(i).frealsum = rsget("realsum")
				FItemList(i).fAuthIdx = rsget("AuthIdx")
				FItemList(i).fUserHp = rsget("UserHp")
				FItemList(i).fSmsYN = rsget("SmsYN")
				FItemList(i).fKakaoTalkYN = rsget("KakaoTalkYN")
				FItemList(i).fIsUsing = rsget("IsUsing")
				FItemList(i).fRegdate = rsget("Regdate")
				FItemList(i).fLastUpdate = rsget("LastUpdate")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'/admin/totalpoint/customer_sell_history_detail.asp
	public sub fsell_history_detail()
		dim sqlStr, i , sqlsearch

		If frectorderno <> "" Then
			sqlsearch = sqlsearch & " AND m.orderno = '" & frectorderno & "' "
		End If

		'데이터 리스트 
		sqlStr = "select top 1000"
		sqlStr = sqlStr & " m.orderno , m.shopid , m.regdate , m.pointuserno , d.itemgubun , d.itemid ,s.shopname" + vbcrlf
		sqlStr = sqlStr & " , d.itemoption , d.itemname ,d.itemoptionname , d.sellprice , d.realsellprice" + vbcrlf
		sqlStr = sqlStr & " ,d.suplyprice , d.itemno ,u.OnLineUSerID , d.cancelyn , u.UserName ,u.UserSeq" + vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shopjumun_master m" + vbcrlf
		sqlStr = sqlStr & " join db_shop.dbo.tbl_shopjumun_detail d" + vbcrlf
		sqlStr = sqlStr & " on m.idx = d.masteridx" + vbcrlf
		sqlStr = sqlStr & " left Join [db_shop].[dbo].[tbl_total_shop_card] AS B " + vbcrlf
		sqlStr = sqlStr & " ON m.pointuserno = B.cardno" + vbcrlf
		sqlStr = sqlStr & " left JOIN [db_shop].[dbo].tbl_total_shop_user AS u" + vbcrlf
		sqlStr = sqlStr & " ON b.UserSeq = u.UserSeq" + vbcrlf
		sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_user s" + vbcrlf
		sqlStr = sqlStr & " on m.shopid = s.userid" + vbcrlf				
		sqlStr = sqlStr & " where m.idx <>0 " & sqlsearch
		sqlStr = sqlStr & " order by m.orderno Desc" + vbcrlf

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.recordcount

		redim preserve FItemList(FTotalCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cTotalPoint_item

				FItemList(i).fshopname = db2html(rsget("shopname"))
				FItemList(i).fUserSeq = db2html(rsget("UserSeq"))
				FItemList(i).fUserName = db2html(rsget("UserName"))
				FItemList(i).fcancelyn = rsget("cancelyn")
				FItemList(i).fOnLineUSerID = rsget("OnLineUSerID")
				FItemList(i).forderno = rsget("orderno")
				FItemList(i).fshopid = rsget("shopid")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fpointuserno = rsget("pointuserno")
				FItemList(i).fitemgubun = rsget("itemgubun")
				FItemList(i).fitemid = rsget("itemid")
				FItemList(i).fitemoption = rsget("itemoption")
				FItemList(i).fitemname = db2html(rsget("itemname"))	
				FItemList(i).fitemoptionname = db2html(rsget("itemoptionname"))
				FItemList(i).fsellprice = rsget("sellprice")
				FItemList(i).frealsellprice = rsget("realsellprice")
				FItemList(i).fsuplyprice = rsget("suplyprice")
				FItemList(i).fitemno = rsget("itemno")
												
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//admin/totalpoint/index.asp
	public sub GetTotalPointList()
		dim strSql, i , strSubSql

		If FUserName <> "" Then
			strSubSql = strSubSql & " AND A.UserName = '" & FUserName & "' "
		End If
		
		If FCardNo <> "" Then
			strSubSql = strSubSql & " AND B.CardNo = '" & FCardNo & "' "
		End If

		If FUserID <> "" Then
			strSubSql = strSubSql & " AND A.OnlineUserID = '" & FUserID & "' "
		End If

		If FUseYN <> "" Then
			strSubSql = strSubSql & " AND B.UseYN = '" & FUseYN & "' "
		End If
		
		IF FCardGubun <> "" Then
			If FCardGubun = "T" Then
				strSubSql = strSubSql & " AND Left(B.CardNo,4) NOT IN('1010','3253') "
			Else
				strSubSql = strSubSql & " AND Left(B.CardNo,4) = '" & FCardGubun & "' "
			End IF
		End IF

		If frectshopid <> "" Then
			strSubSql = strSubSql & " AND b.regShopid = '" & frectshopid & "'"
		End If

		if FRectStartDay<>"" then
			strSubSql = strSubSql + " and b.regdate>='" + CStr(FRectStartDay) + "'"
		end if

		if FRectEndDay<>"" then
			strSubSql = strSubSql + " and b.regdate<'" + CStr(FRectEndDay) + "'"
		end if

		If frectmemberYn = "Y" Then
			strSubSql = strSubSql & " AND B.userseq <> 0 "
		elseIf frectmemberYn = "N" Then
			strSubSql = strSubSql & " AND B.userseq = 0 "
		End If

		If frectuserhp <> "" Then
			strSubSql = strSubSql & " AND replace(a.hpno,'-','') = '" & replace(frectuserhp,"-","") & "'"
		End If

		'총 갯수 구하기
		strSql = "select count(*) as cnt"
		strSql = strSql & " FROM [db_shop].[dbo].[tbl_total_shop_card] AS B"
		strSql = strSql & " Left Join [db_shop].[dbo].[tbl_total_shop_user] AS A"
		strSql = strSql & " 	ON b.UserSeq = a.UserSeq"
		strSql = strSql & " Left Join [db_shop].[dbo].[tbl_shop_user] AS D"
		strSql = strSql & " 	ON a.RegShopID = D.userid"
		strSql = strSql & " WHERE 1=1 " & strSubSql			

		'response.write strSql &"<br>"			
		rsget.Open strSql,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트 
		strSql = "select top " & Cstr(FPageSize * FCurrPage)
		strSql = strSql & " A.UserName, A.Jumin1, A.Jumin2_Enc, B.Regdate, isNull(D.shopname,'') AS shopname"
		strSql = strSql & " , A.Grade, B.CardNo, A.OnlineUserID, isNull(B.Point,0) AS Point, B.UseYN, b.userseq"
		strSql = strSql & " FROM [db_shop].[dbo].[tbl_total_shop_card] AS B"
		strSql = strSql & " Left Join [db_shop].[dbo].[tbl_total_shop_user] AS A"
		strSql = strSql & " 	ON b.UserSeq = a.UserSeq"
		strSql = strSql & " Left Join [db_shop].[dbo].[tbl_shop_user] AS D"
		strSql = strSql & " 	ON a.RegShopID = D.userid"
		strSql = strSql & " WHERE 1=1 " & strSubSql
		strSql = strSql & " ORDER BY A.UserSeq DESC"

		'response.write strSql &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open strSql,dbget,1

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
				set FItemList(i) = new cTotalPoint_item

				FItemList(i).fUserSeq = rsget("UserSeq")
				FItemList(i).fUserName = rsget("UserName")
				FItemList(i).fJumin1 = rsget("Jumin1")
				FItemList(i).fJumin2_Enc = rsget("Jumin2_Enc")
				FItemList(i).fRegdate = rsget("Regdate")
				FItemList(i).fshopname = rsget("shopname")
				FItemList(i).fGrade = rsget("Grade")
				FItemList(i).fCardNo = rsget("CardNo")
				FItemList(i).fOnlineUserID = rsget("OnlineUserID")
				FItemList(i).fPoint = rsget("Point")
				FItemList(i).fUseYN = rsget("UseYN")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'//admin/totalpoint/customer_sell_history_point.asp
	public Function GetTotalPointDetail()
		Dim strSql

		strSql = "SELECT COUNT(A.UserSeq) From " & _
				" [db_shop].[dbo].[tbl_total_shop_user] AS A " & _
				" Left Join [db_shop].[dbo].[tbl_total_shop_card] AS B ON A.UserSeq = B.UserSeq " & _
				" Left Join [db_user].[dbo].[tbl_user_n] AS C ON A.OnlineUserID = C.userid " & _
				" Left Join [db_shop].[dbo].[tbl_shop_user] AS D ON A.RegShopID = D.userid " & _
				" WHERE A.UserSeq = '" & FUserSeq & "' AND B.UseYN = 'Y' "

		'response.write strSql & "<br>"
		rsget.Open strSql,dbget
		IF not rsget.EOF THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
			
		IF FTotCnt > 0 THEN
			strSql = "SELECT " & _
					" A.UserSeq, A.UserName, A.Jumin1, A.Jumin2_Enc, A.Grade, C.SexFlag, " & _
					" isNull(C.userphone,'--') AS TelNo, isNull(C.usercell,'--') AS HpNo, " & _
					" C.ZipCode, " & _
					" C.Zipaddr AS useraddr1, " & _
					" C.useraddr AS useraddr2, C.usermail, A.EmailYN, A.SMSYN, A.LastUpdate, A.Regdate, D.shopname, B.Point, B.CardNo " & _
					" FROM [db_shop].[dbo].[tbl_total_shop_user] AS A " & _
					" Left Join [db_shop].[dbo].[tbl_total_shop_card] AS B ON A.UserSeq = B.UserSeq " & _
					" Left Join [db_user].[dbo].[tbl_user_n] AS C ON A.OnlineUserID = C.userid " & _
					" Left Join [db_shop].[dbo].[tbl_shop_user] AS D ON A.RegShopID = D.userid " & _
					" WHERE A.UserSeq = '" & FUserSeq & "' AND B.UseYN = 'Y' "

			'response.write strSql & "<br>"
			rsget.Open strSql,dbget,1

			IF not rsget.EOF THEN
				FUserName		= rsget("UserName")
				FJumin1			= rsget("Jumin1")
				FJumin2_Enc		= rsget("Jumin2_Enc")
				FCardNo			= rsget("CardNo")
				FPoint			= rsget("Point")
				FGrade			= rsget("Grade")
				FSexFlag		= rsget("SexFlag")
				FTelNo			= rsget("TelNo")
				FHpNo			= rsget("HpNo")
				FZipCode		= rsget("ZipCode")
				FAddress		= rsget("useraddr1")
				FAddressDetail	= rsget("useraddr2")
				FEmail			= rsget("usermail")
				FEmailYN		= rsget("EmailYN")
				FSMSYN			= rsget("SMSYN")
				FLastUpdate		= rsget("LastUpdate")
				FRegdate		= rsget("Regdate")
				FShopName		= rsget("shopname")
			END IF
			rsget.close
		END IF

	End Function

	'####### 로그리스트 #######
	public Function GetTotalPointLogList()
		Dim strSql, subSql

		IF FCardNo <> "" Then
			subSql = " AND B.CardNo = '" & FCardNo & "' "
		End If

		strSql = "	SELECT " & _
				"			C.CardNo, C.Point, D.code_desc, C.RegShopID, isNull(C.LogDesc,'') AS LogDesc, isNull(C.OrderNO,'') AS OrderNO, C.CasherID, C.Regdate, C.PointCode " & _
				"		FROM [db_shop].[dbo].[tbl_total_shop_user] AS A " & _
				"			Left Join [db_shop].[dbo].[tbl_total_shop_card] AS B ON A.UserSeq = B.UserSeq " & _
				"			Left Join [db_shop].[dbo].[tbl_total_shop_log] AS C ON B.CardNo = C.CardNo " & _
				"			Left Join [db_shop].[dbo].[tbl_total_shop_code] AS D ON C.PointCode = D.code_value " & _
				"	WHERE A.UserSeq = '" & FUserSeq & "' " & subSql & " ORDER BY C.Regdate DESC "
		rsget.Open strSql,dbget,1
		'response.write strSql
		
		IF not rsget.EOF THEN
			GetTotalPointLogList = rsget.getRows() 
		END IF	
		rsget.close

		strSql = "	SELECT " & _
				"			isNull(SUM(Point),0) " & _
				"		FROM [db_shop].[dbo].[tbl_total_shop_card] " & _
				"	WHERE UserSeq = '" & FUserSeq & "' "
		rsget.Open strSql,dbget,1
		IF not rsget.EOF THEN
			FTotalPoint = rsget(0)
		END IF	
		rsget.close

	End Function
	
	'/admin/totalpoint/customer_sell_history_point.asp
	public Function GetMemberCardList()
		Dim strSql
		
		strSql = "	SELECT DISTINCT " & _
				"			C.CardNo " & _
				"		FROM [db_shop].[dbo].[tbl_total_shop_card] AS A " & _
				"			Left Join [db_shop].[dbo].[tbl_total_shop_log] AS C ON A.CardNo = C.CardNo " & _
				"	WHERE A.UserSeq = '" & FUserSeq & "' "
		rsget.Open strSql,dbget,1
		'response.write strSql
		
		IF not rsget.EOF THEN
			GetMemberCardList = rsget.getRows() 
		END IF	
		rsget.close

	End Function

End Class

Public Function UserStatus(vStatus)
	IF vStatus = "0" Then
		UserStatus = "등록대기"
	ElseIf vStatus = "1" Then
		UserStatus = "정상회원"
	ElseIf vStatus = "9" Then
		UserStatus = "탈퇴회원"
	Else
		UserStatus = ""
	End If
End Function
%>