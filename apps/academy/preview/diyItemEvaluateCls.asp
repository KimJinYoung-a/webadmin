<%
'####################################################
' Description :  DIY item 상품후기 클래스
' History : 2010.10.25 허진원 생성
'			2010.11.01 한용민 수정
'####################################################

class CEvaluateSearcherItem
	Public Fidx
	public FUserID
	public FTitle
	public FUesdContents
	public FManiaPoint
	public FTotalPoint
	public FPoint
	public FPoint_fun
	public FPoint_dgn
	public FPoint_prc
	public FPoint_stf
	public Fimgsmall
	public FIcon1
	public FIcon2	
	public Flinkimg1
	public Flinkimg2
	public Flinkimg3
	public Flinkimg4
	public Flinkimg5		
	Public FImgContents1
	Public FImgContents2
	Public FImgContents3
	Public FImgContents4
	Public FImgContents5	
	public FItemID
	public Fimglist
	Public Fgubun
	public FRegdate
	Public FItemname
	Public FItemCost
	Public FItemDiv
	Public FItemOption
	Public FOptionName
	Public FMakerName
	Public FMakerID
	Public FOrderSerial
	Public FOrderDate
	Public FImageList100
	Public FImageList120
	Public FEvalRegDate
	Public FEvalCnt	
	Public F100ShopIdx
	Public FCouponName
	Public FCouponValue
	Public FCouponType
	Public FCouponStartDate
	Public FCouponExpireDate
	Public Fminbuyprice
	Public Fhitcount
	Public Fcommentcount
	Public Fscoresum
	Public Fsellcash
	Public Fcontents
	Public Fnourlfile1
	Public Ffile1
	Public Fnourlfile2
	Public Ffile2
	Public Fnourlfile3
	Public Ffile3
	Public Fnourlfile4
	Public Ffile4
	Public Fnourlfile5
	Public Ffile5
	Public Fnourlicon1
	Public FstartDate
	Public FendDate	
	Public FUseGood
	Public FUseBad
	Public FUseETC
	Public FMyBlog

	public Function getUsingTitle(LimitSize)
	
		if Len(FUesdContents) > LimitSize then
			getUsingTitle = Left(FUesdContents,LimitSize) + "..."
		else
			getUsingTitle = FUesdContents
		end if
	
	end Function 
	
	public function IsPhotoExist()
		IsPhotoExist = (Flinkimg1<>"") or (Flinkimg2<>"")
	end function
	
	public Function getLinkImage1()					
		getLinkImage1 = fingersImgUrl + "/contents/academy_diyevaluate/" + GetImageSubFolderByItemid(Fitemid) + "/" + Flinkimg1		
	end function 
	
	public Function getLinkImage2()
		getLinkImage2 = fingersImgUrl + "/contents/academy_diyevaluate/" + GetImageSubFolderByItemid(Fitemid) + "/" + Flinkimg2
	end function 
	
	public Function getLinkImage3()
		getLinkImage3 = fingersImgUrl + "/contents/academy_diyevaluate/" + GetImageSubFolderByItemid(Fitemid) + "/" + Flinkimg3
	end function 
	
	public Function getLinkImage4()
		getLinkImage4 = fingersImgUrl + "/contents/academy_diyevaluate/" + GetImageSubFolderByItemid(Fitemid) + "/" + Flinkimg4
	end function 
	
	public Function getLinkImage5()
		getLinkImage5 = fingersImgUrl + "/contents/academy_diyevaluate/" + GetImageSubFolderByItemid(Fitemid) + "/" + Flinkimg5
	end function 
	
	public Function getIconImage1()
		if Fgubun="0" then
			getIconImage1 =	"http://imgstatic.10x10.co.kr/goodsimage/" + GetImageSubFolderByItemid(Fitemid) + "/" + FIcon1
		else
			getIconImage1 = "http://imgstatic.10x10.co.kr/contents/maniaimg/evaluate/" & CStr(Fgubun) & "/icon1/" + FIcon1
		end if
	end function 

	public Function getIconImage2()
		if Fgubun="0" then
			getIconImage2 =	"http://imgstatic.10x10.co.kr/goodsimage/" + GetImageSubFolderByItemid(Fitemid) + "/" + FIcon2
		else
			getIconImage2 = "http://imgstatic.10x10.co.kr/contents/maniaimg/evaluate/" & CStr(Fgubun) & "/icon2/" + FIcon2
		end if
	end function 

	Private Sub Class_Terminate()
	End Sub
	public sub Class_Initialize()
	end sub
end Class

Class CEvaluateSearcher
	public FItemList()
	public FcdLCnt()
	public FcdLTotalPage
	public FEvalItem
	public FTotTotalCount
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FIdx
	public FRectUserID
	public FRectItemID
	public FECode	'이벤트코드
	public FDiscountRate
	public FRectStartPoint
	public FSortMethod
	public FRectcdL
	public FRectEvaluatedYN
	public FRectOrderSerial
	public FRectOption
	public FRectSearchtype
	public FRectsearchrect

    public FRectDeledEvalInclude
    
	Private Sub Class_Initialize()
		redim preserve FItemList(0)

		FCurrPage     = 1
		FPageSize     = 5
		FResultCount  = 0
		FScrollCount  = 10
		FTotalCount   = 0

		FDiscountRate = 1
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public sub getItemEvalList()
		dim sqlStr,i

		sqlStr = "exec [db_academy].[dbo].sp_academy_Evaluate_Tcnt '" & CStr(FPageSize) & "','" + Cstr(FRectItemID) + "','" + Cstr(FRectStartPoint) + "','" + Cstr(Fidx) + "','" + Cstr(FsortMethod)+ "'" + vbcrlf
		
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType = adOpenStatic
		rsACADEMYget.LockType = adLockOptimistic
		rsACADEMYget.Open sqlStr,dbACADEMYget

		FTotalCount = rsACADEMYget("TotalCnt")
		FTotalPage =  rsACADEMYget("TotalPage")
		rsACADEMYget.close

		sqlStr = "exec [db_academy].[dbo].sp_academy_Evaluate '" +  CStr(FPageSize) + "','" + CStr(FCurrPage) + "','" + Cstr(FRectItemID) + "','" + Cstr(FRectStartPoint) + "','" + Cstr(Fidx) + "','" + Cstr(FsortMethod) + "'" + vbcrlf

		'Response.write sqlStr
		
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType = adOpenStatic
		rsACADEMYget.LockType = adLockOptimistic
		rsACADEMYget.Open sqlStr,dbACADEMYget

		FResultCount = rsACADEMYget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			do until rsACADEMYget.eof
				set FItemList(i) = new CEvaluateSearcherItem
									
				FItemList(i).Fidx			= rsACADEMYget("idx")
				FItemList(i).Fgubun			= rsACADEMYget("Gubun")
				FItemList(i).FUserID		= rsACADEMYget("UserID")
				FItemList(i).FItemID		= rsACADEMYget("ItemID")
				FItemList(i).FTotalPoint	= rsACADEMYget("TotalPoint")
				FItemList(i).FUesdContents 	= db2html(rsACADEMYget("contents"))
				FItemList(i).FPoint_fun		= rsACADEMYget("Point_Function")
				FItemList(i).FPoint_dgn		= rsACADEMYget("Point_Design")
				FItemList(i).FPoint_prc		= rsACADEMYget("Point_Price")
				FItemList(i).FPoint_stf		= rsACADEMYget("Point_Satisfy")
				FItemList(i).FRegdate 		= rsACADEMYget("RegDate")
				FItemList(i).Flinkimg1		= rsACADEMYget("file1")
				FItemList(i).Flinkimg2		= rsACADEMYget("file2")
				
				FItemList(i).FOrderSerial		= rsACADEMYget("OrderSerial")
				FItemList(i).FItemOption		= rsACADEMYget("ItemOption")
				
				FItemList(i).FOptionName	= rsACADEMYget("itemoptionname")

				'// 과거자료 중 0점이 존재 1점으로 표시
				if FItemList(i).FTotalPoint="0" then FItemList(i).FTotalPoint="1"
				if FItemList(i).FPoint_fun="0" then FItemList(i).FPoint_fun="1"
				if FItemList(i).FPoint_dgn="0" then FItemList(i).FPoint_dgn="1"
				if FItemList(i).FPoint_prc="0" then FItemList(i).FPoint_prc="1"
				if FItemList(i).FPoint_stf="0" then FItemList(i).FPoint_stf="1"
				
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
	end sub
	
	public sub getItemEvalOne()
		dim sqlStr,i

		sqlStr = "exec [db_board].[dbo].sp_Ten_Evaluate '1','1','" + Cstr(FRectItemID) + "','','" + Cstr(FIdx) + "',''" + vbcrlf
				
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType = adOpenStatic
		rsACADEMYget.LockType = adLockOptimistic
		rsACADEMYget.Open sqlStr,dbACADEMYget

		FResultCount = rsACADEMYget.RecordCount

		i=0
		set FEvalItem = new CEvaluateSearcherItem
		if  not rsACADEMYget.EOF  then
			FEvalItem.Fidx			= rsACADEMYget("idx")
			FEvalItem.Fgubun		= rsACADEMYget("Gubun")
			FEvalItem.FUserID		= rsACADEMYget("UserID")
			FEvalItem.FItemID		= rsACADEMYget("ItemID")
			FEvalItem.FTotalPoint	= rsACADEMYget("TotalPoint")
			FEvalItem.FUesdContents = db2html(rsACADEMYget("contents"))
			FEvalItem.FPoint_fun	= rsACADEMYget("Point_Function")
			FEvalItem.FPoint_dgn	= rsACADEMYget("Point_Design")
			FEvalItem.FPoint_prc	= rsACADEMYget("Point_Price")
			FEvalItem.FPoint_stf	= rsACADEMYget("Point_Satisfy")
			FEvalItem.FRegdate 		= rsACADEMYget("RegDate")
			FEvalItem.Flinkimg1		= rsACADEMYget("file1")
			FEvalItem.Flinkimg2		= rsACADEMYget("file2")
			FEvalItem.FOptionName	= rsACADEMYget("itemoptionname")
			
		end if

		rsACADEMYget.Close
	end sub

	'// 후기쓴 상품 리스트 '//myfingers/goodsusing/diyitem/diyitem_goodsusing.asp
	Public Sub EvalutedItemList()		
		dim sqlStr,i
			sqlStr = "" &_
				" select Count(e.idx) as TotalCnt , Ceiling(cast(count(e.idx) as Float)/" & Cstr(FPageSize) & ") as TotalPage " &_
				" FROM db_academy.dbo.tbl_diy_item_Evaluate e "&_
				" JOIN db_academy.dbo.tbl_diy_item i "&_
				" on e.itemid=i.itemid "&_
				" WHERE userid='" & FRectUserID & "' "&_
				" and e.isusing='Y' " 

'				if FRectcdL<>"" then 
'					sqlStr = sqlStr & " and i.cate_large='" & FRectcdL & "'"
'				end if
				
				'response.write sqlStr &"<br>"				
				rsACADEMYget.open sqlStr ,dbACADEMYget,1
				
				IF not rsACADEMYget.eof THEN 
					FTotalCount = rsACADEMYget("TotalCnt")
					FTotalPage =  rsACADEMYget("TotalPage")
				End if
				
				rsACADEMYget.close	
								
			sqlStr = " " &_
				" SELECT Top " & Cstr(FPageSize*(FCurrPage)) &_
				"   e.idx , e.gubun , e.contents ,  e.regdate , e.orderserial,e.itemoption, e.userid " &_
				" , e.file1 , e.file2 , e.file3 ,e.file4 , e.file5 "&_
				" , isnull(e.TotalPoint,0) as TotalPoint "&_
				" , isnull(e.Point_function,0) as Point_function "&_
				" , isnull(e.Point_Design,0) as Point_Design "&_
				" , isnull(e.Point_Price,0) as Point_Price "&_ 
				" , isnull(e.Point_satisfy,0) as Point_satisfy "&_
				" , i.itemid , i.itemname , i.sellcash , i.makerID , i.brandname , i.listimage120 , i.listimage , i.itemdiv  "&_
				" , o.optionname "&_
				" FROM db_academy.dbo.tbl_diy_item_Evaluate e "&_
				" JOIN db_academy.dbo.tbl_diy_item i "&_
				" on e.itemid=i.itemid "&_
				" LEFT JOIN db_academy.dbo.tbl_diy_item_option o"&_
				" on e.itemid=o.itemid and e.itemoption = o.itemoption  "&_
				" WHERE userid='" & FRectUserID & "' "&_
				" and e.isusing='Y' " 

'				if FRectcdL<>"" then
'					sqlStr = sqlStr & " and i.cate_large='" & FRectcdL & "'"
'				end if

'				Select Case FSortMethod
'					
'					case "Best" '//베스트 상품순 많이 -- 인기 상품 우선
'						sqlStr = sqlStr & " ORDER by i.itemscore desc, i.itemid desc "
'					case "Buy"	'//구매 일자 순 -- 주문 번호 내림차순
'						sqlStr = sqlStr & " ORDER by e.orderserial desc "
'					case "Reg"	'//작성 일자순 -- 후기 작성 일자,상품 번호 
'						sqlStr = sqlStr & " ORDER by e.regdate desc,i.itemid desc "
'					case "Photo"'//포토 상품 후기순 -- 이미지 있는것 먼저,상품 번호 내림차순
'						sqlStr = sqlStr & " ORDER by e.file1 desc, e.orderserial desc ,e.itemid  "
'				end Select 
			
			'response.write sqlStr &"<br>"
			rsACADEMYget.pagesize = FPageSize
			rsACADEMYget.open sqlStr ,dbACADEMYget,1
			
			FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
			
			redim preserve FItemList(FResultCount) 
			i=0 
			
			IF not rsACADEMYget.eof THEN 
				rsACADEMYget.absolutepage = FCurrPage
				do until rsACADEMYget.eof 
					
					set FItemList(i) = new CEvaluateSearcherItem
					
					FItemList(i).FItemID 			= rsACADEMYget("itemid")
					FItemList(i).Fuserid 			= rsACADEMYget("userid")
					FItemList(i).FItemname 			= db2html(rsACADEMYget("itemname"))
					FItemList(i).FItemCost			= rsACADEMYget("sellcash")
					FItemList(i).FOptionName 		= db2html(rsACADEMYget("optionname"))
					FItemList(i).FItemDiv			= rsACADEMYget("itemdiv")
					FItemList(i).FMakerName			= db2html(rsACADEMYget("brandname"))
					FItemList(i).FMakerID			= rsACADEMYget("makerID")
					FItemList(i).FOrderSerial 		= rsACADEMYget("orderserial")
					FItemList(i).FItemOption 		= rsACADEMYget("itemoption")
					FItemList(i).FOrderDate 		= rsACADEMYget("regdate")
					FItemList(i).FImageList100 	= fingersImgUrl & "/diyitem/webimage/list/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("listimage")
					FItemList(i).FImageList120 	= fingersImgUrl & "/diyitem/webimage/list120/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("listimage120")					
					FItemList(i).Fidx					= rsACADEMYget("idx")
					FItemList(i).Fgubun				= rsACADEMYget("Gubun")					
					FItemList(i).FTotalPoint		= rsACADEMYget("TotalPoint")
					FItemList(i).FUesdContents 	= db2html(rsACADEMYget("contents"))
					FItemList(i).FPoint_fun			= rsACADEMYget("Point_Function")
					FItemList(i).FPoint_dgn			= rsACADEMYget("Point_Design")
					FItemList(i).FPoint_prc			= rsACADEMYget("Point_Price")
					FItemList(i).FPoint_stf			= rsACADEMYget("Point_Satisfy")					
					FItemList(i).Flinkimg1			= rsACADEMYget("file1")
					FItemList(i).Flinkimg2			= rsACADEMYget("file2")
					FItemList(i).Flinkimg3			= rsACADEMYget("file3")
					FItemList(i).Flinkimg4			= rsACADEMYget("file4")
					FItemList(i).Flinkimg5			= rsACADEMYget("file5")					
					FItemList(i).FRegDate		= rsACADEMYget("regdate")

					if FItemList(i).Flinkimg1<>"" then
						FItemList(i).Flinkimg1 = fingersImgUrl  + "/contents/academy_diyevaluate/"  + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + FItemList(i).Flinkimg1
					' 	FItemList(i).Flinkimg1 = "http://imgstatic.10x10.co.kr/contents/academy_evaluate/" + FItemList(i).Flinkimg1
					end if

					if FItemList(i).Flinkimg2<>"" then
						FItemList(i).Flinkimg2 = fingersImgUrl  + "/contents/academy_diyevaluate/"  + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + FItemList(i).Flinkimg2
					'	FItemList(i).Flinkimg2 = "http://imgstatic.10x10.co.kr/contents/academy_evaluate/" + FItemList(i).Flinkimg2
					end if

					i=i+1
					rsACADEMYget.movenext
				loop 
			END IF
			
			rsACADEMYget.close
	
	End Sub
	
	'// 최근 3개월 이내 구매 & 후기 안쓰인 상품 리스트 '/myfingers/goodsusing/diyitem/diyitem_goodsusing.asp
	Public Sub NotEvalutedItemList()		
		dim sqlStr ,i

		sqlStr = "" &_
				" select Count(m.orderserial) as TotalCnt , Ceiling(cast(count(m.orderserial) as Float)/" & Cstr(FPageSize) & ") as TotalPage " &_
				" FROM [db_academy].dbo.tbl_academy_order_master m  "&_
				" JOIN [db_academy].dbo.tbl_academy_order_detail d  "&_
				" on m.OrderSerial= d.OrderSerial and m.sitename='diyitem' and m.ipkumdiv>=7  "&_
				" and m.cancelyn='N' and m.jumundiv<>9 and d.cancelyn<>'Y'  "&_
				" and d.itemid<>0  "&_
				" JOIN [db_academy].dbo.tbl_diy_item i "&_
				" on d.itemid=i.itemid "&_
				" LEFT JOIN db_academy.dbo.tbl_diy_item_Evaluate e  "&_
				" on e.UserID='" & FRectUserID & "' and m.OrderSerial = e.OrderSerial and d.Itemid=e.itemid and d.ItemOption = e.ItemOption   "&_
				" WHERE e.idx is null " &_
				" and m.userid='" & FRectUserID & "'  "
				
				if FRectcdL<>"" then 
					sqlStr = sqlStr & " and i.cate_large='" & FRectcdL & "'"
				end if
				
				' 3개월 제한 - 2007/07월 이후  " m.regdate > dateadd(month,-3,convert(varchar(10),getdate(),121)) "
				
				'response.write sqlStr &"<br>"
				rsACADEMYget.open sqlStr ,dbACADEMYget,1
				
				IF not rsACADEMYget.eof THEN 
					FTotalCount = rsACADEMYget("TotalCnt")
					FTotalPage =  rsACADEMYget("TotalPage")
				End if
				rsACADEMYget.close	

		sqlStr = " " &_
				" SELECT TOP " & Cstr(FPageSize*(FCurrPage)) &_ 
				"  i.itemid , i.sellcash , i.itemname , i.brandname , i.makerid , i.listimage120, i.listimage , i.itemdiv, i.evalcnt "&_
				" , d.itemoption , o.optionname "&_
				" , m.orderserial ,m.regdate" &_
				" FROM [db_academy].dbo.tbl_academy_order_master m  "&_
				" JOIN [db_academy].dbo.tbl_academy_order_detail d  "&_
				" on m.OrderSerial= d.OrderSerial and m.sitename='diyitem' and m.ipkumdiv>=7  "&_
				" and m.cancelyn='N' and m.jumundiv<>9 and d.cancelyn<>'Y'  "&_
				" and d.itemid<>0  "&_
				" JOIN [db_academy].dbo.tbl_diy_item i  "&_
				" 	on d.itemid=i.itemid  "&_
				" LEFT JOIN db_academy.dbo.tbl_diy_item_option o "&_
				" on d.itemid = o.itemid and d.itemoption = o.itemoption "&_
				" LEFT JOIN db_academy.dbo.tbl_diy_item_Evaluate e  "&_
				" on e.UserID='" & FRectUserID & "' and m.OrderSerial = e.OrderSerial and d.Itemid=e.itemid and d.ItemOption = e.ItemOption  "&_
				" WHERE e.idx is null " &_
				" and m.userid='" & FRectUserID & "'  "
				
				if FRectcdL<>"" then 
					sqlStr = sqlStr & " and i.cate_large='" & FRectcdL & "'"
				end if
				
				' 3개월 제한 - 2007/07월 이후 제한  " m.regdate > dateadd(month,-3,convert(varchar(10),getdate(),121)) "
				
				Select Case FSortMethod
				
				case "Best" '//베스트 상품순 -- 인기 상품 우선
					sqlStr = sqlStr & " ORDER by i.itemscore desc, i.itemid desc "
				case "Buy"	'//구매 일자 순 -- 주문 번호 내림차순
					sqlStr = sqlStr & " ORDER by m.orderserial desc "
				case "Reg"	'//작성 유효 일자순 -- 주문 번호 올림차순
					sqlStr = sqlStr & " ORDER by m.orderserial,i.itemid desc "
				case "Photo"'//포토 상품 후기순 -- 이미지 있는것 먼저,상품 번호 내림차순
					sqlStr = sqlStr & " ORDER by e.file1 desc, e.orderserial desc  "
			end Select 
		
			'response.write sqlStr &"<br>"
			rsACADEMYget.pagesize = FPageSize
			rsACADEMYget.open sqlStr ,dbACADEMYget,1
			
			FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
			if (FResultCount<1) then FResultCount=0
			redim preserve FItemList(FResultCount) 
			i=0 
			
			IF not rsACADEMYget.eof THEN 
				rsACADEMYget.absolutepage = FCurrPage
				do until rsACADEMYget.eof 
					
					set FItemList(i) = new CEvaluateSearcherItem
					
					FItemList(i).FItemID 			= rsACADEMYget("itemid")
					FItemList(i).FItemname 			= db2html(rsACADEMYget("itemname"))
					FItemList(i).FItemCost			= rsACADEMYget("sellcash")
					FItemList(i).FItemOption 		= rsACADEMYget("itemoption")
					FItemList(i).FOptionName 		= db2html(rsACADEMYget("optionname"))
					FItemList(i).FItemDiv			= rsACADEMYget("itemdiv")
					FItemList(i).FMakerName			= db2html(rsACADEMYget("brandname"))
					FItemList(i).FMakerID			= rsACADEMYget("makerID")
					FItemList(i).FOrderSerial 		= rsACADEMYget("orderserial")
					FItemList(i).FOrderDate 		= rsACADEMYget("regdate")				
					FItemList(i).FImageList100 	= fingersImgUrl & "/diyitem/webimage/list/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("listimage")
					FItemList(i).FImageList120 	= fingersImgUrl & "/diyitem/webimage/list120/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("listimage120")
					FItemList(i).FRegDate		= rsACADEMYget("regdate")
					FItemList(i).FEvalCnt			= rsACADEMYget("evalcnt")
					
					i=i+1
					rsACADEMYget.movenext
				loop 
			END IF
			
			rsACADEMYget.close
				
	End Sub

	'// 후기 안쓴 상품 '//myfingers/goodsusing/diyitem/diyitem_goodsUsingWrite.asp
	Public Sub getNotEvaluatedItem()
		dim sqlStr
		
		sqlStr = " " &_
			" SELECT top 1 " &_
			"  d.orderserial, d.itemid,d.itemoption  " &_
			" ,i.itemid,i.sellcash,i.itemname,i.brandname , i.listimage , i.listImage120  " &_
			" ,o.optionname " &_
			" , e.idx ,e.gubun,  e.TotalPoint , e.Point_Function , e.Point_Design , e.Point_Price , e.Point_Satisfy " &_
			" , e.icon1, e.icon2 , e.file1 , e.file2 , e.file3 , e.file4 , e.file5" &_ 
			" , e.Contents as UsedContents , e.imgcontents1 , e.imgcontents2 , e.imgcontents3 , e.imgcontents4 , e.imgcontents5 " &_			
			" FROM db_academy.dbo.tbl_academy_order_master m  " &_
			" JOIN db_academy.dbo.tbl_academy_order_detail d  " &_
			" on m.OrderSerial= d.OrderSerial and m.sitename='diyitem' and m.ipkumdiv>=7 and m.cancelyn='N' and m.jumundiv<>'9' and d.cancelyn<>'Y' and d.itemid<>'0'  " &_
			" JOIN db_academy.dbo.tbl_diy_item i  " &_
			" on d.itemid=i.itemid  " &_
			" LEFT JOIN db_academy.dbo.tbl_diy_item_option o  " &_
			" on d.itemid=o.itemid and d.itemoption = o.itemoption  " &_
			" left join db_academy.dbo.tbl_diy_item_Evaluate e " &_
			" on e.orderserial = d.orderserial and e.itemid=d.itemid and e.itemoption = d.itemoption and m.userid= e.userid and e.isusing='Y' " &_
			" WHERE d.itemid='" & FRectItemID & "'  " &_
			" and m.userid='" & FRectUserID & "'  " &_
			" and m.OrderSerial='" & FRectOrderSerial & "'  " &_
			" and d.itemoption ='" & FRectOption & "'  "
		
			' 3개월 제한 - 2007/07월 이후 제한  " m.regdate > dateadd(month,-3,convert(varchar(10),getdate(),121)) "
			'"		and s.couponstartdate<=getdate() and s.couponexpiredate>getdate()  " &_
			
			'response.write sqlStr &"<Br>"			
			rsACADEMYget.open sqlStr ,dbACADEMYget,1
			
			FResultCount = rsACADEMYget.RecordCount
			
			set FEvalItem = new CEvaluateSearcherItem
			IF not rsACADEMYget.eof THEN 
					
				FEvalItem.FItemID 			= rsACADEMYget("itemid")
				FEvalItem.FItemname 			= db2html(rsACADEMYget("itemname"))
				FEvalItem.FItemCost			= rsACADEMYget("sellcash")
				FEvalItem.FItemOption 		= rsACADEMYget("itemoption")
				FEvalItem.FOptionName = db2html(rsACADEMYget("optionname"))
				FEvalItem.FMakerName			= db2html(rsACADEMYget("BrandName"))
				FEvalItem.FImageList100 	= fingersImgUrl & "/diyitem/webimage/list/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("listimage")
				FEvalItem.FImageList120 	= fingersImgUrl & "/diyitem/webimage/list120/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("listimage120")
				FEvalItem.FOrderSerial		= rsACADEMYget("orderserial")					
				FEvalItem.Fidx					= rsACADEMYget("idx")
				FEvalItem.Fgubun				= rsACADEMYget("Gubun")
				FEvalItem.FUesdContents 	= db2html(rsACADEMYget("usedcontents"))
				FEvalItem.FTotalPoint		= rsACADEMYget("TotalPoint")
				FEvalItem.FPoint_fun			= rsACADEMYget("Point_Function")
				FEvalItem.FPoint_dgn			= rsACADEMYget("Point_Design")
				FEvalItem.FPoint_prc			= rsACADEMYget("Point_Price")
				FEvalItem.FPoint_stf			= rsACADEMYget("Point_Satisfy")
				FEvalItem.Flinkimg1			= rsACADEMYget("file1")
				FEvalItem.Flinkimg2			= rsACADEMYget("file2")
				FEvalItem.Flinkimg3			= rsACADEMYget("file3")
				FEvalItem.Flinkimg4			= rsACADEMYget("file4")
				FEvalItem.Flinkimg5			= rsACADEMYget("file5")									
										
			END IF			
			rsACADEMYget.close	
	End Sub
	
	'// 후기 쓴 상품 '//myfingers/goodsusing/diyitem/diyitem_goodsUsingWrite.asp
	Public Sub getEvaluatedItem()
		dim sqlStr
		
		sqlStr = " " &vbCRLF
		sqlStr = sqlStr&" SELECT top 1 " &vbCRLF
		sqlStr = sqlStr&" e.idx , e.gubun , e.orderserial, e.itemoption , o.optionname  " &vbCRLF 
		sqlStr = sqlStr&" , e.TotalPoint , e.Point_Function , e.Point_Design , e.Point_Price , e.Point_Satisfy " &vbCRLF
		sqlStr = sqlStr&" , e.icon1, e.icon2 , e.file1 , e.file2 , e.file3 , e.file4 , e.file5" &vbCRLF
		sqlStr = sqlStr&" ,e.title , e.Contents as UsedContents , e.imgcontents1 , e.imgcontents2 , e.imgcontents3 , e.imgcontents4 , e.imgcontents5 " &vbCRLF
		sqlStr = sqlStr&" ,i.itemid,i.sellcash,i.itemname,i.brandname , i.listimage , i.listImage120, i.makerid " &vbCRLF
		sqlStr = sqlStr&" FROM db_academy.dbo.tbl_diy_item_Evaluate e " &vbCRLF
		sqlStr = sqlStr&" JOIN db_academy.dbo.tbl_diy_item i " &vbCRLF
		sqlStr = sqlStr&" on i.itemid= e.itemid " &vbCRLF
		sqlStr = sqlStr&" LEFT JOIN db_academy.dbo.tbl_diy_item_option o " &vbCRLF
		sqlStr = sqlStr&" on e.itemid=o.itemid and e.itemoption = o.itemoption " &vbCRLF
		sqlStr = sqlStr&" WHERE e.userid='" & CStr(userid) & "' " &vbCRLF
		sqlStr = sqlStr&" and e.itemid='" & CStr(itemid) & "' " &vbCRLF
		sqlStr = sqlStr&" and e.OrderSerial='" & CStr(FRectOrderSerial) & "' " &vbCRLF
		sqlStr = sqlStr&" and e.itemoption ='" & CStr(FRectOption) & "' "
		if (FRectDeledEvalInclude="ON") then
		    '' 상품후기 작성시 상품에서 넘어올경우 삭제 된 내역도 검사해야함. by eastone
		else
		    sqlStr = sqlStr&" and e.isusing='Y' "
	    end if


'response.write sqlStr &"<Br>"
'response.end
			rsACADEMYget.open sqlStr ,dbACADEMYget,1
			
			FResultCount = rsACADEMYget.RecordCount
			
			set FEvalItem = new CEvaluateSearcherItem
			IF not rsACADEMYget.eof THEN 
					
				FEvalItem.FItemID 			= rsACADEMYget("itemid")
				FEvalItem.FItemname 			= db2html(rsACADEMYget("itemname"))
				FEvalItem.FItemCost			= rsACADEMYget("sellcash")
				FEvalItem.FItemOption 		= rsACADEMYget("itemoption")
				FEvalItem.FOptionName = db2html(rsACADEMYget("optionname"))
				FEvalItem.FMakerName			= db2html(rsACADEMYget("BrandName"))
				FEvalItem.FImageList100 	= fingersImgUrl & "/diyitem/webimage/list/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("listimage")
				FEvalItem.FImageList120 	= fingersImgUrl & "/diyitem/webimage/list120/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("listimage120")
				FEvalItem.FOrderSerial		= rsACADEMYget("orderserial")					
				FEvalItem.Fidx					= rsACADEMYget("idx")
				FEvalItem.Fgubun				= rsACADEMYget("Gubun")
				FEvalItem.FTitle				= rsACADEMYget("title")
				FEvalItem.FUesdContents 	= db2html(rsACADEMYget("usedcontents"))					
				FEvalItem.FTotalPoint		= rsACADEMYget("TotalPoint")
				FEvalItem.FPoint_fun			= rsACADEMYget("Point_Function")
				FEvalItem.FPoint_dgn			= rsACADEMYget("Point_Design")
				FEvalItem.FPoint_prc			= rsACADEMYget("Point_Price")
				FEvalItem.FPoint_stf			= rsACADEMYget("Point_Satisfy")					
				FEvalItem.FIcon1				= rsACADEMYget("Icon1")					
				FEvalItem.Flinkimg1			= rsACADEMYget("file1")
				FEvalItem.Flinkimg2			= rsACADEMYget("file2")

				if FEvalItem.Flinkimg1<>"" then
					FEvalItem.Flinkimg1 = fingersImgUrl + "/contents/academy_diyevaluate/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + FEvalItem.Flinkimg1
				' 	FEvalItem.Flinkimg1 = "http://imgstatic.10x10.co.kr/contents/academy_evaluate/" + FEvalItem.Flinkimg1
				end if

				if FEvalItem.Flinkimg2<>"" then
					FEvalItem.Flinkimg2 = fingersImgUrl + "/contents/academy_diyevaluate/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + FEvalItem.Flinkimg2
				'	FEvalItem.Flinkimg2 = "http://imgstatic.10x10.co.kr/contents/academy_evaluate/" + FEvalItem.Flinkimg2
				end if

				FEvalItem.Flinkimg3			= rsACADEMYget("file3")
				FEvalItem.Flinkimg4			= rsACADEMYget("file4")
				FEvalItem.Flinkimg5			= rsACADEMYget("file5")					
				FEvalItem.FImgContents1			= rsACADEMYget("imgcontents1")
				FEvalItem.FImgContents2			= rsACADEMYget("imgcontents2")
				FEvalItem.FImgContents3			= rsACADEMYget("imgcontents3")
				FEvalItem.FImgContents4			= rsACADEMYget("imgcontents4")
				FEvalItem.FImgContents5			= rsACADEMYget("imgcontents5")
				FEvalItem.Fmakerid			= rsACADEMYget("makerid")
										
			END IF			
			rsACADEMYget.close	
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
	

	public Function getMyEvalListLightVer()
		dim sqlStr, arr

		sqlStr = " " &vbCRLF
		sqlStr = sqlStr&" SELECT d.orderserial, d.itemid, d.itemoption" &vbCRLF
		sqlStr = sqlStr&"  ,(select count(*) from db_academy.dbo.tbl_diy_item_Evaluate "&vbCRLF
		sqlStr = sqlStr&" where isusing='Y' and  orderserial=d.orderserial and itemid=d.itemid and itemoption=d.itemoption and userid='" & CStr(FRectUserID) & "') as evalcnt" &vbCRLF
		sqlStr = sqlStr&" FROM db_academy.dbo.tbl_academy_order_detail d  " &vbCRLF
		sqlStr = sqlStr&" JOIN db_academy.dbo.tbl_academy_order_master m " &vbCRLF
		sqlStr = sqlStr&" 	on m.orderserial=d.orderserial " &vbCRLF
		sqlStr = sqlStr&" WHERE m.sitename='diyitem' "
		sqlStr = sqlStr&" and d.itemid<>0 "
		sqlStr = sqlStr&" and d.cancelyn='N' "
		sqlStr = sqlStr&" and m.userid='" & CStr(FRectUserID) & "' " &vbCRLF
		sqlStr = sqlStr&" and d.itemid='" & CStr(FRectItemID) & "' "

'response.write sqlStr
'response.end
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly,adLockReadOnly
		if not rsACADEMYget.EOF then
			arr = rsACADEMYget.getRows()
		end if
		rsACADEMYget.close
		getMyEvalListLightVer = arr
	end Function
End Class

Function fnMyEvalCheck(arr, itemid)
	Dim i, torf, itoption, itorserial
	torf = False
	IF isArray(arr) THEN
		For i =0 To UBound(arr,2)
			If CStr(arr(1,i)) =  CStr(itemid) and CStr(arr(3,i)) =  0  Then	' 
				torf = True
				itorserial = CStr(arr(0,i))
				itoption = CStr(arr(2,i))
				Exit For
			End If
		Next
	End If
	fnMyEvalCheck = torf
End Function

Function fnMyEvalorderserial(arr, itemid)
	Dim i, torf, itoption, itorserial
	itorserial = ""
	IF isArray(arr) THEN
		For i =0 To UBound(arr,2)
			If CStr(arr(1,i)) =  CStr(itemid) and CStr(arr(3,i)) =  0  Then	' 
'				torf = True
				itorserial = CStr(arr(0,i))
'				itoption = CStr(arr(2,i))
				Exit For
			End If
		Next
	End If
	fnMyEvalorderserial = itorserial
End Function

Function fnMyEvalitemoption(arr, itemid)
	Dim i, torf, itoption, itorserial
	itoption = False
	IF isArray(arr) THEN
		For i =0 To UBound(arr,2)
			If CStr(arr(1,i)) =  CStr(itemid) and CStr(arr(3,i)) =  0  Then	' 
		'		torf = True
		'		itorserial = CStr(arr(0,i))
				itoption = CStr(arr(2,i))
				Exit For
			End If
		Next
	End If
	fnMyEvalitemoption = itoption
End Function

%>
