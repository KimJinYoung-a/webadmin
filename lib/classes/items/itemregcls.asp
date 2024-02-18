<%
'####################################################
' Description :  �¶��� ���δ���ǰ
' History : ������ ����
'####################################################

Class CItemListItems
	public Fitemid
	public Fitemname
	public Fsellcash
	public FSuplyCash
	public Fmakername
	public Fregdate
	public Flastupdate
	public FrejectMsg
	public FrejectDate
	public FreRegMsg
	public FreRegDate

	public Fmakerid

	public FCurrState
	public FLinkitemid
	public FImgSmall
	public FSellyn
    
    public Fupchemanagecode
    
	public function GetCurrStateColor()
		GetCurrStateColor = "#000000"
		if FCurrState="1" then
			GetCurrStateColor = "#000000"
		elseif FCurrState="2" then
			GetCurrStateColor = "#FF0000"
		elseif FCurrState="7" then
			GetCurrStateColor = "#0000FF"
		elseif FCurrState="5" then
			GetCurrStateColor = "#008800"
		elseif FCurrState="9" then
			GetCurrStateColor = "#996600"
		elseif FCurrState="0" then
			GetCurrStateColor = "#FF0000"
		else
			GetCurrStateColor = "#000000"
		end if
	end function

	public function GetCurrStateName()
		GetCurrStateName = ""
		if FCurrState="1" then
			GetCurrStateName = "��ϴ��"
		elseif FCurrState="2" then
			GetCurrStateName = "��Ϻ���"
		elseif FCurrState="7" then
			GetCurrStateName = "��ϿϷ�"
		elseif FCurrState="5" then
			GetCurrStateName = "������û"
		elseif FCurrState="0" then
			GetCurrStateName = "��ϺҰ�" ''��ϰź�
		elseif FCurrState="9" then
			GetCurrStateName = "��ü���"
		else
			GetCurrStateName = ""
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

class CWaitItemlist
	public FItemList()

	public FTotalCount
	public FResultCount
	public FRectDesignerID
	public FRectSort
	public FRectUpchemanagecode
	
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FRectCurrState
	public FRectSellyn
	public FRectItemID
	public FRectLectureYN
	Public FRectitemname
	public Fcatecode

	Private Sub Class_Initialize()
	redim FItemList(0)
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	public sub WaitProductList()
		dim sqlStr,i,wheredetail, orderdetail
	    dim noCountQuery : noCountQuery = false
	    if (FRectitemname<>"") then noCountQuery=true  ''��ġ�� ������..;; 2015/04/13
	    
		if (FRectDesignerID<>"") then
			wheredetail = wheredetail + " and A.makerid='" + FRectDesignerID + "'"
		end if 
	 
        if (FRectCurrState="A") then
			wheredetail = wheredetail + " and A.currstate in ('1','2','0','5')"
		else 
			wheredetail = wheredetail + " and A.currstate='"+FRectCurrState+"'"
		end If
        
		if (FRectitemname<>"") then
		    ''wheredetail = wheredetail + " and A.itemname like '" + replace(FRectitemname,"[","[[]") + "%'"  '' ���� %  ����
			wheredetail = wheredetail + " and A.itemname like '" + replace(replace(FRectitemname,"[","[[]"),"'","''") + "%'"&VbCRLF
		end if
		
		if (Fcatecode<>"") then
			wheredetail = wheredetail + " and B.catecode like '" + Fcatecode + "%'"
		end if

		If (FRectUpchemanagecode <> "") Then
			wheredetail = wheredetail & " and A.upchemanagecode='"&FRectUpchemanagecode&"'"&VbCRLF
		End If
		 
		IF  FRectSort = "UD" THEN
		 orderdetail = " A.upchemanagecode Desc "
		ELSEIF FRectSort = "UA" THEN
		  orderdetail = " A.upchemanagecode Asc "
		ELSEIF FRectSort = "ND" THEN
		 orderdetail = " A.itemname Desc "
		ELSEIF FRectSort = "NA" THEN
		 orderdetail = " A.itemname Asc "
		ELSEIF FRectSort = "SD" THEN
		 orderdetail = " A.sellcash Desc "
		ELSEIF FRectSort = "SA" THEN
		 orderdetail = " A.sellcash Asc "
		ELSEIF FRectSort = "BD" THEN
		 orderdetail = " A.buycash Desc "
		ELSEIF FRectSort = "BA" THEN
		 orderdetail = " A.buycash Asc "
		ELSEIF FRectSort = "TD" THEN
		 orderdetail  = " A.itemid Desc "
		ELSEIF FRectSort = "TA" THEN
		 orderdetail = " A.itemid Asc " 
		ELSEIF FRectSort = "LA" THEN
		 orderdetail = " A.lastupdate Asc "
		ELSE
			 orderdetail = " A.lastupdate Desc " 
		END IF	 
		
		'###########################################################################
		'��ϴ�� ��ǰ �� ���� ���ϱ�
		'###########################################################################
		sqlStr = "select count(A.itemid) as cnt"
		sqlStr = sqlStr & " from [db_temp].[dbo].tbl_wait_item as A"
		if (Fcatecode<>"") then
		    sqlStr = sqlStr & "   LEFT OUTER JOIN db_temp.dbo.tbl_display_cate_waititem AS B ON A.itemid = B.itemid and B.isdefault = 'Y' "
		END IF
		sqlStr = sqlStr & " where A.itemid<>0"
		sqlStr = sqlStr & " and A.currstate<9"
		sqlStr = sqlStr & wheredetail

        if (NOT noCountQuery) then
        	'rsget.Open sqlStr,dbget,1
        	rsget.CursorLocation = adUseClient
            rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
        		FTotalCount = rsget("cnt")
        	rsget.Close
    	end if
    	
		'###########################################################################
		'��ϴ�� ��ǰ ������
		'###########################################################################

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " A.itemid,A.makerid,A.itemname,A.sellcash,A.buycash,"
		sqlStr = sqlStr & " A.linkitemid, A.currstate, IsNull(A.makername,'')as maker,A.lastupdate, A.upchemanagecode, A.rejectmsg, A.rejectDate"
		sqlStr = sqlStr & " from [db_temp].[dbo].tbl_wait_item as A "
		if (Fcatecode<>"") then
		    sqlStr = sqlStr & "   LEFT OUTER JOIN db_temp.dbo.tbl_display_cate_waititem AS B ON A.itemid = B.itemid and B.isdefault = 'Y' "
		END IF
		sqlStr = sqlStr & " where A.itemid<>0"
		sqlStr = sqlStr & " and A.currstate<9"
		sqlStr = sqlStr & wheredetail
		sqlStr = sqlStr & " order by "&orderdetail

 
		rsget.pagesize = FPageSize
		''rsget.Open sqlStr,dbget,1
		rsget.CursorLocation = adUseClient                                      ''2016/06/13 ��ļ���
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount =  rsget.RecordCount - (FPageSize*(FCurrPage-1))
        if (noCountQuery) then FTotalCount=FResultCount  ''2015/04/13
            
		''FTotalPage = CInt(FTotalCount\FPageSize) + 1
		FTotalPage = (FTotalCount\FPageSize)                                    ''2016/06/13 ��ļ���
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1  ''2016/06/13 ��ļ���


		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CItemListItems
				FItemList(i).Fitemid = rsget("itemid")
				FItemList(i).Fmakerid = rsget("makerid")
			    FItemList(i).Fitemname = db2html(rsget("itemname"))
				FItemList(i).Fsellcash = rsget("sellcash")
				FItemList(i).FSuplyCash = rsget("buycash")
				FItemList(i).Fmakername = rsget("maker")
				FItemList(i).Flastupdate = rsget("lastupdate")
				FItemList(i).Frejectmsg = rsget("rejectmsg")
				FItemList(i).FrejectDate = rsget("rejectDate")

				FItemList(i).FLinkitemid = rsget("linkitemid")
				FItemList(i).FCurrState = rsget("currstate")
				
				FItemList(i).Fupchemanagecode = db2html(rsget("upchemanagecode"))
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public sub UpdateProductList()
		dim sqlStr,i,wheredetail

		if (FRectDesignerID<>"") then
			wheredetail = wheredetail + " and i.makerid='" + FRectDesignerID + "'"
		end if


		if (FRectSellyn<>"") then
			wheredetail = wheredetail + " and i.sellyn='Y'"
		end if

		if (FRectItemID<>"") then
			wheredetail = wheredetail + " and i.itemid='" + FRectItemID + "'"
		end if



		'###########################################################################
		'��ϴ�� ��ǰ �� ���� ���ϱ�
		'###########################################################################
		sqlStr = "select count(i.itemid) as cnt"
		sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr & " where i.itemid<>0"
		sqlStr = sqlStr & wheredetail

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		'###########################################################################
		'��ϴ�� ��ǰ ������
		'###########################################################################

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " i.itemid,i.makerid,i.itemname,i.sellcash,i.buycash,i.sellyn,"
		sqlStr = sqlStr & " IsNull(i.makername,'')as maker, regdate, i.smallimage as imgsmall"
		sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr & " where i.itemid<>0"
		sqlStr = sqlStr & wheredetail
		sqlStr = sqlStr & " order by regdate Desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount =  rsget.RecordCount - (FPageSize*(FCurrPage-1))

		FTotalPage = CInt(FTotalCount\FPageSize) + 1


		redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CItemListItems

				FItemList(i).Fitemid = rsget("itemid")
				FItemList(i).Fmakerid = rsget("makerid")
			    FItemList(i).Fitemname = db2html(rsget("itemname"))
				FItemList(i).Fsellcash = rsget("sellcash")
				FItemList(i).FSuplyCash = rsget("buycash")
				FItemList(i).Fmakername = rsget("maker")
				FItemList(i).Fregdate = rsget("regdate")
				FItemList(i).Fsellyn = rsget("sellyn")
				FItemList(i).FImgSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("imgsmall")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

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




class CItemOptionItem
	public Fitemoption
	public Fitemoptionname
	public Fisusing
	public Foptsellyn
	public Foptlimityn
	public Foptlimitno
	public Foptlimitsold
	public Fcodeview                'deprecated( Fitemoptionname ���� �����Ѵ�. )
	public FcolorCnt

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

class CWaitItemDetail
'########################################
'�ӽõ�����
'########################################
	public FItemList()
	public FResultCount
	public FTotalCount
	
	public FWaitItemID
	public FMakerid
	public Flarge
	public Fmid
	public Fsmall
	public Fitemdiv
	public Fitemname
	public Fitemcontent
	public Fdesignercomment
	public Fitemsource
	public Fitemsize
	public Fsellcash
	public Fsellvat
	public Fbuycash
	public Fbuyvat
	public Fdeilverytype
	public Fsourcearea
	public Fsourcekind
	public Fmakername
	public Flimityn
	public Flimitno
	public Fvatinclude
	public Fpojangok
	public FrequireMakeDay

	public FMargin
	public FMileage
	public Fsellyn

	public Fitemgubun
	public Fusinghtml
	public Fkeywords
	public Fmwdiv
	public Fdeliverarea
	public Fdeliverfixday
	public Fmaeipdiv
	public Fordercomment
	public Foptioncnt
	public FDFcolorCd
	public FDFcolorImg

    public FCurrState
    public Frejectmsg
    public FrejectDate
    public FreRegMsg
    public FreRegDate
    
    public FsellEndDate
    public Fupchemanagecode
    
    public FRectDesignerID
    
    public FinfoDiv			'ǰ�����й�ȣ
    public FsafetyYn		'����������� ����
    public FsafetyDiv
    public FsafetyNum

	public Ffreight_min		'ȭ�� �ݼۺ�
	public Ffreight_max
	
	public Fctrstate
	public FitemWeight
	public FdeliverOverseas
	public Fisbn13
	public Fisbn10
	public FisbnSub
	public FaddMsg
	public FaddCarve 	
 	public FaddBox 		
 	public FaddSet		
 	public FaddCustom 
 	public FAuthInfo
 	public FAuthImg
 	public fpurchaseType
	public fdeliverytype
	
	public function getMwDiv()
		if (IsNull(Fmaeipdiv) or (Fmaeipdiv="")) then
			getMwDiv = Fmaeipdiv
		else
			getMwDiv = Fmaeipdiv
		end if
	end function

	public function getMwDivName()
		if (Fmaeipdiv = "U") then
		    getMwDivName = "��ü"
		elseif (Fmaeipdiv = "W") then
		    getMwDivName = "��Ź"
		else
		    getMwDivName = "����"
		end if
	end function

	Private Sub Class_Initialize()
		FResultCount = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub


public function fnGetWaitOptAddPrice(ByVal itemid)
 if isNull(itemid) then exit Function
	dim strSql, OptAddPrice
	strSql = " SELECT isNull(sum(optaddprice),0) as OptAddPrice FROM [db_temp].[dbo].tbl_wait_itemoption where itemid = " & itemid
	rsget.Open strSql,dbget,1 
		if Not rsget.Eof then
       fnGetWaitOptAddPrice = rsget(0)
      end if 
  rsget.Close
End Function

	public function getDesignerDefaultMargin()
		dim sqlStr
		sqlStr = "select top 1 defaultmargine from [db_user].[dbo].tbl_user_c"
		sqlStr = sqlStr & " where userid='" & FRectDesignerID & "'"
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			getDesignerDefaultMargin = rsget("defaultmargine")
		end if
		rsget.close
	end function

	public sub WaitProductDetail(byval itemid)
		dim sqlStr

		sqlStr = "select top 1  IsNULL(i.Cate_large,'') as Cate_large, IsNULL(i.Cate_mid,'') as Cate_mid, IsNULL(i.Cate_small,'') as Cate_small, i.itemdiv, i.itemname,"
		sqlStr = sqlStr & " i.itemid, i.makerid, i.itemcontent,i.designercomment,i.itemsource,i.itemsize,"
		sqlStr = sqlStr & " i.sellcash,i.buycash,i.mileage,i.sellyn, i.isbn13, i.isbn10, i.isbn_sub,"
		sqlStr = sqlStr & " i.deliverytype,i.sourcearea,i.makername,i.limityn,i.limitno,"
		sqlStr = sqlStr & " i.vatinclude,i.pojangok,i.itemgubun,i.usinghtml,"
		sqlStr = sqlStr & " i.keywords, i.mwdiv, i.deliverarea, i.deliverfixday, i.ordercomment, c.maeipdiv, i.optioncnt, i.currstate, "
		sqlStr = sqlStr & " i.rejectmsg, i.rejectDate, i.reRegMsg, i.reRegDate, i.sellEndDate, i.upchemanagecode, i.requireMakeDay, "
		sqlStr = sqlStr & " i.infoDiv, i.safetyYn, i.safetyDiv, i.safetyNum, i.freight_min, i.freight_max, i.deliverytype "
		sqlStr = sqlStr & " ,isNull(o.colorCode,'') as DFcolorCd, o.basicImage as DFcolorImg "
		sqlStr = sqlStr & "	,( select top 1 isNull(m.ctrState,0) from db_partner.dbo.tbl_partner as p with (nolock)"
		sqlStr = sqlStr & "		inner join db_partner.dbo.tbl_partner_ctr_master as m with (nolock) on p.groupid = m.groupid and m.ctrState >=0 "
		sqlStr = sqlStr & "		left outer join db_partner.dbo.tbl_partner_ctr_sub as s with (nolock) on m.ctrKey = s.ctrKey  "
		sqlStr = sqlStr & "		where p.id = i.makerid and ((m.makerid ='' and contracttype = 8) or (m.makerid = p.id and contracttype <> 8  and s.sellplace ='On'))"
		sqlStr = sqlStr & "		order by m.ctrState ) as ctrState 	"
		sqlStr = sqlStr & " ,isNull(i.itemWeight,0) as itemWeight, isNull(i.deliverOverseas,'N') as deliverOverseas, isNull(i.sourcekind,'0') as sourcekind "
		sqlStr = sqlStr & " ,i.addMsg, i.addCarve, i.addBox, i.addSet, i.addCustom, p.purchaseType"
		sqlStr = sqlStr & " from [db_temp].[dbo].tbl_wait_item i with (nolock)"
		sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c with (nolock) on i.makerid=c.userid"
		sqlStr = sqlStr & " left join ( "
		sqlStr = sqlStr & "		select top 1 itemid, colorCode, basicImage from [db_temp].[dbo].tbl_wait_item_colorOption with (nolock)"
		sqlStr = sqlStr & "		where itemoption='0000' and itemid='" & itemid & "'"
		sqlStr = sqlStr & "	) as o "
		sqlStr = sqlStr & "		on i.itemid=o.itemid "
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on i.makerid = p.id" & vbcrlf
		sqlStr = sqlStr & " where i.makerid='" & FRectDesignerID & "'"
		sqlStr = sqlStr & " and i.itemid='" & itemid & "'"
 
		'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget.RecordCount
			FResultCount = FTotalCount

			IF Not (rsget.EOF OR rsget.BOF) THEN
			Flarge			= rsget("Cate_large")
			Fmid			= rsget("Cate_mid")
			Fsmall			= rsget("Cate_small")
			Fitemdiv		= rsget("itemdiv")
			FWaitItemID		= rsget("itemid")
			FMakerid		= rsget("makerid")
			Fitemname		= db2html(rsget("itemname"))
			Fitemcontent	= db2html(rsget("itemcontent"))
			Fdesignercomment = db2html(rsget("designercomment"))
			Fitemsource     = db2html(rsget("itemsource"))
			Fitemsize		= db2html(db2html(rsget("itemsize")))
			Fsellcash		= db2html(rsget("sellcash"))
			Fbuycash		= db2html(rsget("buycash"))
			FMileage		= rsget("mileage")
			Fsellyn			= rsget("sellyn")
			Fdeilverytype	= rsget("deliverytype")
			Fsourcearea		= db2html(rsget("sourcearea"))
			Fmakername		= db2html(rsget("makername"))
			Flimityn		= rsget("limityn")
			Flimitno		= rsget("limitno")
			Fvatinclude		= rsget("vatinclude")
			Fpojangok		= rsget("pojangok")
			FrequireMakeDay	= rsget("requireMakeDay")

			Fitemgubun = rsget("itemgubun")
			Fusinghtml = rsget("usinghtml")
			Fkeywords  = db2html(rsget("keywords"))
			Fmwdiv		= rsget("mwdiv")
			Fdeliverarea		= rsget("deliverarea")
			Fdeliverfixday		= rsget("deliverfixday")
			Fmaeipdiv       = rsget("maeipdiv")
			Fordercomment   = db2html(rsget("ordercomment"))
            
            FsellEndDate     = rsget("sellEndDate")
            Fupchemanagecode = rsget("upchemanagecode")
            
			Foptioncnt   = rsget("optioncnt")
            FDFcolorCd	= rsget("DFcolorCd")
			FDFcolorImg	= rsget("DFcolorImg")
            Fcurrstate   = rsget("currstate")
            Frejectmsg	= rsget("rejectmsg")
            FrejectDate	= rsget("rejectDate")
            FreRegMsg	= rsget("reRegMsg")
            FreRegDate	= rsget("reRegDate")

            FinfoDiv	= rsget("infoDiv")
            FsafetyYn	= rsget("safetyYn"):	if(isNull(FsafetyYn) or FsafetyYn="") then FsafetyYn="N"
            FsafetyDiv	= rsget("safetyDiv")
            FsafetyNum	= rsget("safetyNum")

			Ffreight_min	= rsget("freight_min"):	if(isNull(Ffreight_min) or Ffreight_min="") then Ffreight_min=0
			Ffreight_max	= rsget("freight_max"):	if(isNull(Ffreight_max) or Ffreight_max="") then Ffreight_max=0
			Fctrstate = rsget("ctrState"): if (isNull(Fctrstate)) then  Fctrstate = 0
            if FDFcolorCd=0 then FDFcolorCd=""
            if (Fsellcash<>0) then FMargin     =  100-(Fbuycash/Fsellcash*100)
            FitemWeight          = rsget("itemWeight")
            FdeliverOverseas     = rsget("deliverOverseas")
            Fsourcekind					= rsget("sourcekind")
                
             FaddMsg 		= rsget("addMsg")   
             FaddCarve 	= rsget("addCarve")   
             FaddBox 		= rsget("addBox")   
             FaddSet		 = rsget("addSet")   
             FaddCustom = rsget("addCustom")   
             fdeliverytype = rsget("deliverytype")
             fpurchaseType = rsget("purchaseType")
			fisbn13	= rsget("isbn13")
			fisbn10	= rsget("isbn10")
			fisbnsub	= rsget("isbn_sub")
          END IF
		rsget.close
		
		'### ������������� ��� �ش����� ������.
		if FResultCount>0 then
			If FsafetyYn = "Y" Then
				sqlStr = "select tw.safetyDiv, tw.certNum from db_temp.[dbo].[tbl_safetycert_tenReg_waititem] as tw with (nolock)"
				sqlStr = sqlStr & "left join db_temp.[dbo].[tbl_safetycert_info_waititem] as iw with (nolock) on tw.itemid = iw.itemid and tw.certNum = iw.certNum "
				sqlStr = sqlStr & "where tw.itemid = '" & itemid & "'"

				'response.write sqlStr & "<Br>"
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
				if not rsget.eof then
					FAuthInfo = rsget.getRows()
	 			end if
	 			rsget.close
	 		End If
 		End If
	end sub

	public sub WaitProductDetailOption(byval itemid)
		dim sqlStr,i

        'TODO : ���Ŀ� tbl_option_div02 �� ���� ������ �����ؾ� �Ѵ�.
        '����� o.optionname �� �ɼǸ��� ������� �ʴ°�찡 �־� ���Խ��ѵд�.
        sqlStr = " select top 100 o.itemoption, isnull(o.optionname, o2.codeview) as itemoptionname,"
        sqlStr = sqlStr + " isusing, optsellyn, optlimityn, optlimitno, optlimitsold "
        sqlStr = sqlStr + " 	,(select count(colorCode) "
        sqlStr = sqlStr + " 		from [db_temp].[dbo].tbl_wait_item_colorOption with (nolock)"
        sqlStr = sqlStr + " 		where itemid=o.itemid "
        sqlStr = sqlStr + " 	) as colorCnt "
        sqlStr = sqlStr + " from [db_temp].[dbo].tbl_wait_itemoption o with (nolock)"
        sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_option_div02 o2 with (nolock) on ((Left(o.itemoption, 2) = o2.optioncode01) and (Right(o.itemoption, 2) = o2.optioncode02)) "
        sqlStr = sqlStr + " where o.itemid = " + CStr(itemid) + " "
        sqlStr = sqlStr + " and o.itemoption<>''"
        sqlStr = sqlStr + " order by itemoption "

		'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)

			do until rsget.Eof
				set FItemList(i) = new CItemOptionItem

				FItemList(i).Fitemoption    = rsget("itemoption")
				FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).Fisusing       = rsget("isusing")
				FItemList(i).Foptsellyn     = rsget("optsellyn")
				FItemList(i).Foptlimityn    = rsget("optlimityn")
				FItemList(i).Foptlimitno    = rsget("optlimitno")
				FItemList(i).Foptlimitsold  = rsget("optlimitsold")
				FItemList(i).Fcodeview      = db2html(rsget("itemoptionname"))
				FItemList(i).FcolorCnt		= rsget("colorCnt")

				rsget.movenext
				i=i+1
			loop

		rsget.Close
	end sub


end Class

class CWaitItemImagelist
	public Fimgtitle
	public Fimgmain
	public Fimgmain2
	public Fimgmain3
	public Fimgmain4
	public Fimgmain5
	public Fimgmain6
	public Fimgmain7
	public Fimgmain8
	public Fimgmain9
	public Fimgmain10
	public Fimgsmall
	public Fimglist
	public Fimgbasic
	public Fimgmask
	public Ficon1
	public Ficon2
	public Fimgadd
	public Fimgstory
	public Fitemaddcontent
	public FRectItemID
	' �߰��Ѻκ�
	public Fmobileimgmain
	public Fmobileimgmain2
	public Fmobileimgmain3
	public Fmobileimgmain4
	public Fmobileimgmain5
	public Fmobileimgmain6
	public Fmobileimgmain7
	public Fmobileimgmain8
	public Fmobileimgmain9
	public Fmobileimgmain10
	public Fmobileimgmain11
	public Fmobileimgmain12
	'// �߰��Ѻκ�
    
    public Flistimage120
    
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	public sub WaitProductImageList(byval itemid)
		dim sqlStr
		sqlStr = "select top 1 imgtitle,imgmain,imgmain2,imgmain3,imgmain4,imgmain5,imgmain6,imgmain7,imgmain8,imgmain9,imgmain10,imgsmall,imglist,imgbasic,imgmask,"
		sqlStr = sqlStr & " icon1,icon2,imgadd,imgstory,itemaddcontent,listimage120, mobileimg, mobileimg2, mobileimg3, mobileimg4, mobileimg5, mobileimg6, mobileimg7, mobileimg8, mobileimg9, mobileimg10, mobileimg11, mobileimg12"
		sqlStr = sqlStr & " from [db_temp].[dbo].tbl_wait_item_image"
		sqlStr = sqlStr & " where itemid='" & itemid & "'"

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			Fimgtitle = rsget("imgtitle")
			Fimgmain = rsget("imgmain")
			Fimgmain2 = rsget("imgmain2")
			Fimgmain3 = rsget("imgmain3")
			Fimgmain4 = rsget("imgmain4")
			Fimgmain5 = rsget("imgmain5")
			Fimgmain6 = rsget("imgmain6")
			Fimgmain7 = rsget("imgmain7")
			Fimgmain8 = rsget("imgmain8")
			Fimgmain9 = rsget("imgmain9")
			Fimgmain10 = rsget("imgmain10")
			Fimgbasic = rsget("imgbasic")
			Fimgmask = rsget("imgmask")
			Ficon1 = rsget("icon1")
			Ficon2 = rsget("icon2")
			Fimgsmall = rsget("imgsmall")
			Fimglist = rsget("imglist")
			Fimgadd = rsget("imgadd")
			Fimgstory = rsget("imgstory")
			Fitemaddcontent = rsget("itemaddcontent")
			Flistimage120 = rsget("listimage120")
			' �߰��Ѻκ�
			Fmobileimgmain = rsget("mobileimg")
			Fmobileimgmain2 = rsget("mobileimg2")
			Fmobileimgmain3 = rsget("mobileimg3")
			Fmobileimgmain4 = rsget("mobileimg4")
			Fmobileimgmain5 = rsget("mobileimg5")
			Fmobileimgmain6 = rsget("mobileimg6")
			Fmobileimgmain7 = rsget("mobileimg7")
			Fmobileimgmain8 = rsget("mobileimg8")
			Fmobileimgmain9 = rsget("mobileimg9")
			Fmobileimgmain10 = rsget("mobileimg10")
			Fmobileimgmain11 = rsget("mobileimg11")
			Fmobileimgmain12 = rsget("mobileimg12")
			'// �߰��Ѻκ�
		end if
		rsget.close
	end sub
	
    public Function GetAddImageListIMGTYPE1()
	    dim sqlstr, i

	    sqlstr = "select top 100 * from [db_item].[dbo].tbl_item_addimage"
	    sqlstr = sqlstr & " where itemid=" & FRectItemID & " and IMGTYPE = 1 "
	    sqlstr = sqlstr & " ORDER BY GUBUN asc"
	    rsget.Open sqlStr,dbget,1
	    If Not rsget.Eof Then
	    	GetAddImageListIMGTYPE1 = rsget.getrows()
	    End If
        rsget.Close
    end Function

end Class

class CItemReg

	public largename
	public midname
	public smallname
	public optionbigname
	public optionbigno
	public FRectDesignerID
	public Fitemid
	public FItemoption
	public FMainImage
	public optioncodename


	Private Sub Class_Initialize()


	End Sub

	Private Sub Class_Terminate()

	End Sub


	function CheckFiles(ifile)
		dim file1_size, file1_name
		dim extension

		if (ifile="") then
			CheckFiles =0
			exit function
		end if

		file1_size = ifile.FileLen
	    file1_name = ifile.FileName
	    extension = LCase(Mid(file1_name, InStrRev(file1_name, ".")))

	    if (file1_size>100000) then
	    	response.write "<script language='javascript'>alert('���ϻ����� 100,000Byte ���� �����˴ϴ�.'); history.go(-1);</script>"
	        dbget.close()	:	response.End
	    	exit function
	    end if

	    if ((extension <> ".gif") and (extension <> ".jpg") and (extension <> ".png")) then
	    	response.write "<script language='javascript'>alert('�̹���(gif,jpg,png) ȭ�ϸ� �����˴ϴ�.'); history.go(-1);</script>"
	        dbget.close()	:	response.End
	    	exit function
	    end if
	    CheckFiles =0
	end function

'	public sub SearchOptionNameBig(byval optionno)
'		dim sqlStr
'
'		sqlStr = "select codename"
'		sqlStr = sqlStr + " from [db_item].[dbo].tbl_option_div01"
'		sqlStr = sqlStr + " where optioncode01='" + Cstr(optionno) + "'"
'
'		rsget.Open sqlStr,dbget,1
'		if Not rsget.Eof then
'			optioncodename = rsget("codename")
'	    end if
'		rsget.close
'
'	end sub


	public sub SearchCategoryNameLarge(byval largeno)
		dim sqlStr

		sqlStr = "select code_nm"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_Cate_large"
		sqlStr = sqlStr + " where code_large='" + Cstr(largeno) + "'"

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			largename = db2html(rsget("code_nm"))
		end if
		rsget.close
	end sub

	public sub SearchCategoryNameMid(byval largeno,midno)
		dim sqlStr

		sqlStr = "select code_nm"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_Cate_mid"
		sqlStr = sqlStr + " where code_large='" + Cstr(largeno) + "'"
		sqlStr = sqlStr + " and code_mid='" + Cstr(midno) + "'"

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			midname = db2html(rsget("code_nm"))
		end if
		rsget.close
	end sub

	public sub SearchCategoryNameSmall(byval largeno,midno,smallno)
		dim sqlStr

		sqlStr = "select code_nm"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_Cate_small"
		sqlStr = sqlStr + " where code_large='" + Cstr(largeno) + "'"
		sqlStr = sqlStr + " and code_mid='" + Cstr(midno) + "'"
		sqlStr = sqlStr + " and code_small='" + Cstr(smallno) + "'"

		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			smallname = db2html(rsget("code_nm"))
		end if
		rsget.close
	end sub

	public sub SearchOptionName(byval bigno)
		dim sqlStr

		sqlStr = "select optioncode01,codename"
		sqlStr = sqlStr & " from [db_item].[dbo].tbl_option_div01"
		sqlStr = sqlStr & " where optioncode01='" & Cstr(bigno) & "'"
		rsget.Open sqlStr,dbget,1
			optionbigname = rsget("codename")
			optionbigno = rsget("optioncode01")
		rsget.close
	end sub

	function FormatStr(n,orgData)
			dim tmp
			if (n-Len(CStr(orgData))) < 0 then
				FormatStr = CStr(orgData)
				Exit Function
			end if

			tmp = String(n-Len(CStr(orgData)), "0") & CStr(orgData)
			FormatStr = tmp
	end Function


end Class




%>