<%

'' !!!! �Ʒ� ������ ��Ŭ��� �Ǿ� �־�� �Ѵ�.
''/admin/etc/lotte/inc_dailyAuthCheck.asp
''/lib/classes/etc/lotteitemcls.asp
''/admin/etc/incOutMallCommonFunction.asp


Class CxSiteTmpCSItem

    public Fidx
	public Fdivcd
	public Fdivname
	public Fgubunname
    public FSellSite
    public FOutMallOrderSerial
    public FOrgDetailKey
	public FCSDetailKey
	public FOrderSerial
	public Fmakerid
	public FItemID
	public Fitemoption
	public Fitemno
	public FextItemno
	public FOutMallItemName
	public FOutMallItemOptionName
    public FOrderName
    public FOrderEmail
    public FOrderTelNo
    public FOrderHpNo
    public FReceiveName
    public FReceiveTelNo
    public FReceiveHpNo
    public FReceiveZipCode
    public FReceiveAddr1
    public FReceiveAddr2
    public Fdeliverymemo
	public Fcurrstate
	public Fdeleteyn
    public FOutMallRegDate
    public FRegDate
	public Ftencsdivname
	public Ftencscnt
	public Fipkumdate

	public ForgOutMallOrderSerial				'// ���ֹ���ȣ(����� �ֹ��� ���)

	public Fjupsucscnt
	public Fupcheconfirmcscnt
	public Ffinishcscnt
	public Fdelcscnt
	public Fipkumdiv
	public Fcancelyn

	public FOutMallFinishDate
	public FOutMallCurrState

    public Fasid
    public Fascurrstate

	public function IpkumDivColor()
		if Fipkumdiv="0" then
			IpkumDivColor="#FF0000"
		elseif Fipkumdiv="1" then
			IpkumDivColor="#44BBBB"
		elseif Fipkumdiv="2" then
			IpkumDivColor="#000000"
		elseif Fipkumdiv="3" then
			IpkumDivColor="#000000"
		elseif Fipkumdiv="4" then
			IpkumDivColor="#0000FF"
		elseif Fipkumdiv="5" then
			IpkumDivColor="#CC9933"
		elseif Fipkumdiv="6" then
			IpkumDivColor="#FF00FF"
		elseif Fipkumdiv="7" then
			IpkumDivColor="#EE2222"
		elseif Fipkumdiv="8" then
			IpkumDivColor="#EE2222"
		elseif Fipkumdiv="9" then
			IpkumDivColor="#FF0000"
		end if
	end function

	Public function IpkumDivName()
		if Fipkumdiv="0" then
			IpkumDivName="�ֹ����"
		elseif Fipkumdiv="1" then
			IpkumDivName="�ֹ�����"
		elseif Fipkumdiv="2" then
			IpkumDivName="�ֹ�����"
		elseif Fipkumdiv="3" then
			IpkumDivName="�ֹ�����(3)"
		elseif Fipkumdiv="4" then
			IpkumDivName="�����Ϸ�"
		elseif Fipkumdiv="5" then
			IpkumDivName="�ֹ��뺸"
		elseif Fipkumdiv="6" then
			IpkumDivName="��ǰ�غ�"
		elseif Fipkumdiv="7" then
			IpkumDivName="�Ϻ����"
	    elseif Fipkumdiv="8" then
			IpkumDivName="��ǰ���"
		else
			IpkumDivName=Fipkumdiv
		end if
	end Function

	public function GetCurrStateName
		if (Fcurrstate = "B001") then
			GetCurrStateName = "�������"
		elseif (Fcurrstate = "B002") then
			GetCurrStateName = "�����Ϸ�"
		elseif (Fcurrstate = "B007") then
			GetCurrStateName = "��ϿϷ�"
		else
			GetCurrStateName = Fcurrstate
		end if
	end function

	public function GetCurrStateColor
		if (Fcurrstate = "B001") then
			GetCurrStateColor = "black"
		elseif (Fcurrstate = "B002") then
			GetCurrStateColor = "#CC9933"
		elseif (Fcurrstate = "B007") then
			GetCurrStateColor = "blue"
		else
			GetCurrStateColor = "red"
		end if
	end function

	public function GetExtCurrStateName
		if (FOutMallCurrState = "B007") then
			GetExtCurrStateName = "�Ϸ�"
		elseif (FOutMallCurrState = "B001") then
			GetExtCurrStateName = "����"
		elseif (FOutMallCurrState = "B008") then
			GetExtCurrStateName = "����"
		else
			GetExtCurrStateName = FOutMallCurrState
		end if
	end function

public function GetExtCurrStateColor
		if (FOutMallCurrState = "B007") then
			GetExtCurrStateColor = "blue"
		elseif (FOutMallCurrState = "B008") then
			GetExtCurrStateColor = "#FF0000"
		else
			GetExtCurrStateColor = "#000000"
		end if
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class


Class CxSiteCSOrder
    public FItemList()
	public FOneItem
	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage
	public FRectSellSite
	public FRectOrderSerial
	public FRectOutMallOrderSerial
	public FRectCurrState
	public FRectStartDate
	public FRectEndDate
	public FRectDivCD
	public FRectExcNoOrder
	public FRectMakerid
	public FRectOrderBy
	public FResultStr
    public FRectCsRegYN
    public FRectAsID

	public function getExtJungsanNoneList()
		dim i,sqlStr, addSql

		addSql = ""

		sqlStr = " exec [db_dataSummary].[dbo].[usp_Ten_Check_Outmall_OutJungsanNone] '" & FRectStartDate & "', '" & FRectEndDate & "','" & FRectSellSite & "' "

        db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = db3_rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof

				set FItemList(i) = new CxSiteTmpCSItem

				FItemList(i).FSellSite 			= db3_rsget("sitename")
				FItemList(i).Forderserial		= db3_rsget("orderserial")
				FItemList(i).FItemID			= db3_rsget("ItemID")
				FItemList(i).Fitemno			= db3_rsget("itemno")
				FItemList(i).FOutMallOrderSerial	= db3_rsget("extOrgOrderserial")
				FItemList(i).FextItemno			= db3_rsget("extItemno")
				FItemList(i).Fipkumdate			= db3_rsget("ipkumdate")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

	end function

	public function getCSMasterList()
	    dim i,sqlStr, addSql

		addSql = ""
	    if (FRectSellSite<>"") then
    	    addSql = addSql & " and m.SellSite='"&FRectSellSite&"'"
    	end if

	    if (FRectOutMallOrderSerial<>"") then
    	    addSql = addSql & " and m.OutMallOrderSerial = '"&FRectOutMallOrderSerial&"'"
    	end if

	    if (FRectOrderSerial<>"") then
    	    addSql = addSql & " and m.OrderSerial='"&FRectOrderSerial&"'"
    	end if

	    if (FRectCurrState<>"") then
    	    addSql = addSql & " and m.currstate='"&FRectCurrState&"'"
    	end if

	    if (FRectStartDate<>"") then
    	    addSql = addSql & " and m.regdate>='"&FRectStartDate&"'"
    	end if

	    if (FRectEndDate<>"") then
    	    addSql = addSql & " and m.regdate<'"&FRectEndDate&"'"
    	end if

	    if (FRectDivCD<>"") then
    	    addSql = addSql & " and m.divcd = '"&FRectDivCD&"'"
    	end if

        if (FRectAsID <> "") then
            addSql = addSql & " and m.asid = '"&FRectAsID&"'"
        end if

	    if (FRectCsRegYN<>"") then
            if (FRectCsRegYN = "N") then
                addSql = addSql & " and ( "
                addSql = addSql & " 	(m.asid is not NULL and a.deleteyn = 'Y') "
                addSql = addSql & " 	or "
                addSql = addSql & " 	(m.asid is NULL and IsNull(m.OutMallCurrState, 'B001') <> 'B008') "
                addSql = addSql & " 	or "
                addSql = addSql & " 	(m.asid is not NULL and IsNull(m.OutMallCurrState, 'B001') = 'B008') "
                addSql = addSql & " ) "
            elseif (FRectCsRegYN = "Y") then
                '// ���޿Ϸ� ����
                addSql = addSql & " and m.asid is not NULL and a.currstate = 'B007' and a.deleteyn = 'N' and IsNull(m.OutMallCurrState, 'B001') < 'B007' "
            elseif (FRectCsRegYN = "A") then
                '// ���޿Ϸ�
                addSql = addSql & " and m.asid is not NULL and a.currstate = 'B007' and a.deleteyn = 'N' and IsNull(m.OutMallCurrState, 'B001') >= 'B007' "
            else
                addSql = addSql & " and m.asid is not NULL and a.currstate < 'B007' and a.deleteyn = 'N' "
            end if
    	end if

	    if (FRectExcNoOrder<>"") then
    	    addSql = addSql & " and ((m.OrderSerial is not NULL) or (m.SellSite = 'cjmall' and m.divcd = 'A009')) "			'// CJ�� ������ ǥ��
    	end if

	    if (FRectMakerid<>"") then
    	    addSql = addSql & " and d.makerid = '"&FRectMakerid&"'"
    	end if


	    sqlStr = "select count(*) as cnt , CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg"
	    sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPCS m"
        sqlStr = sqlStr & " LEFT JOIN [db_cs].[dbo].tbl_new_as_list a ON m.asid = a.id "
		if (FRectMakerid<>"") then
			sqlStr = sqlStr & " LEFT JOIN [db_order].[dbo].[tbl_order_master] o ON m.orderserial = o.orderserial "
			sqlStr = sqlStr & " LEFT JOIN [db_order].[dbo].[tbl_order_detail] d ON m.orderserial = d.orderserial and m.itemid = d.itemid and m.itemoption = d.itemoption "
		end if
	    sqlStr = sqlStr & " where 1=1"
		sqlStr = sqlStr & addSql

		''response.write sqlstr & "<Br>"
    	rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit function
		end if

		FResultStr = ""

		sqlStr = " SELECT "
		sqlStr = sqlStr & " 	d.makerid, "
		sqlStr = sqlStr & " 	count(*) as cnt "
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_xSite_TMPCS m "
		sqlStr = sqlStr & " JOIN [db_order].[dbo].[tbl_order_master] o ON m.orderserial = o.orderserial "
		sqlStr = sqlStr & " JOIN [db_order].[dbo].[tbl_order_detail] d ON m.orderserial = d.orderserial AND m.itemid = d.itemid AND m.itemoption = d.itemoption "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " GROUP BY "
		sqlStr = sqlStr & " 	d.makerid "
		sqlStr = sqlStr & " HAVING "
		sqlStr = sqlStr & " 	count(*) > 1 "
		sqlStr = sqlStr & " ORDER BY count(*) DESC "
		''rsget.Open sqlStr,dbget,1
		''if  not rsget.EOF  then
		''	do until rsget.eof
		''		if (FResultStr = "") then
		''			FResultStr = rsget("makerid") & "(" & rsget("cnt") & ")"
		''		else
		''			FResultStr = FResultStr & ", " & rsget("makerid") & "(" & rsget("cnt") & ")"
		''		end if
		''		rsget.moveNext
		''	loop
		''end if
		''rsget.Close

		sqlStr = " SELECT TOP " + CStr(FPageSize*FCurrPage) + " m.* "
		sqlStr = sqlStr & " 	,C1.comm_name AS divname "
		sqlStr = sqlStr & " 	,( "
		sqlStr = sqlStr & " 		SELECT IsNull(max(C1.comm_name), '') "
		sqlStr = sqlStr & " 		FROM db_cs.dbo.tbl_new_as_list t "
		sqlStr = sqlStr & " 		LEFT JOIN [db_cs].[dbo].tbl_cs_comm_code C1 ON t.divcd = C1.comm_cd "
		sqlStr = sqlStr & " 		WHERE t.orderserial = m.orderserial and t.deleteyn = 'N' "
		sqlStr = sqlStr & " 		) AS tencsdivname "
		sqlStr = sqlStr & " 	,( "
		sqlStr = sqlStr & " 		SELECT count(*) "
		sqlStr = sqlStr & " 		FROM db_cs.dbo.tbl_new_as_list t "
		sqlStr = sqlStr & " 		LEFT JOIN [db_cs].[dbo].tbl_cs_comm_code C1 ON t.divcd = C1.comm_cd "
		sqlStr = sqlStr & " 		WHERE t.orderserial = m.orderserial and t.deleteyn = 'N' "
		sqlStr = sqlStr & " 		) AS tencscnt "
		sqlStr = sqlStr & " 	,( "
		sqlStr = sqlStr & " 		SELECT count(t.id) "
		sqlStr = sqlStr & " 		FROM db_cs.dbo.tbl_new_as_list t "
		sqlStr = sqlStr & " 		WHERE t.orderserial = m.orderserial and t.deleteyn = 'N' and t.currstate < 'B006' "
		sqlStr = sqlStr & " 		) AS jupsucscnt "
		sqlStr = sqlStr & " 	,( "
		sqlStr = sqlStr & " 		SELECT count(t.id) "
		sqlStr = sqlStr & " 		FROM db_cs.dbo.tbl_new_as_list t "
		sqlStr = sqlStr & " 		WHERE t.orderserial = m.orderserial and t.deleteyn = 'N' and t.currstate = 'B006' "
		sqlStr = sqlStr & " 		) AS upcheconfirmcscnt "
		sqlStr = sqlStr & " 	,( "
		sqlStr = sqlStr & " 		SELECT count(t.id) "
		sqlStr = sqlStr & " 		FROM db_cs.dbo.tbl_new_as_list t "
		sqlStr = sqlStr & " 		WHERE t.orderserial = m.orderserial and t.deleteyn = 'N' and t.currstate = 'B007' "
		sqlStr = sqlStr & " 		) AS finishcscnt "
		sqlStr = sqlStr & " 	,( "
		sqlStr = sqlStr & " 		SELECT count(t.id) "
		sqlStr = sqlStr & " 		FROM db_cs.dbo.tbl_new_as_list t "
		sqlStr = sqlStr & " 		WHERE t.orderserial = m.orderserial and t.deleteyn = 'Y' "
		sqlStr = sqlStr & " 		) AS delcscnt "
		sqlStr = sqlStr & " 	,o.ipkumdiv, o.cancelyn, d.makerid, m.OutMallFinishDate, m.OutMallCurrState, m.asid, a.currstate as ascurrstate "
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_xSite_TMPCS m "
		sqlStr = sqlStr & " LEFT JOIN [db_cs].[dbo].tbl_cs_comm_code C1 ON m.divcd = C1.comm_cd "
        sqlStr = sqlStr & " LEFT JOIN [db_cs].[dbo].tbl_new_as_list a ON m.asid = a.id "
		sqlStr = sqlStr & " LEFT JOIN [db_order].[dbo].[tbl_order_master] o ON m.orderserial = o.orderserial "
		sqlStr = sqlStr & " LEFT JOIN [db_order].[dbo].[tbl_order_detail] d ON m.orderserial = d.orderserial and m.itemid = d.itemid and m.itemoption = d.itemoption "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
    	sqlStr = sqlStr & " order by " ''CSDetailKey desc"
		select case FRectOrderBy
			case "2":
				sqlStr = sqlStr & " m.OutMallRegDate desc, m.idx desc "
			case else:
				sqlStr = sqlStr & " m.idx desc "
		end select

	If (session("ssBctID")="kjy8517") Then
		response.write sqlStr & "<Br>"
	End If
	    rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CxSiteTmpCSItem

				FItemList(i).Fidx             		= rsget("idx")
				FItemList(i).FSellSite             	= rsget("SellSite")
				FItemList(i).FOutMallOrderSerial    = rsget("OutMallOrderSerial")
				FItemList(i).FOrgDetailKey          = rsget("OrgDetailKey")
				FItemList(i).FCSDetailKey          	= rsget("CSDetailKey")
				FItemList(i).FOrderSerial           = rsget("OrderSerial")
				FItemList(i).Fdivcd             	= rsget("divcd")

				FItemList(i).Fdivname             	= rsget("divname")
				if ((FItemList(i).FSellSite = "lotteCom" or FItemList(i).FSellSite = "lotteimall") and FItemList(i).Fdivcd = "A004") then
					'// �Ե������� ��ǰ/��ȯ ���� ������ �� ����.
					FItemList(i).Fdivname             	= "��ǰ/��ȯ"
				elseif (FItemList(i).FSellSite = "cjmall" and FItemList(i).Fdivcd = "A004") then
					'// ��ǰ : �����ǰ���� �ٹ��ǰ������ ��ǰ�� ���� �����Ѵ�.
					FItemList(i).Fdivname             	= "��ǰ"
				elseif (FItemList(i).FSellSite = "cjmall" and FItemList(i).Fdivcd = "A011") then
					'// ��ȯ : ���豳ȯ���� �ٹ豳ȯ������ ��ǰ�� ���� �����Ѵ�.
					FItemList(i).Fdivname             	= "��ȯ"
				elseif (FItemList(i).Fdivcd = "A088" or FItemList(i).Fdivcd = "A044" or FItemList(i).Fdivcd = "A090") then
					if (FItemList(i).Fdivcd = "A088") then
						FItemList(i).Fdivname             	= "�ֹ���� öȸ"
					elseif (FItemList(i).Fdivcd = "A044") then
						FItemList(i).Fdivname             	= "��ǰ öȸ"
					elseif (FItemList(i).Fdivcd = "A090") then
						FItemList(i).Fdivname             	= "��ȯ öȸ"
					end if
				elseif (FItemList(i).FSellSite = "coupang" and FItemList(i).Fdivcd = "A004") then
					'// ������ ���/��ǰ�� �����ִ�.
					FItemList(i).Fdivname             	= "���(��ǰ)"
				else
					'//
				end if

				FItemList(i).Fgubunname             = rsget("gubunname")
				FItemList(i).FOrderName             = db2html(rsget("OrderName"))
				FItemList(i).FOrderEmail            = rsget("OrderEmail")
				FItemList(i).FOrderTelNo            = rsget("OrderTelNo")
				FItemList(i).FOrderHpNo             = rsget("OrderHpNo")
				FItemList(i).FReceiveName           = db2html(rsget("ReceiveName"))
				FItemList(i).FReceiveTelNo          = rsget("ReceiveTelNo")
				FItemList(i).FReceiveHpNo           = rsget("ReceiveHpNo")
				FItemList(i).FReceiveZipCode        = rsget("ReceiveZipCode")
				FItemList(i).FReceiveAddr1          = rsget("ReceiveAddr1")
				FItemList(i).FReceiveAddr2          = rsget("ReceiveAddr2")
				FItemList(i).Fdeliverymemo          = rsget("deliverymemo")
				FItemList(i).FOutMallRegDate        = rsget("OutMallRegDate")
				FItemList(i).FRegDate             	= rsget("RegDate")

				FItemList(i).FItemID             	= rsget("ItemID")
				FItemList(i).Fitemoption            = rsget("itemoption")
				FItemList(i).Fmakerid	            = rsget("makerid")

				FItemList(i).Fitemno             	= rsget("itemno")
				if ((FItemList(i).FSellSite = "lotteCom" or FItemList(i).FSellSite = "lotteimall") and FItemList(i).Fdivcd = "A004") then
					'// �Ե������� ��ǰ/��ȯ ���� ���������� ����.
					FItemList(i).Fitemno             	= ""
				end if

				FItemList(i).FOutMallItemName       = db2html(rsget("OutMallItemName"))
				FItemList(i).FOutMallItemOptionName = db2html(rsget("OutMallItemOptionName"))

				FItemList(i).Ftencsdivname          = rsget("tencsdivname")
				FItemList(i).Ftencscnt             	= rsget("tencscnt")

				FItemList(i).Fcurrstate             = rsget("currstate")

				FItemList(i).ForgOutMallOrderSerial	= rsget("orgOutMallOrderSerial")

				FItemList(i).Fjupsucscnt			= rsget("jupsucscnt")
				FItemList(i).Fupcheconfirmcscnt		= rsget("upcheconfirmcscnt")
				FItemList(i).Ffinishcscnt			= rsget("finishcscnt")
				FItemList(i).Fdelcscnt				= rsget("delcscnt")
				FItemList(i).Fipkumdiv				= rsget("ipkumdiv")
				FItemList(i).Fcancelyn				= rsget("cancelyn")

				FItemList(i).FOutMallFinishDate		= rsget("OutMallFinishDate")
				FItemList(i).FOutMallCurrState		= rsget("OutMallCurrState")

                FItemList(i).Fasid					= rsget("asid")
                FItemList(i).Fascurrstate			= rsget("ascurrstate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FTotalPage =0
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

Class CxSiteCSOrderXML
    public FItemList()
	public FOneItem
	public FResultCount
	public FTotalCount
	public FRectDivCD
	public FRectSellSite
	public FRectOutMallOrderSerial
	public FRectYYYYMMDD
	public FRectStartYYYYMMDD
	public FRectEndYYYYMMDD

	public ErrMsg
	private objXML
	private xmlDOM

	private xmlURL
	private objData

	public function SavexSiteCSOrderListtoDB()
		ErrMsg = ""

		if (ErrMsg = "") then
			xmlURL = GetXMLURL()
			''response.write xmlURL

			if (xmlURL = "") and (ErrMsg = "") then
				ErrMsg = "��ϵ��� ���� ���޸��Դϴ�.[0]"
			end if
		end if

        ''response.write xmlURL
		''response.write "<br>�׽�Ʈ ���Դϴ�.<br>"
		''dbget.close()
		''response.end

		if (ErrMsg = "") then
			Call GetXmlFromWeb()

			if (objData = "") and (ErrMsg = "") then
				ErrMsg = "�Ե����İ� ����߿� ������ �߻��߽��ϴ�."
			end if
		end if

		if (ErrMsg = "") then
			Call ActSavexSiteCSOrderListtoDB()
		end if

    end function

	public function SendxSiteSongjangNo(ord_no, ord_dtl_sn, sendQnt, sendDate, outmallGoodsID, hdc_cd, inv_no)
		ErrMsg = ""

		if (ErrMsg = "") then
			xmlURL = GetXMLURL()
			'response.write xmlURL

			if (xmlURL = "") and (ErrMsg = "") then
				ErrMsg = "��ϵ��� ���� ���޸��Դϴ�.[0]"
			end if
		end if

		if (ErrMsg = "") then
			Call GetXmlFromWeb()

			if (objData = "") and (ErrMsg = "") then
				ErrMsg = "�Ե����İ� ����߿� ������ �߻��߽��ϴ�."
			end if
		end if

		if (ErrMsg = "") then
			Call GetSongjangSendResult()
		end if

    end function

	function ActSavexSiteCSOrderListtoDB()
		dim i, j
		dim objMasterListXML, objMasterOneXML, objDetailListXML, objDetailOneXML
		dim masterCnt, detailCnt
		dim divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo, OutMallRegDate
		dim OrgDetailKey, CSDetailKey, itemno
		dim strSql

		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		xmlDOM.LoadXML replace(objData,"&","��")

		if (FRectSellSite = "lotteCom") or (FRectSellSite = "lotteimall") then
			'// �Ե�����, �Ե�i��

			'// ������
			set objMasterListXML = xmlDOM.selectNodes("/Response/Result/OrderInfo")

			masterCnt = objMasterListXML.length

			if (masterCnt > 0) then
				for i = 0 to masterCnt - 1
					set objMasterOneXML = objMasterListXML.item(i)

					divcd				= FRectDivCD
					gubunname			= objMasterOneXML.selectSingleNode("DelvInfo/ClmCausNm").text
					SellSite			= FRectSellSite
					OutMallOrderSerial	= objMasterOneXML.selectSingleNode("OrdNo").text
					OrderName			= objMasterOneXML.selectSingleNode("DelvInfo/OrdManNm").text
					OrderEmail			= ""
					OrderTelNo			= ""
					OrderHpNo			= ""
					ReceiveName			= objMasterOneXML.selectSingleNode("DelvInfo/RmitNm").text
					ReceiveTelNo		= ""
					ReceiveHpNo			= ""
					ReceiveZipCode		= ""
					ReceiveAddr1		= objMasterOneXML.selectSingleNode("DelvInfo/Addr").text
					ReceiveAddr2		= ""
					deliverymemo		= ""

					if (FRectDivCD = "A008") then
						OutMallRegDate		= objMasterOneXML.selectSingleNode("DelvInfo/CnclDtime").text
					else
						OutMallRegDate		= Left(now, 10)
					end if

					'// ������
					set objDetailListXML = objMasterOneXML.selectNodes("ProdInfo")
					detailCnt = objDetailListXML.length
					for j = 0 to detailCnt - 1
						set objDetailOneXML = objDetailListXML.item(j)

						OrgDetailKey	= objDetailOneXML.selectSingleNode("OrgOrdDtlSn").text
						CSDetailKey		= objDetailOneXML.selectSingleNode("OrdDtlSn").text

						if (FRectDivCD = "A008") then
							itemno			= objDetailOneXML.selectSingleNode("CnclQty").text
						else
							itemno			= objDetailOneXML.selectSingleNode("OrdQty").text
						end if

						strSql = " if not exists (select idx from db_temp.dbo.tbl_xSite_TMPCS where SellSite = '" + CStr(SellSite) + "' and OutMallOrderSerial = '" + CStr(OutMallOrderSerial) + "' and OrgDetailKey = '" + CStr(OrgDetailKey) + "' and CSDetailKey = '" + CStr(CSDetailKey) + "') "
						strSql = strSql + " begin "
						strSql = strSql + " insert into db_temp.dbo.tbl_xSite_TMPCS(divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
						strSql = strSql + " , OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) "
						strSql = strSql + " values('" + CStr(divcd) + "', '" + html2db(CStr(gubunname)) + "', '" + html2db(CStr(SellSite)) + "', '" + html2db(CStr(OutMallOrderSerial)) + "', '" + html2db(CStr(OrderName)) + "', '" + html2db(CStr(OrderEmail)) + "', '" + html2db(CStr(OrderTelNo)) + "', '" + html2db(CStr(OrderHpNo)) + "', '" + html2db(CStr(ReceiveName)) + "', '" + html2db(CStr(ReceiveTelNo)) + "', '" + html2db(CStr(ReceiveHpNo)) + "', '" + html2db(CStr(ReceiveZipCode)) + "', '" + html2db(CStr(ReceiveAddr1)) + "', '" + html2db(CStr(ReceiveAddr2)) + "', '" + html2db(CStr(deliverymemo)) + "' "
						strSql = strSql + " , '" + html2db(CStr(OutMallRegDate)) + "', '" + html2db(CStr(OrgDetailKey)) + "', '" + html2db(CStr(CSDetailKey)) + "', " + CStr(itemno) + ") "
						strSql = strSql + " end "
						''rw strSql
						rsget.Open strSql, dbget, 1

						set objDetailOneXML = Nothing
					next

					set objDetailListXML = Nothing
					set objMasterOneXML = Nothing
				next
			end if

			set objMasterListXML = Nothing

			strSql = " update c "
			strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
			strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
			strSql = strSql + " from "
			strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
			strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
			strSql = strSql + " on "
			strSql = strSql + " 	1 = 1 "
			strSql = strSql + " 	and c.SellSite = o.SellSite "
			strSql = strSql + " 	and c.OutMallOrderSerial = Replace(o.OutMallOrderSerial, '-', '') "
			strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
			strSql = strSql + " where "
			strSql = strSql + " 	1 = 1 "
			strSql = strSql + " 	and c.orderserial is NULL "
			strSql = strSql + " 	and o.orderserial is not NULL "
			''rw strSql
			rsget.Open strSql, dbget, 1
		else
			ErrMsg = "�Ľ̿� �����߽��ϴ�."
		end if
		Set xmlDOM = Nothing
	end function

	function GetSongjangSendResult()
		dim i, j
		dim objMasterListXML, objMasterOneXML, objDetailListXML, objDetailOneXML
		dim masterCnt, detailCnt
		dim divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo, OutMallRegDate
		dim OrgDetailKey, CSDetailKey, itemno
		dim strSql
		dim IsSuccess

		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		xmlDOM.LoadXML replace(objData,"&","��")

		ErrMsg = ""

		if (FRectSellSite = "lotteimall") then
			'// �Ե���i��

			IsSuccess = False
			set objMasterListXML = xmlDOM.selectNodes("/Response/Result")
			if (objMasterListXML.length > 0) then
				IsSuccess = True
			end if

			if IsSuccess then
				'// ����
				strSql = " update db_temp.dbo.tbl_xSite_TMPOrder"
				strSql = strSql & " set sendstate=1"
				strSql = strSql & " ,sendreqCnt=IsNULL(sendreqCnt,0)+1"
				strSql = strSql & " where outmallorderserial='"&ord_no&"'"
				strSql = strSql & " and orgdetailkey='"&ord_dtl_sn&"'"
				strSql = strSql & " and IsNULL(sendstate,0)=0"
				strSql = strSql & " and IsNULL(matchstate,'') <> 'D' and ordercsgbn = 0"
				'rw strSql
				dbget.Execute strSql

				ErrMsg = "OK"
			else
				'// ����
				set objMasterListXML = xmlDOM.selectNodes("/Response/Errors")
				set objMasterOneXML = objMasterListXML.item(0)

				ErrMsg = objMasterOneXML.selectSingleNode("Error/Message").text

				strSql = " update db_temp.dbo.tbl_xSite_TMPOrder"
				strSql = strSql & " set sendreqCnt=IsNULL(sendreqCnt,0)+1"
				strSql = strSql & " where outmallorderserial='"&ord_no&"'"
				strSql = strSql & " and orgdetailkey='"&ord_dtl_sn&"'"
				strSql = strSql & " and IsNULL(sendstate,0)=0"
				strSql = strSql & " and IsNULL(matchstate,'') <> 'D' and IsNULL(ordercsgbn, 0) = 0"
				''response.write strSql
				dbget.Execute strSql

				'// ���� 3ȸ �̻��̸� ����ó��
				Dim errCount
				strSql = ""
				strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
				strSql = strSql & "	where OutMallOrderSerial='"&ord_no&"'"
				strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"
				strSql = strSql & " and sendReqCnt >= 3" & VBCRLF
				rsget.Open strSql,dbget,1
				If Not rsget.Eof Then
					errCount = rsget("cnt")
				End If
				rsget.Close

				If errCount > 0 Then
					response.write  "<select name='updateSendState' id=""updateSendState"">" &_
									"	<option value=''>����</option>" &_
									"	<option value='901'>�߼�ó������ �����ϰ�</option>" &_
									"	<option value='902'>����� ��������</option>" &_
									"	<option value='903'>��ǰó����</option>" &_
									"</select>&nbsp;&nbsp;"
					response.write "<input type='button' value='�Ϸ�ó��' onClick=""finCancelOrd2('"&ord_no&"','"&ord_dtl_sn&"',document.getElementById('updateSendState').value)""><br>"
					response.write "<script language='javascript'>"&VbCRLF
					response.write "function finCancelOrd2(ord_no,ord_dtl_sn,selectValue){"&VbCRLF
					response.write "    if(selectValue == ''){"&VbCRLF
					response.write "    	alert('�������ּ���');"&VbCRLF
					response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
					response.write "    	return;"&VbCRLF
					response.write "    }"&VbCRLF
					response.write "    var uri = 'xSiteCSOrder_lotteimall_Process.asp?mode=updateSendState&ord_no='+ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
					response.write "    var popwin = window.open(uri,'finCancelOrd2','width=200,height=200');"&VbCRLF
					response.write "    popwin.focus()"&VbCRLF
					response.write "}"&VbCRLF
					response.write "</script>"&VbCRLF
				End If
			end if

			'' '// ������
			'' set objMasterListXML = xmlDOM.selectNodes("/Response/Result/OrderInfo")

			'' masterCnt = objMasterListXML.length

			'' if (masterCnt > 0) then
			'' 	for i = 0 to masterCnt - 1
			'' 		set objMasterOneXML = objMasterListXML.item(i)

			'' 		divcd				= FRectDivCD
			'' 		gubunname			= objMasterOneXML.selectSingleNode("DelvInfo/ClmCausNm").text
			'' 		SellSite			= FRectSellSite
			'' 		OutMallOrderSerial	= objMasterOneXML.selectSingleNode("OrdNo").text
			'' 		OrderName			= objMasterOneXML.selectSingleNode("DelvInfo/OrdManNm").text
			'' 		OrderEmail			= ""
			'' 		OrderTelNo			= ""
			'' 		OrderHpNo			= ""
			'' 		ReceiveName			= objMasterOneXML.selectSingleNode("DelvInfo/RmitNm").text
			'' 		ReceiveTelNo		= ""
			'' 		ReceiveHpNo			= ""
			'' 		ReceiveZipCode		= ""
			'' 		ReceiveAddr1		= objMasterOneXML.selectSingleNode("DelvInfo/Addr").text
			'' 		ReceiveAddr2		= ""
			'' 		deliverymemo		= ""

			'' 		if (FRectDivCD = "A008") then
			'' 			OutMallRegDate		= objMasterOneXML.selectSingleNode("DelvInfo/CnclDtime").text
			'' 		else
			'' 			OutMallRegDate		= Left(now, 10)
			'' 		end if

			'' 		'// ������
			'' 		set objDetailListXML = objMasterOneXML.selectNodes("ProdInfo")
			'' 		detailCnt = objDetailListXML.length
			'' 		for j = 0 to detailCnt - 1
			'' 			set objDetailOneXML = objDetailListXML.item(j)

			'' 			OrgDetailKey	= objDetailOneXML.selectSingleNode("OrgOrdDtlSn").text
			'' 			CSDetailKey		= objDetailOneXML.selectSingleNode("OrdDtlSn").text

			'' 			if (FRectDivCD = "A008") then
			'' 				itemno			= objDetailOneXML.selectSingleNode("CnclQty").text
			'' 			else
			'' 				itemno			= objDetailOneXML.selectSingleNode("OrdQty").text
			'' 			end if

			'' 			strSql = " if not exists (select idx from db_temp.dbo.tbl_xSite_TMPCS where SellSite = '" + CStr(SellSite) + "' and OutMallOrderSerial = '" + CStr(OutMallOrderSerial) + "' and OrgDetailKey = '" + CStr(OrgDetailKey) + "' and CSDetailKey = '" + CStr(CSDetailKey) + "') "
			'' 			strSql = strSql + " begin "
			'' 			strSql = strSql + " insert into db_temp.dbo.tbl_xSite_TMPCS(divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
			'' 			strSql = strSql + " , OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) "
			'' 			strSql = strSql + " values('" + CStr(divcd) + "', '" + html2db(CStr(gubunname)) + "', '" + html2db(CStr(SellSite)) + "', '" + html2db(CStr(OutMallOrderSerial)) + "', '" + html2db(CStr(OrderName)) + "', '" + html2db(CStr(OrderEmail)) + "', '" + html2db(CStr(OrderTelNo)) + "', '" + html2db(CStr(OrderHpNo)) + "', '" + html2db(CStr(ReceiveName)) + "', '" + html2db(CStr(ReceiveTelNo)) + "', '" + html2db(CStr(ReceiveHpNo)) + "', '" + html2db(CStr(ReceiveZipCode)) + "', '" + html2db(CStr(ReceiveAddr1)) + "', '" + html2db(CStr(ReceiveAddr2)) + "', '" + html2db(CStr(deliverymemo)) + "' "
			'' 			strSql = strSql + " , '" + html2db(CStr(OutMallRegDate)) + "', '" + html2db(CStr(OrgDetailKey)) + "', '" + html2db(CStr(CSDetailKey)) + "', " + CStr(itemno) + ") "
			'' 			strSql = strSql + " end "
			'' 			''rw strSql
			'' 			rsget.Open strSql, dbget, 1

			'' 			set objDetailOneXML = Nothing
			'' 		next

			'' 		set objDetailListXML = Nothing
			'' 		set objMasterOneXML = Nothing
			'' 	next
			'' end if

			'' set objMasterListXML = Nothing

			'' strSql = " update c "
			'' strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
			'' strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
			'' strSql = strSql + " from "
			'' strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
			'' strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
			'' strSql = strSql + " on "
			'' strSql = strSql + " 	1 = 1 "
			'' strSql = strSql + " 	and c.SellSite = o.SellSite "
			'' strSql = strSql + " 	and c.OutMallOrderSerial = Replace(o.OutMallOrderSerial, '-', '') "
			'' strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
			'' strSql = strSql + " where "
			'' strSql = strSql + " 	1 = 1 "
			'' strSql = strSql + " 	and c.orderserial is NULL "
			'' strSql = strSql + " 	and o.orderserial is not NULL "
			'' ''rw strSql
			'' rsget.Open strSql, dbget, 1
		else
			ErrMsg = "�Ľ̿� �����߽��ϴ�."
		end if
		Set xmlDOM = Nothing
	end function

	'// ������
	function GetxSiteCSOrderCount_XXX()
		dim objNode, objNodes

		FResultCount = 0

		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		xmlDOM.LoadXML replace(objData,"&","��")

		if (FRectSellSite = "lotteCom") then
			'// �Ե�����
			FResultCount = xmlDOM.selectNodes("/Response/Result/OrderInfo").length
		else
			FResultCount = 0
		end if
		Set xmlDOM = Nothing
	end function

	public function GetXmlFromWeb()
		objData = ""

		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,60000,60000,90000  ''2013/08/01 �߰�
		objXML.Send()

		if objXML.Status = "200" then
			objData = BinaryToText(objXML.ResponseBody, "euc-kr")
		end if

		Set objXML  = Nothing
	end function

	public function GetXMLURL()
		dim tmp

		tmp = GetxSiteDateFormat(FRectStartYYYYMMDD)

		if (tmp = "") then
			GetXMLURL = ""
			ErrMsg = "���������� �������� �ʾҽ��ϴ�."
			exit function
		end if

		if (sellsite = "lotteCom") then
			if (FRectDivCD = "A008") then
				'// ���
				GetXMLURL = lotteAPIURL + "/openapi/searchCnclList.lotte?subscriptionId=" + CStr(lotteAuthNo) + "&start_date=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "&end_date=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD))
			elseif (FRectDivCD = "A004") then
				'// ��ǰ
				GetXMLURL = lotteAPIURL + "/openapi/searchReturnList.lotte?subscriptionId=" + CStr(lotteAuthNo) + "&start_date=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "&end_date=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD)) + "&ord_dtl_stat_cd=20"
			else
				GetXMLURL = ""
				ErrMsg = "��ϵ��� ���� ���޸��Դϴ�.[1]"
			end if
		elseif (sellsite = "lotteimall") then
			if (FRectDivCD = "sendsongjang") then
				'// ��������

				if application("Svr_Info")="Dev" then
					'// aaaaaaaaaaaaaaaaaaaaaaa
					'// ���߼���
					ltiMallAuthNo = "fe790295f0fec596ebc6a8a13a55513208e7830182501abf7b70d1fbc4e09ffde03afe430407f378238bccd00eda50718c4695904037247c5da9451d4f75dddc"
					ltiMallAPIURL = "http://openapi.lotteimall.com"
				end if

				'// sfin : ���Ϸ�
				'// dfin : ��ۿϷ�(������� ����)
				''GetXMLURL = ltiMallAPIURL + "/openapi/registDeliver.lotte?subscriptionId=" + CStr(ltiMallAuthNo) + "&ord_no=" + CStr(ord_no) + "&ord_dtl_sn=" + CStr(ord_dtl_sn) + "&proc_gubun=dfin&hdc_cd=" + CStr(hdc_cd) + "&inv_no=" + CStr(inv_no) + "&dlv_fin_dtime=" + CStr(GetxSiteDateFormat(sendDate))
				GetXMLURL = ltiMallAPIURL + "/openapi/registDeliver.lotte?subscriptionId=" + CStr(ltiMallAuthNo) + "&ord_no=" + CStr(ord_no) + "&ord_dtl_sn=" + CStr(ord_dtl_sn) + "&proc_gubun=sfin&hdc_cd=" + CStr(hdc_cd) + "&inv_no=" + CStr(inv_no) + "&dlv_fin_dtime=" + CStr(GetxSiteDateFormat(sendDate))
			elseif (FRectDivCD = "A008") then
				'// ���
				GetXMLURL = ltiMallAPIURL + "/openapi/searchCnclList.lotte?subscriptionId=" + CStr(ltiMallAuthNo) + "&start_date=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "&end_date=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD))
			elseif (FRectDivCD = "A004") then
				'// ��ǰ
				GetXMLURL = ltiMallAPIURL + "/openapi/searchReturnList.lotte?subscriptionId=" + CStr(ltiMallAuthNo) + "&start_date=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "&end_date=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD)) + "&ord_dtl_stat_cd=20"
			else
				GetXMLURL = ""
				ErrMsg = "��ϵ��� ���� ���޸��Դϴ�.[1]"
			end if
		else
			GetXMLURL = ""
			ErrMsg = "��ϵ��� ���� ���޸��Դϴ�.[2]"
		end if
	end function

	public function GetxSiteDateFormat(dt)
		if (FRectSellSite = "lotteCom") then
			GetxSiteDateFormat = Replace(dt, "-", "")
		elseif (FRectSellSite = "lotteimall") then
			GetxSiteDateFormat = Replace(dt, "-", "")
		else
			GetxSiteDateFormat = ""
		end if
	end function

	public function ResetXML()
		Set objXML = Nothing
		Set xmlDOM = Nothing
    end function

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FResultCount = 0
		FTotalCount = 0

		Call ResetXML()
	End Sub

	Private Sub Class_Terminate()
		Call ResetXML()
	End Sub

End Class

%>
