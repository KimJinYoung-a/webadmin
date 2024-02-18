<%
Class CSimpleReportItem
	Public FYyyymmdd
	Public FMwOrderCnt
	Public FPcOrderCnt
	Public FAppOrderCnt
	Public FMwSumAmount
	Public FPcSumAmount
	Public FAppSumAmount

End Class

Class CSimpleReport
	Public FOneItem
	Public FItemList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

	Public FRectSDate
	Public FRectEDate
    Public FRectChannel
    public FRectReportType
    Public FRectOrderType
    Public FRectGroupType
    Public FRectDateGbn
    Public FRectAddParam1
    Public FRectAddParam2
    Public FRectchkvs
	public FRectitemid

'    Public FRectPType
'    Public FRectPValue
'    public FRectUPTypeValue
'    Public FRectPreTime
'    public FRectCompTerms
'    public FRectRdsiteGrp

    public FRectHH

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
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

    public function getSimpleReportStatistics(byref colRows)
        Dim strSql

	    if (vReportType="top100item") then
	        FRectEDate = DateAdd("d",-1,FRectEDate)
	        strSql = " exec [db_statistics].[dbo].[usp_TEN_Statistics_BestItem_Multi] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectOrderType&"',"&FPageSize&",'"&FRectchkvs&"'"
	    else
            'strSql = " exec db_EVT.[dbo].[sp_TEN_Conversion_daily_by_pType] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"'"
        end if

        if (strSql="") then Exit function
'  rw strSql
'response.end
        rsSTSget.CursorLocation = adUseClient
        rsSTSget.Open strSql,dbSTSget,adOpenForwardOnly,adLockReadOnly
		If not rsSTSget.EOF Then

		    colRows = Array()
		    For Each fld In rsSTSget.Fields
		        reDim Preserve colRows(UBound(colRows) + 1)
                colRows(UBound(colRows))=fld.Name

            Next

			getSimpleReportStatistics = rsSTSget.getRows()
		End If
		rsSTSget.Close
	end function

	public function getSimpleReport(byref colRows)
	    Dim strSql

	    if (vReportType="bestitemcoupon") then
	        if (FRectAddParam1<>"") then
	            strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_itemCoupon_Sales_ByCoupon] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectChannel&"','"&FRectOrderType&"','"&FRectAddParam1&"'"
	        else
	            strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_itemCoupon_Sales] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectChannel&"','"&FRectOrderType&"','"&FRectAddParam2&"'"
	        end if
	    elseif (vReportType="salesitemcpnbyuserlevel") then
	        strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_itemCoupon_Sales_ByLevel] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectChannel&"','"&FRectAddParam1&"'"
	    elseif (vReportType="itemcpnevalwithsales") then
	        strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_itemCoupon_Eval_Vs_Spend] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectChannel&"','"&FRectAddParam1&"','"&FRectOrderType&"'"
	    elseif (vReportType="outmallsales") then
	        if (FRectAddParam1<>"") then
	            strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_Outmall_Sales_BySitename] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectOrderType&"','"&FRectAddParam1&"'"
	        else
	            strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_Outmall_Sales] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectOrderType&"'"
	        end if
	    elseif (vReportType="evtsubscript") then
	        strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_Event_subscript] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"','"&FRectAddParam1&"'"
	    elseif (vReportType="rdsitesales") then
	        strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_rdsite_Sales] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectOrderType&"'"
	    elseif (vReportType="newitembycate") then
	        strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_NewItem_ByCate] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectChannel&"','"&FRectOrderType&"'"
	    elseif (vReportType="newitembybrandcate") then
	        strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_NewItem_ByBrandCate] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectChannel&"','"&FRectOrderType&"','"&FRectAddParam1&"'"
        elseif (vReportType="dealsales") then
	        if (FRectAddParam1<>"" and FRectitemid<>"") then
	            strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_Deal_Sales_ByDealCode_itemid] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectChannel&"','"&FRectOrderType&"',"&FRectAddParam1&","&FRectitemid&""
	        elseif FRectAddParam1<>"" or FRectitemid<>"" then
				if (FRectAddParam1<>"") then
					strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_Deal_Sales_ByDealCode] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectChannel&"','"&FRectOrderType&"',"&FRectAddParam1&""
				elseif (FRectitemid<>"") then
					strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_Deal_Sales_Byitemid] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectChannel&"','"&FRectOrderType&"',"&FRectitemid&""
				end if
	        else
	            strSql = " exec db_analyze_data_raw.[dbo].[usp_TEN_Analytics_Deal_Sales] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectDateGbn&"','"&FRectChannel&"','"&FRectOrderType&"'"
	        end if
	    else
            'strSql = " exec db_EVT.[dbo].[sp_TEN_Conversion_daily_by_pType] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"'"
        end if

        if (strSql="") then Exit function

		'response.write strSql & "<br>"
		'response.end
        rsAnalget.CursorLocation = adUseClient
		dbAnalget.CommandTimeout = 120  ''2�� (�⺻ 30��)
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then

		    colRows = Array()
		    For Each fld In rsAnalget.Fields
		        reDim Preserve colRows(UBound(colRows) + 1)
                colRows(UBound(colRows))=fld.Name

            Next

			getSimpleReport = rsAnalget.getRows()
		End If
		rsAnalget.Close

    end function

    public function x_getSimpleReport()
        Dim strSql , oData, fld, tempArray
        Set oData = Server.CreateObject("Scripting.Dictionary")

        strSql = " exec db_EVT.[dbo].[sp_TEN_Conversion_daily_by_pType] '"&FRectSDate&"', '"&FRectEDate&"','"&FRectChannel&"'"

        rsAnalget.CursorLocation = adUseClient
        rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
		If not rsAnalget.EOF Then
		    For Each fld In rsAnalget.Fields
                oData.Add fld.Name, Array()
            Next

			Do Until rsAnalget.EOF
                For Each fld In rsAnalget.Fields
                    tempArray = oData(fld.Name)
                    ReDim Preserve tempArray(UBound(tempArray) + 1)
                    tempArray(UBound(tempArray)) = rsAnalget(fld.Name)
                    oData(fld.Name) = tempArray
                Next
                rsAnalget.MoveNext
            Loop
		End If
		rsAnalget.Close

		SET getSimpleReport = oData
    end function

End Class


''-------------------------------------------
function drawReportDescription(iptype)
    dim ret
    Select Case iptype
        CASE "dealsales" :
            ret = "* �ֹ��� ����, �ֹ����� �̻�, ����, �ؿ� ����, 1�ð� ���� ������, ���� ���Ե� ��ǰ�� ������."
        CASE "top100item" :
            ret = "* ������ ����, 1�� ���� ������, ����α� ����."
        CASE ELSE
            ret = ""
    End Select

    response.write ret
end function

function drawReportSelectBox(iboxname,iselname)
    dim ret
    ret = "<select name='"&iboxname&"' >"
    ret = ret&"<option value='' "&CHKIIF(iselname="","selected","")&">����</option>"
    ret = ret&"<option value='bestitemcoupon' "&CHKIIF(iselname="bestitemcoupon","selected","")&">��ǰ��������</option>"
    ret = ret&"<option value='itemcpnevalwithsales' "&CHKIIF(iselname="itemcpnevalwithsales","selected","")&">��ǰ�������������</option>"

    'ret = ret&"<option value='salesitemcpnbyuserlevel' "&CHKIIF(iselname="salesitemcpnbyuserlevel","selected","")&">��ǰ��������-ȸ�����</option>"
    ret = ret&"<option value='outmallsales' "&CHKIIF(iselname="outmallsales","selected","")&">���޸�����</option>"
    ret = ret&"<option value='rdsitesales' "&CHKIIF(iselname="rdsitesales","selected","")&">RDSITE����</option>"
    ret = ret&"<option value='newitembycate' "&CHKIIF(iselname="newitembycate","selected","")&">ī�װ����Ż�ǰ(�ǸŽ�����)</option>"
    ret = ret&"<option value='newitembybrandcate' "&CHKIIF(iselname="newitembybrandcate","selected","")&">ī�װ����űԺ귣��</option>"

    ret = ret&"<option value='evtsubscript' "&CHKIIF(iselname="evtsubscript","selected","")&">�̺�Ʈ����������</option>"
    ret = ret&"<option value='dealsales' "&CHKIIF(iselname="dealsales","selected","")&">���������</option>"

    if (FALSE) then
        ret = ret&"<option value='sellbycategory' "&CHKIIF(iselname="sellbycategory","selected","")&">ī�װ�������</option>"
        ret = ret&"<option value='bestkeyword' "&CHKIIF(iselname="bestkeyword","selected","")&">����ƮŰ����</option>"

    end if
    ret = ret&"</select>"

    response.write ret
end function

function drawConversionChannelSelectBox(iboxname,iselname)
    dim ret
    ret = "<select name='"&iboxname&"' >"
    ret = ret&"<option value='' "&CHKIIF(iselname="","selected","")&">ALL</option>"
    ret = ret&"<option value='pc' "&CHKIIF(iselname="pc","selected","")&">WEB</option>"
    ret = ret&"<option value='mw' "&CHKIIF(iselname="mw","selected","")&">MOB</option>"
    ret = ret&"<option value='app' "&CHKIIF(iselname="app","selected","")&">APP</option>"
    ret = ret&"</select>"

    response.write ret
end function

function drawConversionTypeSelectBox(iboxname,iselname)
    dim ret
    ret = "<select name='"&iboxname&"' >"
    ret = ret&"<option value='' "&CHKIIF(iselname="","selected","")&">ALL</option>"
    ret = ret&"<option value='pRtr' "&CHKIIF(iselname="pRtr","selected","")&">�˻�</option>"
    ret = ret&"<option value='pBtr' "&CHKIIF(iselname="pBtr","selected","")&">�귣��</option>"
    ret = ret&"<option value='pEtr' "&CHKIIF(iselname="pEtr","selected","")&">�̺�Ʈ</option>"
    ret = ret&"<option value='pCtr' "&CHKIIF(iselname="pCtr","selected","")&">ī�װ�</option>"
    ret = ret&"<option value='rc' "&CHKIIF(iselname="rc","selected","")&">��ǰ��õ</option>"
    ret = ret&"<option value='gaparam' "&CHKIIF(iselname="gaparam","selected","")&">gaparam</option>"
    ret = ret&"</select>"

    response.write ret
end function

function drawConversionTypeGroupSelectBox(iboxname,iselname, iptype)
    dim ret, strSql, arrVal, i
    strSql = " exec [db_EVT].[dbo].[sp_TEN_conversion_get_comm_code] '"&iptype&"', ''"

    rsAnalget.CursorLocation = adUseClient
    rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
	If not rsAnalget.EOF Then
		arrVal = rsAnalget.getRows()
	End If
	rsAnalget.Close

	if isArray(arrVal) then
	    ret = "<select name='"&iboxname&"' >"
	    ret = ret&"<option value='' "&CHKIIF(iselname="","selected","")&">����</option>"
	    for i=0 To UBound(arrVal,2)
	        ret = ret&"<option value='"&arrVal(0,i)&"' "&CHKIIF(iselname=arrVal(0,i),"selected","")&">"&arrVal(0,i)&"</option>"
	    next
	    ret = ret&"</select>"
	end if
	response.write ret
end function

function drawConversionTypeGroupSelectBox2(iboxname,iselname,iptype,idepth,iupvalue)
    dim ret, strSql, arrVal, i
    strSql = " exec [db_EVT].[dbo].[sp_TEN_Conversion_get_comm_code_depth] '"&iptype&"',"&idepth&",'"&iupvalue&"'"
    rsAnalget.CursorLocation = adUseClient
    rsAnalget.Open strSql,dbAnalget,adOpenForwardOnly,adLockReadOnly
	If not rsAnalget.EOF Then
		arrVal = rsAnalget.getRows()
	End If
	rsAnalget.Close

	if isArray(arrVal) then
	    ret = "<select name='"&iboxname&"' >"
	    ret = ret&"<option value='' "&CHKIIF(iselname="","selected","")&">����</option>"
	    for i=0 To UBound(arrVal,2)
	        ret = ret&"<option value='"&arrVal(0,i)&"' "&CHKIIF(iselname=arrVal(0,i),"selected","")&">"&arrVal(1,i)&"</option>"
	    next

	    if (iselname="UNKNOWN") then
	        ret = ret&"<option value='UNKNOWN' selected>UNKNOWN</option>"
	    else
	        ret = ret&"<option value='UNKNOWN' >UNKNOWN</option>"
	    end if
	    ret = ret&"</select>"
	end if
	response.write ret
end function


function drawPreTimeSelectBox(iboxname,iselname)
    dim ret
    ret = "<select name='"&iboxname&"' >"
    ret = ret&"<option value='3' "&CHKIIF(iselname="3","selected","")&">�ֱ�3�ð�</option>"
    ret = ret&"<option value='6' "&CHKIIF(iselname="6","selected","")&">�ֱ�6�ð�</option>"
    ret = ret&"<option value='12' "&CHKIIF(iselname="12","selected","")&">�ֱ�12�ð�</option>"
    ret = ret&"<option value='24' "&CHKIIF(iselname="24","selected","")&">�ֱ�24�ð�</option>"
    ret = ret&"<option value='48' "&CHKIIF(iselname="48","selected","")&">�ֱ�48�ð�</option>"
    ret = ret&"</select>"

    response.write ret
end function

function getpTypeName(ipType)
    select CASE ipType
        CASE "pRtr"
            getpTypeName = "�˻���"
        CASE "pBtr"
            getpTypeName = "�귣��"
        CASE "pEtr"
            getpTypeName = "�̺�Ʈ"
        CASE "pCtr"
            getpTypeName = "ī�װ�"
        CASE "rc"
            getpTypeName = "��ǰ��õ"
        CASE "gaparam"
            getpTypeName = "gaparam"
        CASE ELSE
            getpTypeName = ".."
    end Select
end function
%>