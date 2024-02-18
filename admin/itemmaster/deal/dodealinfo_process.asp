<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###########################################################
' Page : /admin/itemmaster/deal/dodealinfo_process.asp
' Description :  �� ��ǰ - ���
' History : 2020.07.31 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<%
'--------------------------------------------------------
' �������� & �Ķ���� �� �ޱ�
'--------------------------------------------------------
Dim k, sqlStr, i, sailsdash, pricesdash, dealitemid
Dim startdate, enddate, shour, sminute, ehour, eminute
dim discountitemid, saleitemid, salePer, orgprice, sailprice, sellcash
dim realitemid

If request.form("shour")="" Then
    shour="00"
Else
    shour=request.form("shour")
End If
If request.form("sminute")="" Then
    sminute="00"
Else
    sminute=request.form("sminute")
End If
If request.form("ehour")="" Then
    ehour="23"
Else
    ehour=request.form("ehour")
End If
If request.form("eminute")="" Then
    eminute="59"
Else
    eminute=request.form("eminute")
End If
startdate = request.form("startdate") & " " & shour & ":" & sminute
enddate = request.form("enddate") & " " & ehour & ":" & eminute
sailsdash = request.form("sailsdash")
If sailsdash<>"Y" Then sailsdash="N"
pricesdash = request.form("pricesdash")
If pricesdash<>"Y" Then pricesdash="N"
discountitemid = request.form("discountitemid")
saleitemid = request.form("saleitemid")

'�̹� ��ϵ� ���ڵ����� �˻� (�ߺ� â���� ��� �� �� ����� ����� �ȵ� 2022.06.15 ������)
sqlStr = "select dealitemid from [db_event].[dbo].[tbl_deal_event] where idx='" & request.form("idx") & "'"
rsget.Open sqlStr, dbget, 1
if Not rsget.Eof then
    dealitemid = rsget("dealitemid")
end if
rsget.close
if dealitemid > 0 then
response.write "<script>alert('������ �� �ڵ�� ��ǰ�� ��ϵǾ����ϴ�.\n����Ʈ���� ��Ϲ�ư�� ���� �ٽ� ������ּ���.');</script>"
response.end
end if

If discountitemid<>"" Then
    '// ���� ���� ��ǰ ������ �������� //
    sqlStr =	"select orgprice, sailprice from [db_item].[dbo].tbl_item where itemid='" & discountitemid & "'"
    rsget.Open sqlStr, dbget, 1 
    if Not rsget.Eof then
        orgprice = rsget("orgprice")
        sailprice = rsget("sailprice")
    end if
    rsget.close
    If sailprice <> "0" Then
        salePer=Cint(((orgprice-sailprice)/orgprice)*100)
    Else
        salePer=0
    End IF
Else
    salePer=request.form("masterdiscountrate")
End If

If saleitemid<>"" Then
    '// ��������ǰ ������ �������� //
    sqlStr =	"select sellcash from [db_item].[dbo].tbl_item where itemid='" & saleitemid & "'"
    rsget.Open sqlStr, dbget, 1 
    if Not rsget.Eof then
        sellcash = rsget("sellcash")
    end if
    rsget.close
Else
    sellcash=request.form("mastersellcash")
End If

Public function GetDealItemList(byval masteridx)
    dim strSQL
    strSQL = "exec [db_event].[dbo].[sp_Ten_DealItemList] " & masteridx & ""
    rsget.CursorLocation = adUseClient
    rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
    If Not(rsget.EOF) then
        GetDealItemList = rsget.getRows
    end if
    rsget.Close
End Function

'=============================== �� �߰� ���� ==========================================
Dim ArrDealItem, intLoop, DealBrandName, DealBrandCheck, brandname, maxOrgprice
DealBrandCheck="Y"
DealBrandName=""
ArrDealItem=GetDealItemList(request.form("idx"))
If isArray(ArrDealItem) Then
    For intLoop = 0 To UBound(ArrDealItem,2)
        If intLoop=0 Then DealBrandName=ArrDealItem(7,intLoop)
        If ArrDealItem(7,intLoop) <> DealBrandName Then
            DealBrandCheck="N"
        End If
    Next
End If

If DealBrandCheck="N" Then
    brandname=""
Else
    brandname=DealBrandName
End If

'����ǰ ���� ���� ���� ���� �Ǵ�
Dim Isusing
If request.form("Isusing")="Y" Then
    Isusing="Y"
Else
    Isusing="N"
End If

'// Ʈ������ ����
''On Error Resume Next
dbget.beginTrans

'// ��ǰ��ȣ�� �޴´� //
    sqlStr = "Select isnull(max(itemid),0) as maxitemid  from [db_temp].[dbo].tbl_deal_item_temp"
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        realitemid = rsget("maxitemid") + 1
    end if
    rsget.close

'// ��ǰ ����� ���� ��� ���� ������ �޴´� //
    sqlStr = "Select isnull(max(i.orgprice),0) as orgprice"
    sqlStr = sqlStr & " from [db_event].[dbo].[tbl_deal_event_item] as d"
    sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item] as i on d.itemid=i.itemid"
    sqlStr = sqlStr & " where d.dealcode=" & request.form("idx")
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        maxOrgprice = rsget("orgprice")
    end if
    rsget.close

'// ��ǰ ������ �Է� //
    sqlStr =	"insert into [db_temp].[dbo].tbl_deal_item_temp" & vbCrlf &_
                "	(itemid,cate_large,cate_mid,cate_small,itemdiv " & vbCrlf &_
                " 		, makerid,frontMakerid,itemname " & vbCrlf &_
                "		, sellcash ,buycash, orgprice, orgsuplycash " & vbCrlf &_
                "		, mileage, sellyn, deliverytype " & vbCrlf &_
                "		, limityn,limitno,limitsold,limitdispyn,orderMinNum,orderMaxNum " & vbCrlf &_
                "		, vatinclude, pojangok, deliverarea, deliverfixday, mwdiv" & vbCrlf &_
                "		, itemscore, upchemanagecode, itemrackcode, tenOnlyYn, deliverOverseas , sellSTDate, brandname ,isusing" & vbCrlf &_
                "		, smallimage, listimage, listimage120, basicimage, basicimage600, icon1image, icon2image, optioncnt) " & vbCrlf &_
                "	 select top 1 " & realitemid & vbCrlf &_
                "	, cate_large, cate_mid, cate_small, 21, makerid, frontMakerid, convert(varchar(64),N'"& html2db(request.form("itemname")) &"') as itemname, " &sellcash&" ,buycash, " & maxOrgprice & ", orgsuplycash" & vbCrlf &_
                "	, mileage, 'Y', deliverytype, limityn, limitno, limitsold,'Y',orderMinNum,orderMaxNum, vatinclude, pojangok, deliverarea, deliverfixday, mwdiv" & vbCrlf &_
                "	, itemscore, upchemanagecode, itemrackcode, tenOnlyYn, deliverOverseas , getdate(), brandname, '" & Isusing & "'" & vbCrlf &_
                "	, smallimage, listimage, listimage120, basicimage, basicimage600, icon1image , icon2image, '" & salePer & "'" & vbCrlf &_
                "	from [db_item].[dbo].tbl_item where itemid='" & request.form("itemid") & "'"
    dbget.execute(sqlStr)

'// �� ī�װ� ���� : ��Ͻ� �⺻ 1 CateGory�� **//
    sqlStr = "Insert into [db_temp].dbo.tbl_deal_Item_category_temp" &_
            " (itemid,code_large,code_mid,code_small,code_div)" &_
            " select  top 1 " & realitemid & ", code_large,code_mid,code_small,code_div" &_
            "  from [db_item].dbo.tbl_Item_category WITH (READUNCOMMITTED) where itemid='" & Trim(request.form("itemid")) & "'"
'	Response.write sqlStr
'	Response.end
    dbget.execute(sqlStr)

    ''-- ��ǰ ������
    sqlStr = "insert into [db_temp].[dbo].tbl_deal_item_Contents_temp " & vbCrlf
    sqlStr = sqlStr & "(itemid, keywords, sourcearea, makername, " & vbCrlf
    sqlStr = sqlStr & " itemsource,itemsize,usinghtml,itemcontent, " & vbCrlf
    sqlStr = sqlStr & " ordercomment,designercomment, requireMakeDay,infoDiv,safetyYn,safetyDiv,safetyNum,freight_min,freight_max, sourcekind)"  & vbCrlf
    sqlStr = sqlStr & " select top 1 " & vbCrlf
    sqlStr = sqlStr & " "  & realitemid & vbCrlf
        sqlStr = sqlStr & " , '"&html2db(request.form("keywords"))&"', sourcearea, makername, itemsource,itemsize,usinghtml,itemcontent, ordercomment,designercomment, requireMakeDay,infoDiv,safetyYn,safetyDiv,safetyNum,freight_min,freight_max, sourcekind" & vbCrlf
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_item_Contents where itemid='" & request.form("itemid") & "'" & vbCrlf
'Response.write sqlStr
'	Response.end
    dbget.execute(sqlStr)

    '// ����ī�װ� �ֱ� //
    If (request.form("catecode").Count>0) Then
        sqlStr = "delete from db_temp.dbo.tbl_deal_display_cate_item_temp Where itemid='" & realitemid & "';" & vbCrLf
        sqlStr = sqlStr & "update db_temp.dbo.tbl_deal_item_temp set dispcate1=null Where itemid='" & realitemid & "';" & vbCrLf
        for i=1 to request.form("catecode").Count
            sqlStr = sqlStr & "Insert into db_temp.dbo.tbl_deal_display_cate_item_temp (catecode, itemid, depth, sortNo, isDefault) values "
            sqlStr = sqlStr & "('" & request.form("catecode")(i) & "'"
            sqlStr = sqlStr & ",'" & realitemid & "'"
            sqlStr = sqlStr & ",'" & request.form("catedepth")(i) & "',9999"
            sqlStr = sqlStr & ",'" & request.form("isDefault")(i) & "');" & vbCrLf
            if request.form("isDefault")(i)="y" then
                sqlStr = sqlStr & "update db_temp.dbo.tbl_deal_item_temp set dispcate1='" & left(request.form("catecode")(i),3) & "' Where itemid='" & realitemid & "';" & vbCrLf
            end if
        next
        dbget.execute(sqlStr)
    end if

    '####### PC ��ǰ�����̹��� �� (�ִ� 7������ ����) 20150603 #######
    If request.form("dealcontents") <> "" Then
        sqlStr = " IF Not Exists(SELECT IDX FROM [db_temp].[dbo].tbl_deal_item_addimage_temp WHERE ITEMID='" & realitemid & "' and IMGTYPE=1 and GUBUN=1)"
        sqlStr = sqlStr + "	BEGIN "
        sqlStr = sqlStr+ " 		INSERT INTO [db_temp].[dbo].tbl_deal_item_addimage_temp (ITEMID,IMGTYPE,GUBUN,ADDIMAGE_400)"
        sqlStr = sqlStr + "     	VALUES ('" & realitemid & "',1,1,'" & request.form("dealcontents") & "')"
        sqlStr = sqlStr + " 	END"
        sqlStr = sqlStr + " ELSE"
        sqlStr = sqlStr + " 	BEGIN "
        sqlStr = sqlStr + "		UPDATE [db_temp].[dbo].tbl_deal_item_addimage_temp "
        sqlStr = sqlStr + " 		SET ADDIMAGE_400 ='" & request.form("dealcontents") & "'"
        sqlStr = sqlStr + " 		WHERE ITEMID = '" & realitemid & "'"
        sqlStr = sqlStr + " 		and IMGTYPE=1"
        sqlStr = sqlStr + " 		and GUBUN =1"
        sqlStr = sqlStr + " 	END "
        dbget.execute sqlStr
    End If
    '####### ����� ��ǰ�����̹��� �� (�ִ� 7������ ����) 20150603 #######
    If request.form("mobiledealcontents") <> "" Then
        sqlStr = " IF Not Exists(SELECT IDX FROM [db_temp].[dbo].tbl_deal_item_addimage_temp WHERE ITEMID='" & realitemid & "' and IMGTYPE=1 and GUBUN=2)"
        sqlStr = sqlStr + "	BEGIN "
        sqlStr = sqlStr+ " 		INSERT INTO [db_temp].[dbo].tbl_deal_item_addimage_temp (ITEMID,IMGTYPE,GUBUN,ADDIMAGE_400)"
        sqlStr = sqlStr + "     	VALUES ('" & realitemid & "',1,2,'" & request.form("mobiledealcontents") & "')"
        sqlStr = sqlStr + " 	END"
        sqlStr = sqlStr + " ELSE"
        sqlStr = sqlStr + " 	BEGIN "
        sqlStr = sqlStr + "		UPDATE [db_temp].[dbo].tbl_deal_item_addimage_temp "
        sqlStr = sqlStr + " 		SET ADDIMAGE_400 ='" & request.form("mobiledealcontents") & "'"
        sqlStr = sqlStr + " 		WHERE ITEMID = '" & realitemid & "'"
        sqlStr = sqlStr + " 		and IMGTYPE=1"
        sqlStr = sqlStr + " 		and GUBUN =2"
        sqlStr = sqlStr + " 	END "
        dbget.execute sqlStr
    End If

'####################### �� ��ǰ ���� ���� ################################
sqlStr =  "UPDATE [db_event].[dbo].[tbl_deal_event]"
sqlStr = sqlStr + " 		SET status=0"
sqlStr = sqlStr + " 		, dealitemid ='" & realitemid & "'"
sqlStr = sqlStr + " 		, masteritemcode ='" & request.form("itemid") & "'"
sqlStr = sqlStr + " 		, viewdiv ='" & request.form("viewdiv") & "'"
sqlStr = sqlStr + " 		, startdate ='" &startdate & "'"
sqlStr = sqlStr + " 		, enddate ='" & enddate & "'"
sqlStr = sqlStr + " 		, mastersellcash ='" & request.form("mastersellcash") & "'"
sqlStr = sqlStr + " 		, masterdiscountrate ='" & request.form("masterdiscountrate") & "'"
sqlStr = sqlStr + " 		, regname ='" & Cstr(request.form("auser")) & "'"
sqlStr = sqlStr + " 		, pricesdash ='" & pricesdash & "'"
sqlStr = sqlStr + " 		, sailsdash ='" & sailsdash & "'"
sqlStr = sqlStr + " 		, work_notice ='" & request.form("work_notice") & "'"
sqlStr = sqlStr + " 		, mainTitle ='" & request.form("mainTitle") & "'"
sqlStr = sqlStr + " 		, subTitle ='" & request.form("subTitle") & "'"
sqlStr = sqlStr + " 		WHERE idx = '" & request.form("idx") & "' "
dbget.execute sqlStr

'####################### ��ǰ ���� �α� ���� ################################
sqlStr =  "insert db_log.dbo.tbl_NsqMesQue(title,topic,memo,ownername,callip,message,result)"
sqlStr = sqlStr + " values('��ǰ���','DEAL_PRODUCT','','" & session("ssBctId") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & realitemid & "','dealready')"
dbget.execute sqlStr

If Err.Number = 0 Then
    dbget.CommitTrans
    if (application("Svr_Info")="Dev") then
        '�׽�Ʈ ������ API ȣ�� ���� ����
        message = "�Ϸ� �Ǿ����ϴ�.\n����Ʈ������ �̵� �� API ȣ��� ��ǰ ���� �������ּ���."
    else
        dim message, oXML
        '####################### ���� API ȣ�� ################################
        set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP���۳�Ʈ ����
        oXML.open "GET", "http://110.93.128.100:8090/scmapi/nsqmessage/containcollect", false
        oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
        oXML.send	'����
        'response.write oXML.responseText & "<br>"
        if oXML.status=200 then
            message = "��� �Ϸ�"
        else
            message = "��� ����[001]"
        end if
        Set oXML = Nothing	'���۳�Ʈ ����

        IF (Err) then
            message = "���� ����[002]"
        end if
    end if
%>
<script type="text/javascript">
$(function() {
	alert("<%=message%>");
    location.replace("index.asp");

});
</script>
<%
else
    dbget.RollBackTrans
%>
<script type="text/javascript">
alert("ó���� ������ �߻��߽��ϴ�.");
</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->