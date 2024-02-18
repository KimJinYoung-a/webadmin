<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������� [���]���� ����(�ν�) ����
' History : 2009.04.07 �̻� ����
'			2017.04.11 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->
<%
Dim params : params = request.Form("params")
Dim shopid : shopid = requestCheckVar(request.Form("shopid"),32)
Dim makerid : makerid = requestCheckVar(request.Form("makerid"),32)
Dim lossDate : lossDate = requestCheckVar(request.Form("lossDate"),30)
Dim cksel : cksel = request.Form("cksel") + ","
Dim AssignrealcheckErr : AssignrealcheckErr = request.Form("AssignrealcheckErr") + ","
Dim shopitemprice : shopitemprice = request.Form("shopitemprice") + ","
Dim shopbuyprice : shopbuyprice = request.Form("shopbuyprice") + ","
Dim shopsuplycash : shopsuplycash = request.Form("shopsuplycash") + ","
Dim itemgubun : itemgubun = request.Form("itemgubun") + ","
Dim itemid    : itemid = request.Form("itemid") + ","
Dim itemoption  : itemoption = request.Form("itemoption") + ","

Const CLOSSBRandID = "shopitemloss"
Dim sqlStr, idx, i, cnt, vix

rw shopid
rw makerid
rw cksel
rw AssignrealcheckErr
rw shopsuplycash

cksel     = split(cksel,",")
itemgubun = split(itemgubun,",")
itemid    = split(itemid,",")
itemoption= split(itemoption,",")
AssignrealcheckErr = split(AssignrealcheckErr,",")
shopsuplycash      = split(shopsuplycash,",")
shopbuyprice        = split(shopbuyprice,",")
shopitemprice       = split(shopitemprice,",")

''2���� ���� �ڷ�� �Է� ����..
Dim STOCKBASEDATE : STOCKBASEDATE = Left(dateAdd("m",-1,now()),7) + "-01" 
IF (CDate(lossDate)<CDate(STOCKBASEDATE)) THEN
   response.write STOCKBASEDATE & " ���� ��¥�� ���� �Ұ�"
   dbget.Close() : response.end
End if

if IsArray(cksel) then
    cnt = Ubound(cksel)
else
    cnt = 0
end if

''' �ν� ��� �Է�

''isreq �԰��û. Flag , isbaljuExists 'Y'
	sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_master"
	sqlStr = sqlStr + " (chargeid,shopid,divcode,vatcode,scheduledate,statecd,songjangdiv,songjangno,reguserid,isbaljuExists,comment)"
	sqlStr = sqlStr + " values('" + CLOSSBRandID + "'"
	sqlStr = sqlStr + " ,'" + shopid + "'"
	sqlStr = sqlStr + " ,'999'"
	sqlStr = sqlStr + " ,'008'"
	sqlStr = sqlStr + " ,'" + lossDate + "'"
	sqlStr = sqlStr + " ,0" 
	sqlStr = sqlStr + " ,'99'"
	sqlStr = sqlStr + " ,'�ν�ó��'"
	sqlStr = sqlStr + " ,'" + session("ssBctId") + "'"
	sqlStr = sqlStr + " ,'N'"
	sqlStr = sqlStr + " ,'�ǻ���� �ν�ó��'"
	sqlStr = sqlStr + " )"
	
	'response.write sqlStr &"<br>"
	dbget.Execute(sqlStr)

	sqlStr = " select ident_current('[db_shop].[dbo].tbl_shop_ipchul_master') as idx "
	rsget.Open sqlStr, dbget, 1
		idx = rsget("idx")
	rsget.close

	for i=0 to cnt
	    vix = cksel(i)
	    if (vix<>"") then
    	    If (itemgubun(vix)<>"") and (itemid(vix)<>"") and (itemoption(vix)<>"") and (shopsuplycash(vix)<>"") and (shopsuplycash(vix)<>"") then
        		sqlStr = " insert into [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
        		sqlStr = sqlStr + " (masteridx,itemgubun,shopitemid,itemoption," + vbCrlf
        		sqlStr = sqlStr + " designerid,sellcash,suplycash,shopbuyprice,itemno,reqno)"  + vbCrlf
        		sqlStr = sqlStr + " values(" + CStr(idx)  + "," + vbCrlf
        		sqlStr = sqlStr + "'" + requestCheckVar(Trim(itemgubun(vix)),2) + "'," + vbCrlf
        		sqlStr = sqlStr + "" + requestCheckVar(Trim(itemid(vix)),10) + "," + vbCrlf
        		sqlStr = sqlStr + "'" + requestCheckVar(Trim(itemoption(vix)),4) + "'," + vbCrlf
        		sqlStr = sqlStr + "'" + makerid + "'," + vbCrlf
        		sqlStr = sqlStr + "" + requestCheckVar(Trim(shopitemprice(vix)),20) + "," + vbCrlf
        		sqlStr = sqlStr + "" + requestCheckVar(Trim(shopsuplycash(vix)),20) + "," + vbCrlf
        		sqlStr = sqlStr + "" + requestCheckVar(Trim(shopbuyprice(vix)),20) + "," + vbCrlf
        		sqlStr = sqlStr + "" + requestCheckVar(Trim(AssignrealcheckErr(vix)*-1),10) + "," + vbCrlf
        		sqlStr = sqlStr + "" + requestCheckVar(Trim(AssignrealcheckErr(vix)*-1),10) + vbCrlf
        		sqlStr = sqlStr + "" + ")"
        		
        		dbget.Execute(sqlStr)
        		
        		'''���� ����
        		sqlStr = "exec [db_summary].[dbo].sp_Ten_Shop_realchekErr_Input '" & shopid & "','" & requestCheckVar(Trim(itemgubun(vix)),2) & "'," & requestCheckVar(Trim(itemid(vix)),10) & ",'" & requestCheckVar(Trim(itemoption(vix)),4) & "'," & requestCheckVar(AssignrealcheckErr(vix),10) & ",'" & lossDate & "','" & session("ssBctID") & "'"
                rw sqlStr
                dbget.Execute sqlStr
                    
            end if
        end if
		
	next

	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " set itemname=T.shopitemname" + vbCrlf
	sqlStr = sqlStr + " ,itemoptionname=T.shopitemoptionname" + vbCrlf
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_detail.masteridx=" + CStr(idx)
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.shopitemid=T.shopitemid"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_detail.itemoption=T.itemoption"

	dbget.Execute(sqlStr)

	'' Master Summary
	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNULL(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalshopbuyprice=IsNULL(T.totshopbuy,0)" + vbCrlf
	sqlStr = sqlStr + " ,statecd='8'"  + vbCrlf
	sqlStr = sqlStr + " ,execdt='"&lossDate&"'"  + vbCrlf
	sqlStr = sqlStr + " ,upcheconfirmdate=getdate()"  + vbCrlf
	sqlStr = sqlStr + " ,upcheconfirmuserid='" + session("ssBctId") + "'" + vbCrlf
	sqlStr = sqlStr + " ,lastupdate=getdate()" + vbCrlf
	sqlStr = sqlStr + " from (" + vbCrlf
	sqlStr = sqlStr + " select sum(sellcash*itemno) as totsell " + vbCrlf
	sqlStr = sqlStr + " ,sum(suplycash*itemno) as totsupp " + vbCrlf
	sqlStr = sqlStr + " ,sum(shopbuyprice*itemno) as totshopbuy " + vbCrlf
	sqlStr = sqlStr + " from " + vbCrlf
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " where masteridx="  + CStr(idx) + vbCrlf
	sqlStr = sqlStr + " and deleteyn='N'" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_master.idx=" + CStr(idx)

	dbget.Execute(sqlStr)

''��� �ݿ�.
    sqlStr = "exec db_summary.dbo.sp_Ten_Shop_BrandIpchulUpdateByIdx " & CStr(idx) & ",1"
    dbget.Execute(sqlStr)

    Dim retURL : retURL="/common/offshop/shopErrSummary.asp?"&params
    response.write "<script>alert('ó�� �Ǿ����ϴ�.');</script>"
    response.write "<script>location.href='"&retURL&"';</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
