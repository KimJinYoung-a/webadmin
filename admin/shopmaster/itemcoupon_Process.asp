<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dblogicsopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
 Server.ScriptTimeout= 150	'5��

Dim otimer : otimer=Timer()
Dim pTimer : pTimer=otimer

function debugRwite (istep)
	if Application("Svr_info")="Dev" then
		rw istep&":"& FormatNumber(Timer()-pTimer,4)
		pTimer=Timer()
	end if
end function

function IsLastUpdateNotAssign(itemcouponidx)
    ''������ ���� ������ ������ 2000�� �̻��ϰ�� lastupdate�� ������Ʈ ���ϱ� ����.
    Dim sqlStr

    IsLastUpdateNotAssign = false

    sqlStr = "select count(*) as CNT from [db_item].[dbo].tbl_item_coupon_detail where itemcouponidx=" + CStr(itemcouponidx)
    rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
	    IsLastUpdateNotAssign =(rsget("CNT")>=2000)   ''��ǰ���� 2000�� ������ ����/����� lastupdate ó�� ����.
	end if
	rsget.Close

end function

function AplyItemCountUpdate(itemcouponidx)
	dim sqlStr
	''�����ǰ���� ������Ʈ
	sqlStr = "update [db_item].[dbo].tbl_item_coupon_master" + VbCrlf
	sqlStr = sqlStr + " set applyitemcount=IsNULL(T.cnt,0)" + VbCrlf
	sqlStr = sqlStr + ",lastupDt=getdate()" + VbCrlf
	sqlStr = sqlStr + " from (" + VbCrlf
	sqlStr = sqlStr + "     select count(*) as cnt from [db_item].[dbo].tbl_item_coupon_detail where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
	sqlStr = sqlStr + " ) as T" + VbCrlf
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf

	dbget.Execute sqlStr
	
	''2������ 2018/06/19 
	sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
	dblogicsget.Execute sqlStr
end function

function AplyToItem(itemcouponidx, chklastupdate)
	dim sqlStr
	dim ocouponGubun, oitemcoupontype, oitemcouponvalue, oitemcouponstartdate, oitemcouponexpiredate, openstate, currdatetime
	dim couponExpired
	dim resultCnt
    dim notUpdate

    '' 2010-10 �߰�
    ''��ǰ lastupdate �������� ����
    ''chklastupdate �� ���� ���� ���½�/ ��������ø� üũ��.
    ''notUpdate = true�ΰ�� lastupdate ������Ʈ ���� ����.
    notUpdate = false
    if (chklastupdate) then
        notUpdate = IsLastUpdateNotAssign(itemcouponidx)

        if (notUpdate) then response.write "�����ǰ���� 2000���̻��̶� ��ǰ lastupdate �������.<br>"
    end if
	Call debugRwite("stepA-1")

	applyitemcount = 0
	couponExpired = false

	sqlStr = "select top 1 couponGubun, margintype, itemcoupontype, itemcouponvalue, openstate, applyitemcount,"
	sqlStr = sqlStr + " convert(varchar(19),itemcouponstartdate,21) as itemcouponstartdate,"
	sqlStr = sqlStr + " convert(varchar(19),itemcouponexpiredate,21) as itemcouponexpiredate,"
	sqlStr = sqlStr + " (case when (itemcouponstartdate>getdate()) or (itemcouponexpiredate<getdate()) then 'Y' else 'N' end ) as couponexpired, "
	sqlStr = sqlStr + " convert(varchar(19),getdate()) as currdatetime"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_master" + VbCrlf
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx)

	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
	    ocouponGubun   = rsget("couponGubun")
		itemcoupontype = rsget("itemcoupontype")
		itemcouponvalue = rsget("itemcouponvalue")
		itemcouponstartdate = rsget("itemcouponstartdate")
		itemcouponexpiredate = rsget("itemcouponexpiredate")
		openstate = rsget("openstate")
		applyitemcount = rsget("applyitemcount")
		currdatetime = rsget("currdatetime")

		couponExpired = rsget("couponexpired")

		response.write "couponExpired :" + CStr(couponExpired) + "<br>"
	end if
	rsget.Close
	Call debugRwite("stepA-2")
    
	''�߱޴�����̰ų� �߱޿������ ��ŵ.
	if ((openstate<>"7") and (openstate<>"9")) then exit function

	'' Naver �����ΰ�� lastupdate  2018/08/08
	if (ocouponGubun="V") and (Not notUpdate) then
		''��ġ�� ó���ϴ°����� ��������.
		''EXEC db_AppWish.dbo.[sp_TEN_CP_tbl_item_coupon_master_detail_Change]
		' sqlStr = "EXEC db_AppWish.dbo.[sp_TEN_CP_tbl_item_coupon_master_detail_Change_By_CPnIDX] "&CStr(itemcouponidx)
		' dblogicsget.Execute sqlStr

		' ''2������ 2018/06/19 
		' sqlStr = "update I "
		' sqlStr = sqlStr + " set lastupdate=getdate()"
		' sqlStr = sqlStr + " from [db_AppWish].[dbo].tbl_item I"
		' sqlStr = sqlStr + " 	Join [db_AppWish].[dbo].tbl_item_coupon_detail d"
		' sqlStr = sqlStr + " 	on I.itemid=d.itemid"
		' sqlStr = sqlStr + " where d.itemcouponidx=" + CStr(itemcouponidx)
		
		' dblogicsget.Execute sqlStr
	end if

	''Ÿ������, ��������, ����������ΰ�� ��ŵ.
    if (ocouponGubun<>"C") then exit function

	''�߱� ����� �����ΰ�� -> N�� ����
	if (openstate="9") or (couponExpired="Y") then

		sqlStr = "update [db_item].[dbo].tbl_item"
		sqlStr = sqlStr + " set itemcouponyn='N'"
		sqlStr = sqlStr + " ,itemcoupontype='1'"
		sqlStr = sqlStr + " ,itemcouponvalue=0"
		sqlStr = sqlStr + " ,curritemcouponidx=NULL"
		IF (Not notUpdate) then
		    sqlStr = sqlStr + " ,lastupdate=getdate()"
	    end if
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_detail"
		sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx)
		sqlStr = sqlStr + " and [db_item].[dbo].tbl_item.itemid=[db_item].[dbo].tbl_item_coupon_detail.itemid"

		'response.write sqlStr + "<br>"
		dbget.Execute sqlStr
		Call debugRwite("stepA-3")		

		''2������ 2018/06/19 
	    sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
	    dblogicsget.Execute sqlStr
		Call debugRwite("stepA-4")
	end if

	''��ǰ�� �����Ȱ�� -> N�� ���� // �����Ѱ��� lastupdate����
	sqlStr = "update [db_item].[dbo].tbl_item"
	sqlStr = sqlStr + " set itemcouponyn='N'"
	sqlStr = sqlStr + " ,itemcoupontype='1'"
	sqlStr = sqlStr + " ,itemcouponvalue=0"
	sqlStr = sqlStr + " ,curritemcouponidx=NULL"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select i.itemid  "
	sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item i"
	sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_item_coupon_detail d"
	sqlStr = sqlStr + " 	on d.itemcouponidx=" + CStr(itemcouponidx) + " and i.itemid=d.itemid "
	sqlStr = sqlStr + " 	where i.curritemcouponidx=" + CStr(itemcouponidx)
	sqlStr = sqlStr + " 	and d.itemcouponidx is null"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=T.itemid"

	'response.write sqlStr + "<br>"
		dbget.Execute sqlStr, resultCnt
	response.write "�����Ǽ�=" + CStr(resultCnt) + "<br>"
	Call debugRwite("stepA-5")
    
	''2������ 2018/06/19 
    sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
    dblogicsget.Execute sqlStr
	Call debugRwite("stepA-6") ''���Ⱑ ����.

	''itemcouponidx�� ��ϵ� ��ǰ�� ��� �������� ������ Update
	sqlStr = "update [db_item].[dbo].tbl_item"
	sqlStr = sqlStr + " set itemcouponyn='Y'"
	sqlStr = sqlStr + " ,itemcoupontype=T.itemcoupontype"
	sqlStr = sqlStr + " ,itemcouponvalue=T.itemcouponvalue"
	sqlStr = sqlStr + " ,curritemcouponidx=T.itemcouponidx"
	IF (Not notUpdate) then
	    sqlStr = sqlStr + " ,lastupdate=getdate()"
    end if
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select m.itemcouponidx, m.itemcoupontype, m.itemcouponvalue, d.itemid "
	sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item_coupon_master m,"
	sqlStr = sqlStr + " 	[db_item].[dbo].tbl_item_coupon_detail d"
	sqlStr = sqlStr + " 	where m.itemcouponidx=d.itemcouponidx"
	sqlStr = sqlStr + " 	and m.openstate='7'"
	sqlStr = sqlStr + " 	and d.itemcouponidx=" + CStr(itemcouponidx)
	sqlStr = sqlStr + " 	and m.itemcouponstartdate<=getdate()"
	sqlStr = sqlStr + " 	and m.itemcouponexpiredate>=getdate()"
	sqlStr = sqlStr + " ) T "
	sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=T.itemid"
	sqlStr = sqlStr + " and Not ("
	sqlStr = sqlStr + " 		 	[db_item].[dbo].tbl_item.itemcouponyn='Y'"
	sqlStr = sqlStr + " 		and [db_item].[dbo].tbl_item.itemcoupontype=T.itemcoupontype"
	sqlStr = sqlStr + " 		and [db_item].[dbo].tbl_item.itemcouponvalue=T.itemcouponvalue"
	sqlStr = sqlStr + " 		and [db_item].[dbo].tbl_item.curritemcouponidx=T.itemcouponidx"
	sqlStr = sqlStr + "			)"

	'response.write sqlStr + "<br>"
	dbget.Execute sqlStr, resultCnt
	Call debugRwite("stepA-7")
    response.write "�����Ǽ�=" + CStr(resultCnt)
    
    ''2������ 2018/06/19 
    sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
    dblogicsget.Execute sqlStr
	Call debugRwite("stepA-8")
end function

'### ���� �α� ����
Sub AddSCMChangeLog(couponIdx,logMessage)
	Dim strSql
	if logMessage<>"" then
		strSql = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
		strSql = strSql & "VALUES('" & session("ssBctId") & "', 'itemCoupon', '" & couponIdx & "', '" & requestCheckVar(Request("menupos"),9) & "', "
		strSql = strSql & "'" & logMessage & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
		dbget.execute(strSql)
	end if
End Sub

dim refer
refer = request.ServerVariables("HTTP_REFERER")


dim itemcouponidx
dim couponGubun
dim itemcoupontype
dim itemcouponvalue
dim itemcouponstartdate
dim itemcouponexpiredate
dim itemcouponname
dim itemcouponimage
dim itemcouponexplain
dim applyitemcount
dim openstate
dim margintype
dim defaultmargin
dim mode
dim IsEditMode
dim sqlstr,i
dim buf
dim itemidarr, couponbuypricearr, couponsellcasharr, makerid, sailyn
dim ErrStr

dim sType, addSql, itemid, itemname, sellyn, usingyn, danjongyn, disp, couponyn, minmargin, itemcostup, itemcostdown
dim limityn, mwdiv, cdl, cdm, cds, deliverytype, coupontype, groupId
dim itemcouponidxarr
dim exceptnotepmapitem

itemcouponidx      	= requestCheckVar(request("itemcouponidx"),9)
couponGubun         = requestCheckVar(request("couponGubun"),9)
itemcoupontype      = requestCheckVar(request("itemcoupontype"),9)
itemcouponvalue     = requestCheckVar(request("itemcouponvalue"),9)
itemcouponstartdate = request("itemcouponstartdate") + " " + request("itemcouponstartdate2")
itemcouponexpiredate= request("itemcouponexpiredate") + " " + request("itemcouponexpiredate2")
itemcouponname      = html2Db(request("itemcouponname"))
itemcouponimage     = request("itemcouponimage")
applyitemcount      = request("applyitemcount")
openstate         	= request("openstate")
margintype          = request("margintype")
defaultmargin		= request("defaultmargin")
mode 				= request("mode")
itemidarr			= request("itemidarr")
couponbuypricearr	= request("couponbuypricearr")
couponsellcasharr   = request("couponsellcasharr")
itemcouponexplain	= html2Db(request("itemcouponexplain"))
makerid				= request("makerid")
sailyn				= request("sailyn")
sType               = request("sType")

addSql              = request("addSql")
itemid              = request("itemid")
itemname            = request("itemname")
sellyn              = request("sellyn")
usingyn             = request("usingyn")
danjongyn           = request("danjongyn")
limityn             = request("limityn")
mwdiv               = request("mwdiv")
cdl                 = request("cdl")
cdm                 = request("cdm")
cds                 = request("cds")
deliverytype        = request("deliverytype")
coupontype			= requestCheckVar(request("coupontype"),1)
itemcouponidxarr    = request("itemcouponidxarr")
disp                = requestCheckVar(request("disp"),30)
couponyn            = requestCheckVar(request("couponyn"),10)
minmargin           = requestCheckVar(request("minmargin"),10)
itemcostup          = requestCheckVar(request("itemcostup"),10)
itemcostdown        = requestCheckVar(request("itemcostdown"),10)
exceptnotepmapitem	= requestCheckVar(request("exceptnotepmapitem"),10)
groupId				= requestCheckVar(request("groupId"),8)

if itemcouponidx="" then itemcouponidx="0"
if defaultmargin="" then defaultmargin=0
if coupontype="" then coupontype="N"
if (itemcouponidx<>"0") then
	IsEditMode = true
else
	IsEditMode = false
end if

if mode="couponmaster" then
	on Error Resume Next
		buf = CDate(itemcouponstartdate)
		if Err then
			response.Write "<script>alert('�߱޽����� ����-" + Err.Description + "')</script>"
			response.Write "<script>history.back()</script>"
			dbget.close()	:	response.End
		end if
	on Error Goto 0

	on Error Resume Next
		buf = CDate(itemcouponexpiredate)
		if Err then
			response.Write "<script>alert('�߱������� ����-" + Err.Description + "')</script>"
			response.Write "<script>history.back()</script>"
			dbget.close()	:	response.End
		end if
	on Error Goto 0

	if (itemcoupontype="1") then
		if (itemcouponvalue>=100) or (itemcouponvalue<1) then
			response.Write "<script>alert('���������� 1~99% ���� ���� �����մϴ�.')</script>"
			response.Write "<script>history.back()</script>"
			dbget.close()	:	response.End
		end if
	elseif (itemcoupontype="2") Then
		If session("ssBctId")="bborami" Then '// ���� �̺�Ʈ ���� �Է¶����� �Ӻ��� �븮�� ���� Ǯ��� ��, ��¥���� �ɾ��..
			If Left(Now(), 10) >= "2015-05-12" And Left(Now(), 10) < "2015-05-27" Then
			Else
				if (itemcouponvalue<100) or (itemcouponvalue>=150001) then  ''100000 => 150000 ������ ��û
					response.Write "<script>alert('���������� 1~150000 ���� ���� �����մϴ�.')</script>"
					response.Write "<script>history.back()</script>"
					dbget.close()	:	response.End
				end If
			End If
		Else
			if (itemcouponvalue<100) or (itemcouponvalue>=300001) then  ''150000 => 300000 �豤�� ��û 150708
				response.Write "<script>alert('���������� 1~300000 ���� ���� �����մϴ�.')</script>"
				response.Write "<script>history.back()</script>"
				dbget.close()	:	response.End
			end If
		End If
	elseif (itemcoupontype="3") then
		if (Cint(itemcouponvalue)<>Cint(getDefaultBeasongPayByDate(now()))) then
			response.Write "<script>alert('������ ���������� " + Cstr(getDefaultBeasongPayByDate(now())) + " ���� �����մϴ�.')</script>"
			response.Write "<script>history.back()</script>"
			dbget.close()	:	response.End
		end if
	else
		response.Write "<script>alert('����Ÿ���� �������� �ʾҽ��ϴ�.')</script>"
		response.Write "<script>history.back()</script>"
		dbget.close()	:	response.End
	end if


	if (IsEditMode) then
		''����
		dim orgDefaultMargin ,orgDefaultMargintype
		sqlstr = "SELECT defaultmargin,margintype FROM db_item.dbo.tbl_item_coupon_master "
		sqlstr = sqlstr + " where itemcouponidx=" + CStr(itemcouponidx)

		rsget.CursorLocation = adUseClient
        rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly

		IF not rsget.eof Then
			orgDefaultMargin = rsget("defaultmargin")
			orgDefaultMargintype = rsget("margintype")
		End IF

		rsget.close

		sqlstr = "update [db_item].[dbo].tbl_item_coupon_master" + VbCrlf
		sqlstr = sqlstr + " set itemcoupontype='" + itemcoupontype + "'" + VbCrlf
		sqlstr = sqlstr + " ,couponGubun='" + couponGubun + "'" + VbCrlf
		sqlstr = sqlstr + " ,itemcouponvalue=" + CStr(itemcouponvalue) + VbCrlf
		sqlstr = sqlstr + " ,itemcouponstartdate='" + itemcouponstartdate + "'" + VbCrlf
		sqlstr = sqlstr + " ,itemcouponexpiredate='" + itemcouponexpiredate + "'" + VbCrlf
		sqlstr = sqlstr + " ,itemcouponname='" + itemcouponname + "'" + VbCrlf
		sqlstr = sqlstr + " ,itemcouponexplain='" + itemcouponexplain + "'" + VbCrlf
		sqlstr = sqlstr + " ,margintype='" + margintype + "'" + VbCrlf
		sqlstr = sqlstr + " ,defaultmargin='" + defaultmargin + "'" + VbCrlf
		sqlstr = sqlstr + " ,coupontype='" + coupontype + "'" + VbCrlf
		sqlStr = sqlStr + ",lastupDt=getdate()" + VbCrlf
		sqlStr = sqlStr + ", itemcouponimage='" + itemcouponimage + "'" + VbCrlf
		sqlstr = sqlstr + " where itemcouponidx=" + CStr(itemcouponidx)

		dbget.Execute sqlStr

        ''2������ 2018/06/19 
	    sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
	    dblogicsget.Execute sqlStr
	    
		'���� ���� ����� ��� ��ǰ ��ü ����
		IF (Cint(orgDefaultMargin) <> Cint(defaultmargin)) or (CStr(orgDefaultMargintype)<>CStr(margintype)) Then
				sqlStr =" UPDATE [db_item].[dbo].tbl_item_coupon_detail  "& VbCRLF
				sqlStr = sqlStr& " SET couponbuyprice="& VbCRLF
				SELECT Case margintype
					Case "00"  	''��ǰ�������� - ���԰� 0 �ΰ�� �����԰�
						sqlStr = sqlStr & " 0 " & VbCrlf
					Case "10"	''�ٹ����ٺδ� - �����԰�
						sqlStr = sqlStr & " 0 " & VbCrlf
					Case "20"	''�������� : �߰� [2008-09-23]
						if itemcoupontype="1" then			''������
							sqlStr = sqlStr & " convert(int,i.sellcash*"& Cstr((100-itemcouponvalue)/100) &"*"& Cstr((100-defaultmargin)/100) &")"
						elseif itemcoupontype="2" then   	''�ݾ�
							sqlStr = sqlStr & " convert(int,(i.sellcash-" & CStr(itemcouponvalue) + ")*"& Cstr((100-defaultmargin)/100) &")"
						else
							sqlStr = sqlStr & " 0 " & VbCrlf
						end if
					Case "30"	''���ϸ��� - ���縶�� : �߰� [2008-09-23]
						if itemcoupontype="1" then			''������
							sqlStr = sqlStr & " convert(int,i.sellcash*" & CStr((100-itemcouponvalue)/100) & "*i.buycash/i.sellcash)"
						elseif itemcoupontype="2" then   	''�ݾ�
							sqlStr = sqlStr & " convert(int,(i.sellcash-" & CStr(itemcouponvalue) & ")*i.buycash/i.sellcash)"
						else
							sqlStr = sqlStr & " 0 " & VbCrlf
						end if
					Case "50"	''�ݹݺδ�
						if itemcoupontype="1" then			''������
							sqlStr = sqlStr & " i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & "*0.5)"
						elseif itemcoupontype="2" then   	''�ݾ�
							sqlStr = sqlStr & " i.buycash - convert(int," & CStr(itemcouponvalue) & "*0.5)"
						else
							sqlStr = sqlStr & " 0 " & VbCrlf
						end if
					Case "60"	''��ü�δ� - ���԰� ����
						if itemcoupontype="1" then			''������
							sqlStr = sqlStr & " i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) + ")"
						elseif itemcoupontype="2" then   	''�ݾ�
							sqlStr = sqlStr & " i.buycash - " & CStr(itemcouponvalue)
						else
							sqlStr = sqlStr & " 0 " & VbCrlf
						end if
			        Case "80"   ''���������� -500
			            sqlStr = sqlStr & " case when i.mwdiv='M' then 0 else i.buycash - 500 end "
					Case "90"	''20%��ü��� - �����ΰ�� �����԰�.
						if itemcoupontype="1" then			''������
							sqlStr = sqlStr & " case when i.mwdiv='M' then 0 else i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & "*0.5) end "
						elseif itemcoupontype="2" then   	''�ݾ�
							sqlStr = sqlStr & " case when i.mwdiv='M' 0 else i.buycash - convert(int," & CStr(itemcouponvalue) & "*0.5)  end "
						else
							sqlStr = sqlStr & " 0 " & VbCrlf
						end if
					Case else
						sqlStr = sqlStr & " 0 " & VbCrlf
				End SELECT
						sqlStr = sqlStr & " , couponmargin=" 
				SELECT Case margintype
					Case "00"  	''��ǰ�������� - ���԰� 0 �ΰ�� �����԰� 
            		    sqlStr = sqlStr & " 0" & VbCrlf
            		Case "10"	''�ٹ����ٺδ� - �����԰� 
                        sqlStr = sqlStr & " 0" & VbCrlf
            		Case "20"	''�������� : �߰� [2008-09-23] 
            			if itemcoupontype="1" then			''������ 
                            sqlStr = sqlStr & " (((i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)) - (convert(int,i.sellcash*"& Cstr((100-itemcouponvalue)/100) &"*"& Cstr((100-defaultmargin)/100) &")))/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100"
            			elseif itemcoupontype="2" then   	''�ݾ� 
            				sqlStr = sqlStr & " (((i.sellcash-"&Cstr(itemcouponvalue)&") -(convert(int,(i.sellcash-" & CStr(itemcouponvalue) & ")*"& Cstr((100-defaultmargin)/100) &")))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100"
            			else 
            				sqlStr = sqlStr & " 0 " & VbCrlf
            			end if
            		Case "30"	''���ϸ��� - ���縶�� : �߰� [2008-09-23]
            			if itemcoupontype="1" then			''������ 
            				sqlStr = sqlStr & "  (((i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)) - (convert(int,i.sellcash*" & CStr((100-itemcouponvalue)/100) & "*i.buycash/i.sellcash)))/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100 "
            			elseif itemcoupontype="2" then   	''�ݾ� 
            				sqlStr = sqlStr & "  (((i.sellcash-"&Cstr(itemcouponvalue)&") - (convert(int,(i.sellcash-" & CStr(itemcouponvalue) & ")*i.buycash/i.sellcash)))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100 "
            			else 
            				sqlStr = sqlStr & " 0 " & VbCrlf
            			end if
            		Case "50"	''�ݹݺδ�
            			if itemcoupontype="1" then			''������ 
            				sqlStr = sqlStr  & "  (((i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)) - ( i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & "*0.5)))/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100 "
            			elseif itemcoupontype="2" then   	''�ݾ� 
            				sqlStr = sqlStr & "  (((i.sellcash-"&Cstr(itemcouponvalue)&")- ( i.buycash - convert(int," & CStr(itemcouponvalue) & "*0.5)))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100 "
            			else 
            				sqlStr = sqlStr & " 0 " & VbCrlf
            			end if
            		Case "60"	''��ü�δ� - ���԰� ����
            			if itemcoupontype="1" then			''������ 
            				sqlStr = sqlStr  & " (((i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)) - (i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & ")))/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100 "
            			elseif itemcoupontype="2" then   	''�ݾ� 
            				sqlStr = sqlStr  & " (((i.sellcash-"&Cstr(itemcouponvalue)&") - (i.buycash - " & CStr(itemcouponvalue) &"))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100 "
            			else 
            				sqlStr = sqlStr & " 0 " & VbCrlf
            			end if
                    Case "80"   ''���������� -500 
                            sqlStr = sqlStr  & " case when i.mwdiv='M' then 0 else ((i.sellcash- (i.buycash - 500))/i.sellcash)*100 end "
            		Case "90"	''20%��ü��� - �����ΰ�� �����԰�.
            			if itemcoupontype="1" then			''������ 
            				sqlStr = sqlStr & " case when i.mwdiv='M' then 0 else (((i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash))-(i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & "*0.5)))/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100 end "
            			elseif itemcoupontype="2" then   	''�ݾ� 
            				sqlStr = sqlStr & " case when i.mwdiv='M' 0 else (((i.sellcash-"&Cstr(itemcouponvalue)&")-(i.buycash - convert(int," & CStr(itemcouponvalue) & "*0.5)))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100  end "
            			else 
            				sqlStr = sqlStr & " 0 " & VbCrlf
            			end if
            
            		Case else 
            			sqlStr = sqlStr & " 0 " & VbCrlf
				End SELECT
				sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_item_coupon_detail d " & VbCrlf
				sqlStr = sqlStr & " JOIN [db_item].[dbo].tbl_item i "
				sqlStr = sqlStr & " 	on d.itemid = i.itemid "
				sqlStr = sqlStr & " WHERE d.itemcouponidx=" & CStr(itemcouponidx)
          
        '  response.write sqlStr
       ' response.end
				dbget.Execute sqlStr
				
				''2������ 2018/06/19 
        	    sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
        	    dblogicsget.Execute sqlStr
		End IF

	else
		''�ű� ���
		sqlStr = "select * from [db_item].[dbo].tbl_item_coupon_master where 1=0"
		rsget.Open sqlStr,dbget,1,3
		rsget.AddNew

		rsget("itemcoupontype") = itemcoupontype
		rsget("couponGubun") = couponGubun
		rsget("itemcouponvalue") = itemcouponvalue
		rsget("itemcouponstartdate") = itemcouponstartdate
		rsget("itemcouponexpiredate") = itemcouponexpiredate
		rsget("itemcouponname") = itemcouponname
		rsget("itemcouponexplain") = itemcouponexplain

		rsget("openstate") = "0"
		rsget("margintype") = margintype
		rsget("defaultmargin")	= defaultmargin
		rsget("reguserid") = session("ssBctId")
		rsget("coupontype") = coupontype

		rsget.update
			itemcouponidx = rsget("itemcouponidx")
		rsget.close

	end if
elseif mode="I" then
    '' �߰� �˾�â���� �Ѿ� �� ���.
	ErrStr = ""

	''����Ÿ�� ��������
	margintype = "00"

	sqlStr = "select top 1 margintype, itemcoupontype, itemcouponvalue,couponGubun,"
	sqlStr = sqlStr + " convert(varchar(19),itemcouponstartdate,21) as itemcouponstartdate,"
	sqlStr = sqlStr + " convert(varchar(19),itemcouponexpiredate,21) as itemcouponexpiredate"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_master" + VbCrlf
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx)
	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		margintype = rsget("margintype")
		itemcoupontype = rsget("itemcoupontype")
		itemcouponvalue = rsget("itemcouponvalue")
		itemcouponstartdate = rsget("itemcouponstartdate")
		itemcouponexpiredate = rsget("itemcouponexpiredate")
		couponGubun = rsget("couponGubun")  ''�Ϲ�/���̹�/�����ε�.
	end if
	rsget.close

	itemidarr = trim(itemidarr)
	if Right(itemidarr,1)="," then itemidarr=Left(itemidarr,Len(itemidarr)-1)

	'' ������ �����ϰ��, ��ü��ǰ �� �ٹ蹫���� ���رݾ� �ʰ� ��ǰ �ȳ�
	if itemcoupontype=3 then
		sqlStr = "Select top 100 itemid, mwdiv, sellcash " & vbCRLF
		sqlStr = sqlStr & " from db_item.dbo.tbl_item " & vbCRLF
		sqlStr = sqlStr & " Where itemid in (" & itemidarr & ")"
		
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if not rsget.Eof then
			do until rsget.Eof
				if rsget("mwdiv")="U" then ErrStr = ErrStr + "-��ü��� ��ǰ (��ǰ��ȣ : " + CStr(rsget("itemid")) + ") ��ϺҰ� \n"
				if rsget("mwdiv")<>"U" and rsget("sellcash")>=30000 then ErrStr = ErrStr + "- ������ ��ǰ (��ǰ��ȣ : " + CStr(rsget("itemid")) + ") ��ϺҰ� \n"
				rsget.moveNext
			loop

			if ErrStr<>"" then
				response.write "<script language=javascript>alert('��۷����� ��������\n\n" + ErrStr + "');</script>"
				response.End
			end if
		end if
		rsget.close
	end if
	
	Call debugRwite("step1")
    ''�˻��� ��ü ��ǰ�� ���.. �˻��� ��� ���� insert  ó��
    addSql = ""
    IF (sType="all") THEN

         '// �߰� ����
		
        if (makerid <> "") then
            addSql = addSql & " and i.makerid='" + makerid + "'"
        end if

        if (itemid <> "") then
			itemid = trim(itemid)
			itemid = replace(itemid,chr(13),"")
			itemid = replace(itemid,chr(10),",")
			if Right(itemid,1)="," then itemid=Left(itemid,Len(itemid)-1)

            addSql = addSql & " and i.itemid in (" + itemid + ")"
        end if

        ''if (itemname <> "") then
        ''    addSql = addSql & " and i.itemname like '%" + html2db(itemname) + "%'"
        ''end if

        if (sellyn="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif( sellyn="SR") then
        	  addSql = addSql & " and i.sellyn='N' and r.itemid is not null "
        elseif (sellyn <> "") then
            addSql = addSql & " and i.sellyn='" + sellyn + "'"
        end if
        
        if (usingyn <> "") then
            addSql = addSql & " and i.isusing='" + usingyn + "'"
        end if

        if danjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif danjongyn="YM" then
            addSql = addSql + " and i.danjongyn<>'N'"
            addSql = addSql + " and i.danjongyn<>'S'"
        elseif danjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + danjongyn + "'"
        end if
       
        if mwdiv="MW" then
            addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif mwdiv<>"" then
            addSql = addSql + " and i.mwdiv='" + mwdiv + "'"
        end if
        
		if limityn="Y0" then
            addSql = addSql + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif limityn<>"" then
            addSql = addSql + " and i.limityn='" + limityn + "'"
        end if

        if cdl<>"" then
            addSql = addSql + " and i.cate_large='" + cdl + "'"
        end if

        if cdm<>"" then
            addSql = addSql + " and i.cate_mid='" + cdm + "'"
        end if

        if cds<>"" then
            addSql = addSql + " and i.cate_small='" + cds + "'"
        end if

        if disp<>"" then
		    if LEN(disp)>3 then
		         addSql = addSql + " and i.dispcate1='"&LEFT(disp,3)&"'" ''2015/03/27�߰�
		    end if
			addSql = addSql + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item WITH(NOLOCK) where catecode like '" + disp + "%' and isDefault='y') "
		end if
		
		if couponyn<>"" then
            addSql = addSql + " and i.itemCouponyn='" + couponyn + "'"
        end if

        if sailyn<>"" then
            addSql = addSql + " and i.sailyn='" + sailyn + "'"
        end if

        if deliverytype <> "" then
        	addSql = addSql + " and i.deliverytype='" + deliverytype + "'"
        end if

        If minmargin <> "" Then
        	addSql = addSql + " and i.itemid <> 0 and i.isusing = 'Y' and i.itemdiv <> '82' and i.sellcash <> 0 and ((1-(i.buycash/i.sellcash))*100) >= " & minmargin & " "
        End If
        
        if (itemcostup<>"") then
            addSql = addSql & " and i.sellcash>="&itemcostup&""&vbCRLF
        end if
        
        if (itemcostdown<>"") then
            addSql = addSql & " and i.sellcash<="&itemcostdown&""&vbCRLF
        end if

		if (groupId<>"") then
            addSql = addSql & " and Exists ( "
            addSql = addSql & " 	select 1 "
            addSql = addSql & " 	from db_partner.dbo.tbl_partner as sp "
            addSql = addSql & " 		join db_user.dbo.tbl_user_c as sc "
            addSql = addSql & " 			on sp.id=sc.userid "
            addSql = addSql & " 	where sp.id=i.makerid "
            addSql = addSql & " 		and sp.isusing='Y' and sc.isusing='Y' "
            addSql = addSql & " 		and sc.userdiv='02' "
            addSql = addSql & " 		and sp.groupid='" & groupId & "' "
            addSql = addSql & " ) "
		end if

        ''EP ��������
        addSql = addSql & " and i.itemdiv<>'21'"
        addSql = addSql & " and i.makerid not in (select makerid from db_temp.dbo.tbl_EpShop_not_in_makerid WITH(NOLOCK) where mallgubun='naverep' and isusing='N')"
        addSql = addSql & " and i.itemid not in (Select itemid From db_temp.dbo.tbl_EpShop_not_in_itemid WITH(NOLOCK) Where mallgubun='naverep' AND isusing = 'Y')"
        ''addSql = addSql & " and i.itemid not in (select itemid from db_temp.dbo.tbl_EpShop_Mapping_item)"
        ''addSql = addSql & " and Not Exists(select 1 from db_temp.dbo.tbl_naver_item_map nn where nn.serviceyn='y' and nn.tenitemid=i.itemid)"   ''tbl_nvshop_mapItem ���κ��� 
        ''addSql = addSql & " and i.itemid not in (select itemid from db_temp.dbo.tbl_EpShop_RecentSell_item where (sellNDays>=6 or sell1Days>=2))"  ''�ֱ� �Ǹų��� N���̻� ���� (�ּ�ó�� 2018/07/19)
        
        ''2018/07/18
        addSql = addSql & " and i.makerid not in ( select makerid from db_temp.dbo.tbl_Epshop_itemcoupon_Except_Brand WITH(NOLOCK) where isNULL(AsignMaxDt,'2099-01-01')>getdate() )"
        addSql = addSql & " and i.itemid not in ( select itemid from db_temp.dbo.tbl_Epshop_itemcoupon_Except_item WITH(NOLOCK) where isNULL(AsignMaxDt,'2099-01-01')>getdate() )"
        
		''2019/11/04 �����߰�.
        addSql = addSql & " and Not Exists(select 1 from [db_temp].dbo.[tbl_Epshop_fixedPrice] fx WITH(NOLOCK) where fx.itemid=i.itemid)"
		''���� �б�ó��. ����� ��û 2019/11/04
        if (exceptnotepmapitem="") then
			addSql = addSql & " and Not Exists(select 1 from [db_etcmall].dbo.[tbl_nvshop_mapItem] nn WITH(NOLOCK) where nn.itemid=i.itemid)" 
		end if

        ''��Ͽ�����������
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & "     select itemid"
        addSql = addSql & "     from [db_item].[dbo].tbl_item_coupon_master m WITH(NOLOCK) "
        addSql = addSql & "         Join [db_item].[dbo].tbl_item_coupon_detail d WITH(NOLOCK) "
        addSql = addSql & "         on m.itemcouponidx=d.itemcouponidx"
        addSql = addSql & "         and m.openstate<9"
        addSql = addSql & "     where m.itemcouponexpiredate>getdate()"
		addSql = addSql + " 	and NOT ("
		addSql = addSql + " 		(m.itemcouponstartdate>'" + CStr(itemcouponexpiredate) + "')"
		addSql = addSql + " 		or "
		addSql = addSql + " 		(m.itemcouponexpiredate<'" + CStr(itemcouponstartdate) + "')"
        ' addSql = addSql & "     (m.itemcouponstartdate<='"&itemcouponstartdate&"' and m.itemcouponexpiredate>'"&itemcouponstartdate&"')"
        ' addSql = addSql & "     or"
        ' addSql = addSql & "     (m.itemcouponstartdate<='"&itemcouponexpiredate&"' and m.itemcouponexpiredate>'"&itemcouponexpiredate&"')"
        addSql = addSql & "     )"
		'if (couponGubun="V") then  ''�ߺ����� üũ��� ���� //2019/03/25 => �˻��� ��� �Է��ϴ� ���̽��� �̹���� Ÿ�� �ʴ´�.
		'	addSql = addSql + " and m.couponGubun='V'"
		'else
		'	addSql = addSql + " and m.couponGubun<>'V'"
		'end if
        addSql = addSql & " )"

        if (addSql="") then
            addSql = "select i.itemid from [db_item].[dbo].tbl_item i WITH(NOLOCK) where 1=0 "
        else
            addSql = "select i.itemid from [db_item].[dbo].tbl_item i WITH(NOLOCK) where i.itemid<>0 " & addSql
        end if
         
        '' counting ����
        dim iCountQuery, paraitemcount, isubcnt : isubcnt=0
        iCountQuery = replace(addSql,"select i.itemid","select count(*) cnt")

        rsget.CursorLocation = adUseClient
        rsget.Open iCountQuery, dbget, adOpenForwardOnly, adLockReadOnly
	    if not rsget.Eof then
	        isubcnt = rsget("cnt")
	    end if
	    rsget.close
	    
	    paraitemcount = request("itemcount")
	    
	    if (CStr(paraitemcount)<>CStr(isubcnt)) then
	        
	        response.write "<script>alert('���� ���� :"&paraitemcount&":"&isubcnt&"');</script>"
	        rw addSql
	        dbget.Close() : response.end
	    end if
    ELSE
    	addSql = trim(itemidarr)
	END IF
	Call debugRwite("step2")

	'' �ٸ� ������ ��ǰ�� ��ϵǾ� ������� üũ
	sqlStr = " select top 100 m.itemcouponidx, d.itemid from"
	sqlStr = sqlStr + " [db_item].[dbo].tbl_item_coupon_master m WITH(NOLOCK) "
	sqlStr = sqlStr + " Join [db_item].[dbo].tbl_item_coupon_detail d WITH(NOLOCK) "
	sqlStr = sqlStr + " on m.itemcouponidx=d.itemcouponidx"
	sqlStr = sqlStr + " where m.itemcouponidx<>" + CStr(itemcouponidx)
	sqlStr = sqlStr + " and m.openstate<9"			''�߱������ΰ� ����
	'sqlStr = sqlStr + " and m.couponGubun<>'P'"		''�����ι߱������� ���� (20140617; ������) , �ߺ����� �Ұ���(2018.01.22)
	' sqlStr = sqlStr + " and ( "
	' sqlStr = sqlStr + " 	(m.itemcouponstartdate<='" + CStr(itemcouponstartdate) + "' and m.itemcouponexpiredate>'" + CStr(itemcouponstartdate) + "')"
	' sqlStr = sqlStr + " 	or "
	' sqlStr = sqlStr + " 	(m.itemcouponstartdate<='" + CStr(itemcouponexpiredate) + "' and m.itemcouponexpiredate>'" + CStr(itemcouponexpiredate) + "')"
	' sqlStr = sqlStr + " 	)"
	sqlStr = sqlStr + " and m.itemcouponexpiredate>getdate()"
	sqlStr = sqlStr + " and NOT ( "
	sqlStr = sqlStr + " 	(m.itemcouponstartdate>'" + CStr(itemcouponexpiredate) + "')"
	sqlStr = sqlStr + " 	or "
	sqlStr = sqlStr + " 	(m.itemcouponexpiredate<'" + CStr(itemcouponstartdate) + "')"
	sqlStr = sqlStr + " 	)"

	if (sType<>"all") then  ''�˻��� ��� �Է��ϴ� ���̽��� �̹���� Ÿ�� �ʴ´�. //2019/03/29
		if (couponGubun="V") then  ''�ߺ����� üũ��� ���� //2019/03/25
			sqlStr = sqlStr + " and m.couponGubun='V'"
		else
			sqlStr = sqlStr + " and m.couponGubun not in ('V','P','T')"  ''P,T�߰�(secret==P �������.) 2019/06/11
		end if
	end if
	sqlStr = sqlStr + " and d.itemid in (" + addSql + ")"  + VbCrlf

	

	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.Eof then
		do until rsget.Eof
			ErrStr = ErrStr + "������ȣ : " + CStr(rsget("itemcouponidx")) + " - ��ǰ��ȣ : " + CStr(rsget("itemid")) + " ����� \n"
			rsget.moveNext
		loop
	end if
	rsget.close

	Call debugRwite("step3")
	'' ���� ��Ͽ� ���� ��ǰ�� �߰�.
	sqlStr = "insert into [db_item].[dbo].tbl_item_coupon_detail " & VbCrlf
	sqlStr = sqlStr & " (itemcouponidx, itemid, couponbuyprice, couponmargin)" & VbCrlf
	sqlStr = sqlStr & " select "& CStr(itemcouponidx) & "," & VbCrlf
	sqlStr = sqlStr & " i.itemid, " & VbCrlf
	Select Case margintype
		Case "00"  	''��ǰ�������� - ���԰� 0 �ΰ�� �����԰�
			sqlStr = sqlStr & " 0 " & VbCrlf
		'Case "10"	''�ٹ����ٺδ� - ���԰� ����x
		'	if itemcoupontype="1" then			''������
		'		sqlStr = sqlStr + " i.buycash - convert(int,i.sellcash*" + CStr(itemcouponvalue/100) + ")"
		'	elseif itemcoupontype="2" then   	''�ݾ�
		'		sqlStr = sqlStr + " i.buycash - " + CStr(itemcouponvalue)
		'	else
		'		sqlStr = sqlStr + " 0 " + VbCrlf
		'	end if
		    sqlStr = sqlStr & ", 0" & VbCrlf
		Case "10"	''�ٹ����ٺδ� - �����԰�
			sqlStr = sqlStr & " 0 " & VbCrlf
            sqlStr = sqlStr & ", 0" & VbCrlf
		Case "20"	''�������� : �߰� [2008-09-23]
		 
			if itemcoupontype="1" then			''������
				sqlStr = sqlStr & " convert(int,i.sellcash*"& Cstr((100-itemcouponvalue)/100) &"*"& Cstr((100-defaultmargin)/100) &")"
                sqlStr = sqlStr & ", ( ( (i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)) - (convert(int,i.sellcash*"& Cstr((100-itemcouponvalue)/100) &"*"& Cstr((100-defaultmargin)/100) &")) )/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100"
			elseif itemcoupontype="2" then   	''�ݾ�
				sqlStr = sqlStr & " convert(int,(i.sellcash-" & CStr(itemcouponvalue) + ")*"& Cstr((100-defaultmargin)/100) &")"
				sqlStr = sqlStr & " ,(((i.sellcash-"&Cstr(itemcouponvalue)&") -(convert(int,(i.sellcash-" & CStr(itemcouponvalue) & ")*"& Cstr((100-defaultmargin)/100) &")))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100"
			else
				sqlStr = sqlStr & " 0 " & VbCrlf
				sqlStr = sqlStr & " ,0 " & VbCrlf
			end if
		Case "30"	''���ϸ��� - ���縶�� : �߰� [2008-09-23]
			if itemcoupontype="1" then			''������
				sqlStr = sqlStr & " convert(int,i.sellcash*" & CStr((100-itemcouponvalue)/100) & "*i.buycash/i.sellcash)"
				sqlStr = sqlStr & " , (((i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)) - (convert(int,i.sellcash*" & CStr((100-itemcouponvalue)/100) & "*i.buycash/i.sellcash)))/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100 "
			elseif itemcoupontype="2" then   	''�ݾ�
				sqlStr = sqlStr & " convert(int,(i.sellcash-" & CStr(itemcouponvalue) & ")*i.buycash/i.sellcash)"
				sqlStr = sqlStr & " , (((i.sellcash-"&Cstr(itemcouponvalue)&") - (convert(int,(i.sellcash-" & CStr(itemcouponvalue) & ")*i.buycash/i.sellcash)))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100 "
			else
				sqlStr = sqlStr & " 0 " & VbCrlf
				sqlStr = sqlStr & " ,0 " & VbCrlf
			end if
		Case "50"	''�ݹݺδ�
			if itemcoupontype="1" then			''������
				sqlStr = sqlStr  & " i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & "*0.5)"
				sqlStr = sqlStr  & " , (((i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)) - ( i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & "*0.5)))/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100"
			elseif itemcoupontype="2" then   	''�ݾ�
				sqlStr = sqlStr & " i.buycash - convert(int," & CStr(itemcouponvalue) & "*0.5)"
				sqlStr = sqlStr & " , (((i.sellcash-"&Cstr(itemcouponvalue)&")- ( i.buycash - convert(int," & CStr(itemcouponvalue) & "*0.5)))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100 "
			else
				sqlStr = sqlStr  & " 0 "  & VbCrlf
				sqlStr = sqlStr & " ,0 " & VbCrlf
			end if
		Case "60"	''��ü�δ� - ���԰� ����
			if itemcoupontype="1" then			''������
				sqlStr = sqlStr  & " i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & ")"
				sqlStr = sqlStr  & " , (((i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)) - (i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & ")))/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100"
			elseif itemcoupontype="2" then   	''�ݾ�
				sqlStr = sqlStr  & " i.buycash - " & CStr(itemcouponvalue)
				sqlStr = sqlStr  & "  , (((i.sellcash-"&Cstr(itemcouponvalue)&") - (i.buycash - " & CStr(itemcouponvalue) &"))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100"
			else
				sqlStr = sqlStr  & " 0 "  & VbCrlf
				sqlStr = sqlStr & " ,0 " & VbCrlf
			end if
        Case "80"   ''���������� -500
                sqlStr = sqlStr  & " case when i.mwdiv='M' then 0 else i.buycash - 500 end "
                sqlStr = sqlStr  & ", case when i.mwdiv='M' then 0 else ((i.sellcash- (i.buycash - 500))/i.sellcash)*100 end "
		Case "90"	''20%��ü��� - �����ΰ�� �����԰�.
			if itemcoupontype="1" then			''������
				sqlStr = sqlStr & " case when i.mwdiv='M' then 0 else i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & "*0.5) end "
				sqlStr = sqlStr & ", case when i.mwdiv='M' then 0 else (((i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash))-(i.buycash - convert(int,i.sellcash*" & CStr(itemcouponvalue/100) & "*0.5)))/(i.sellcash-("&Cstr(itemcouponvalue/100)&"*i.sellcash)))*100 end "
			elseif itemcoupontype="2" then   	''�ݾ�
				sqlStr = sqlStr & " case when i.mwdiv='M' 0 else i.buycash - convert(int," & CStr(itemcouponvalue) & "*0.5)  end "
				sqlStr = sqlStr & ", case when i.mwdiv='M' 0 else (((i.sellcash-"&Cstr(itemcouponvalue)&")-(i.buycash - convert(int," & CStr(itemcouponvalue) & "*0.5)))/(i.sellcash-"&Cstr(itemcouponvalue)&"))*100  end "
			else
				sqlStr = sqlStr  & " 0 "  & VbCrlf
				sqlStr = sqlStr & " ,0 " & VbCrlf
			end if

		Case else
			sqlStr = sqlStr  & " 0 "  & VbCrlf
			sqlStr = sqlStr & " ,0 " & VbCrlf
	end Select

	sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i WITH(NOLOCK) " &VbCrlf
	sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item_coupon_detail d WITH(NOLOCK) " & VbCrlf
	sqlStr = sqlStr & " 	on d.itemcouponidx=" & CStr(itemcouponidx) & VbCrlf
	sqlStr = sqlStr & " 	and d.itemid=i.itemid" & VbCrlf
	sqlStr = sqlStr & " where i.itemid in (" & addSql & ")"  & VbCrlf
	sqlStr = sqlStr & " 	and d.itemid is null"
	sqlStr = sqlStr + "		and i.itemdiv<>'21' "  ''����ǰ ����
	sqlStr = sqlStr & " 	and i.itemid not in ("
	sqlStr = sqlStr & " 		select distinct d.itemid from"
	sqlStr = sqlStr & " 			[db_item].[dbo].tbl_item_coupon_master m WITH(NOLOCK) ,"
	sqlStr = sqlStr & " 			[db_item].[dbo].tbl_item_coupon_detail d WITH(NOLOCK) "
	sqlStr = sqlStr & " 		where m.itemcouponidx=d.itemcouponidx"
	sqlStr = sqlStr & " 			and m.itemcouponidx<>" & CStr(itemcouponidx)
	sqlStr = sqlStr & " 			and m.openstate<9"  ''�߱������ΰ� ����
	sqlStr = sqlStr + " 			and m.itemcouponexpiredate>getdate()"
	sqlStr = sqlStr + " 			and NOT ( "
	sqlStr = sqlStr + " 				((m.itemcouponstartdate>'" + CStr(itemcouponexpiredate) + "')"
	sqlStr = sqlStr + " 				or "
	sqlStr = sqlStr + " 				(m.itemcouponexpiredate<'" + CStr(itemcouponstartdate) + "'))"
	sqlStr = sqlStr + " 			)"
	if (sType<>"all") then  ''�˻��� ��� �Է��ϴ� ���̽��� �̹���� Ÿ�� �ʴ´�. //2019/03/29
		if (couponGubun="V") then  ''�ߺ����� üũ��� ���� //2019/03/25
			sqlStr = sqlStr + " and m.couponGubun='V'"
		else
			sqlStr = sqlStr + " and m.couponGubun not in ('V','P','T') "
		end if
	end if
	sqlStr = sqlStr & " 	and d.itemid in (" & addSql & ")"  & VbCrlf
	sqlStr = sqlStr & " ) "
 
	dbget.CommandTimeout = 150	'5��
	dbget.Execute sqlStr
	Call debugRwite("step4")
	
	Call AplyToItem(itemcouponidx,false)
	Call debugRwite("step5")
	''�����ǰ��.
	AplyItemCountUpdate itemcouponidx
	
	Call debugRwite("step6")
	if Not(itemid="" and itemidarr="") then
		Call AddSCMChangeLog(itemcouponidx, "- ��ǰ����>��ǰ�߰� : " & itemid & itemidarr)
	end if
elseif mode="delcouponitemarr" then
	itemidarr = trim(itemidarr)
	if Right(itemidarr,1)="," then itemidarr=Left(itemidarr,Len(itemidarr)-1)

	sqlStr = "delete from [db_item].[dbo].tbl_item_coupon_detail" + VbCrlf
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
	sqlStr = sqlStr + " and itemid in (" + itemidarr + ")"  + VbCrlf

	dbget.Execute sqlStr

    ''2������ 2018/06/19 
    sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
    dblogicsget.Execute sqlStr
    ''�����ΰ�� 2�������� lastupdate�� ������.. Naver EP����.
    sqlStr = "update [db_AppWish].[dbo].tbl_item  "+ VbCrlf
    sqlStr = sqlStr + " set lastupdate =getdate()"+ VbCrlf
    sqlStr = sqlStr + " where itemid in (" + itemidarr + ")"  + VbCrlf
    dblogicsget.Execute sqlStr
    
	''������ ���� ��ǰ���̺��� ���� ���� N �� ����
	Call AplyToItem(itemcouponidx,false)

	''�����ǰ��.
	AplyItemCountUpdate itemcouponidx

	if itemidarr<>"" then
		Call AddSCMChangeLog(itemcouponidx, "- ��ǰ����>��ǰ���� : " & itemidarr)
	end if
elseif mode="delBrandAll" then
	'// �귣�� ��ǰ �ϰ� ����
	if makerid<>"" then
		sqlStr = "delete from cd " + VbCrlf
		sqlStr = sqlStr & " from [db_item].[dbo].tbl_item_coupon_detail as cd with(noLock) " + VbCrlf
		sqlStr = sqlStr & " 	join [db_item].[dbo].tbl_item as i with(noLock) " + VbCrlf
		sqlStr = sqlStr & " 		on cd.itemid=i.itemid " + VbCrlf
		sqlStr = sqlStr & " where cd.itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
		sqlStr = sqlStr & " 	and i.makerid='" + makerid + "' " + VbCrlf
		dbget.Execute sqlStr

		''�����ΰ�� 2������������ ���� lastupdate�� ������.. Naver EP����.
		sqlStr = "update i "+ VbCrlf
		sqlStr = sqlStr + " set i.lastupdate =getdate()"+ VbCrlf
		sqlStr = sqlStr & " from [db_AppWish].[dbo].tbl_item as i with(noLock) " + VbCrlf
		sqlStr = sqlStr & "		join [db_AppWish].[dbo].tbl_item_coupon_detail as cd with(noLock) " + VbCrlf
		sqlStr = sqlStr & " 		on i.itemid=cd.itemid " + VbCrlf
		sqlStr = sqlStr & " where cd.itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
		sqlStr = sqlStr & " 	and i.makerid='" + makerid + "' " + VbCrlf
		dblogicsget.Execute sqlStr

		''2������ 2018/06/19 
		sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
		dblogicsget.Execute sqlStr


		''������ ���� ��ǰ���̺��� ���� ���� N �� ����
		Call AplyToItem(itemcouponidx,false)

		''�����ǰ��.
		AplyItemCountUpdate itemcouponidx

		Call AddSCMChangeLog(itemcouponidx, "- ��ǰ����>�귣���ǰ���� : " & makerid)
	end if

elseif mode="delcouponitemmulti" then  ''2018/05/17
    dim midxArrQue
    itemcouponidxarr = split(itemcouponidxarr,",")
    itemidarr        = split(itemidarr,",")
    
    if (Lbound(itemcouponidxarr)<>Lbound(itemidarr)) or (Ubound(itemcouponidxarr)<>Ubound(itemidarr)) then
        response.Write "<script>alert('param ����')</script>"
		response.Write "<script>history.back()</script>"
		dbget.close()	:	response.End
    end if
    
    for i=Lbound(itemcouponidxarr) to Ubound(itemcouponidxarr)
        if (itemcouponidxarr(i)<>"") and (itemidarr(i)<>"") then
            ''rw itemcouponidxarr(i)&","&itemidarr(i)
            
            sqlStr = "delete from [db_item].[dbo].tbl_item_coupon_detail" + VbCrlf
        	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidxarr(i)) + VbCrlf
        	sqlStr = sqlStr + " and itemid in (" + itemidarr(i) + ")"  + VbCrlf
        
        	dbget.Execute sqlStr
        	
        	''2������ 2018/06/19 
        	sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
            dblogicsget.Execute sqlStr
            ''�����ΰ�� 2�������� lastupdate�� ������.. Naver EP����.
            sqlStr = "update [db_AppWish].[dbo].tbl_item  "+ VbCrlf
            sqlStr = sqlStr + " set lastupdate =getdate()"+ VbCrlf
            sqlStr = sqlStr + " where itemid in (" + itemidarr(i) + ")"  + VbCrlf
            dblogicsget.Execute sqlStr
    
        	if Not(InStr(midxArrQue,itemcouponidxarr(i)&",")>0) then
        	    midxArrQue = midxArrQue&itemcouponidxarr(i)&","
        	end if

			Call AddSCMChangeLog(itemcouponidxarr(i), "- ��ǰ����>��ǰ���� : " & itemidarr(i))
        end if
    next
    
    midxArrQue = split(midxArrQue,",")
    for i=Lbound(midxArrQue) to Ubound(midxArrQue)
        if (midxArrQue(i)<>"") then
            ''������ ���� ��ǰ���̺��� ���� ���� N �� ����
        	Call AplyToItem(midxArrQue(i),false)
        
        	''�����ǰ��.
        	AplyItemCountUpdate midxArrQue(i)


        end if
    next
    
elseif mode="modicouponitemarr" then
	itemidarr = trim(itemidarr)
	couponbuypricearr  = trim(couponbuypricearr)
    couponsellcasharr = trim(couponsellcasharr)
    
	if Right(itemidarr,1)="," then itemidarr=Left(itemidarr,Len(itemidarr)-1)
	if Right(couponbuypricearr,1)="," then couponbuypricearr=Left(couponbuypricearr,Len(couponbuypricearr)-1)
	if Right(couponsellcasharr,1)="," then couponsellcasharr=Left(couponsellcasharr,Len(couponsellcasharr)-1)

	itemidarr = split(itemidarr,",")
	couponbuypricearr = split(couponbuypricearr,",")
    couponsellcasharr = split(couponsellcasharr,",")
    
	for i=LBound(itemidarr) to UBound(itemidarr)
		if trim(itemidarr(i))<>"" then
			sqlStr = "update D" + VbCrlf
			sqlStr = sqlStr + " set couponbuyprice=" + CStr(couponbuypricearr(i)) + VbCrlf
			if (TRIM(couponbuypricearr(i))="0") or (TRIM(couponsellcasharr(i))="0") or (TRIM(couponsellcasharr(i))="") then
			    sqlStr = sqlStr + " ,couponmargin=0" + VbCrlf
			else
			    sqlStr = sqlStr + " ,couponmargin=(1-" +CStr(couponbuypricearr(i))+"*1.0/"+CStr(couponsellcasharr(i))+")*100"
			end if
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_coupon_detail D" + VbCrlf
			sqlStr = sqlStr + " where D.itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
			sqlStr = sqlStr + " and D.itemid=" + CStr(itemidarr(i)) + VbCrlf

			dbget.Execute sqlStr
			
			''2������ 2018/06/19 
        	sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
            dblogicsget.Execute sqlStr
            ''2������ lastupdate for Naver coupon
            sqlStr = "update [db_AppWish].[dbo].tbl_item  "+ VbCrlf
            sqlStr = sqlStr + " set lastupdate =getdate()"+ VbCrlf
            sqlStr = sqlStr + " where itemid in (" + itemidarr(i) + ")"  + VbCrlf
            dblogicsget.Execute sqlStr
		end if
	next

	Call AplyToItem(itemcouponidx,false)

	''�����ǰ��.
	AplyItemCountUpdate itemcouponidx
elseif mode="opencoupon" Then

	sqlStr = "update [db_item].[dbo].tbl_item_coupon_master" + VbCrlf
	sqlStr = sqlStr + " set openstate='7'"
	sqlStr = sqlStr + ",lastupDt=getdate()" + VbCrlf
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf

	dbget.Execute sqlStr
'response.write sqlStr
	Call AplyToItem(itemcouponidx,true)
	
	''2������ 2018/06/19 
    sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
    dblogicsget.Execute sqlStr

	Call AddSCMChangeLog(itemcouponidx, "- ��ǰ����>��������")

elseif mode="reservecoupon" Then

	sqlStr = "update [db_item].[dbo].tbl_item_coupon_master" + VbCrlf
	sqlStr = sqlStr + " set openstate='6'"
	sqlStr = sqlStr + ",lastupDt=getdate()" + VbCrlf
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf

	dbget.Execute sqlStr

    ''2������ 2018/06/19 
    sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
    dblogicsget.Execute sqlStr

	Call AddSCMChangeLog(itemcouponidx, "- ��ǰ����>��������")

elseif mode="closecoupon" Then

    dim MayExpireDt
    MayExpireDt = Left(CStr(DateAdd("d",-1,Now())),10) & " 23:59:59"

    ''response.write MayExpireDt

    ''�� �߱� �� ���� Expire
    sqlStr = "update [db_item].[dbo].tbl_user_item_coupon" + VbCrlf
    sqlStr = sqlStr + " set itemcouponexpiredate='" & MayExpireDt & "'" + VbCrlf
    sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
    sqlStr = sqlStr + " and itemcouponexpiredate>'" & MayExpireDt & "'" + VbCrlf
    sqlStr = sqlStr + " and usedyn='N'" + VbCrlf

    dbget.Execute sqlStr
    

	sqlStr = "update [db_item].[dbo].tbl_item_coupon_master" + VbCrlf
	sqlStr = sqlStr + " set openstate='9'"
	sqlStr = sqlStr + ",lastupDt=getdate()" + VbCrlf
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf

	dbget.Execute sqlStr
	
	''2������ 2018/06/19 
    sqlStr = replace(sqlStr,"[db_item].[dbo].","[db_AppWish].[dbo].")
    dblogicsget.Execute sqlStr

	Call AplyToItem(itemcouponidx,true)

	Call AddSCMChangeLog(itemcouponidx, "- ��ǰ����>��������")

elseif mode="imageupload" Then
		''����

		sqlstr = "update [db_item].[dbo].tbl_item_coupon_master" + VbCrlf
		sqlstr = sqlstr + " set itemcouponimage='" + itemcouponimage + "'" + VbCrlf
		sqlstr = sqlstr + " where itemcouponidx=" + CStr(itemcouponidx)

		dbget.Execute sqlStr
end if

%>
<% if (mode="couponmaster") then %>
	<% if (IsEditMode) then %>
	<script language='javascript'>
	alert('���� �Ǿ����ϴ�.');
	location.replace('/admin/shopmaster/itemcouponmasterreg.asp?itemcouponidx=<%= itemcouponidx %>');
	</script>
	<% else %>
	<script language='javascript'>
	alert('���� �Ǿ����ϴ�. ��ǰ�� ��� �� �ּ���');
	opener.location.reload();
	window.close();
	//location.replace('/admin/shopmaster/itemcouponmasterreg.asp?itemcouponidx=<%= itemcouponidx %>');
	</script>
	<% end if %>
<% elseif mode="I" then %>
	<script language='javascript'>
	<%
	if ErrStr<>"" then
		ErrStr = ErrStr + "\n\n ������ �ߺ����� ���� �� �� �����ϴ�."
		response.write "alert('" + ErrStr + "')"
	end if
	%>
	alert('��ǰ�� ��� �Ǿ����ϴ�.');
	//location.replace('/admin/shopmaster/itemcouponitemlisteidt.asp?itemcouponidx=<%= itemcouponidx %>&makerid=<%= makerid %>&sailyn=<%= sailyn %>');
	</script>
<% elseif mode="delcouponitemarr" or mode="delBrandAll" then %>
	<script language='javascript'>
	alert('���� �Ǿ����ϴ�.');
	opener.location.reload();
	location.replace('/admin/shopmaster/itemcouponitemlisteidt.asp?itemcouponidx=<%= itemcouponidx %>&makerid=<%= makerid %>&sailyn=<%= sailyn %>');
	</script>
<% elseif mode="delcouponitemmulti" then %>
	<script language='javascript'>
	alert('���� �Ǿ����ϴ�.');
	opener.location.reload();
	location.replace('/admin/shopmaster/itemcouponitemlisteidtMulti.asp?makerid=<%= makerid %>');
	</script>
<% elseif mode="modicouponitemarr" then %>
	<script language='javascript'>
	alert('���� �Ǿ����ϴ�.');
	opener.location.reload();
	location.replace('/admin/shopmaster/itemcouponitemlisteidt.asp?itemcouponidx=<%= itemcouponidx %>&makerid=<%= makerid %>&sailyn=<%= sailyn %>');
	</script>
<% elseif mode="opencoupon" then %>
	<script language='javascript'>
	alert('������ ���� �Ǿ����ϴ�.');
	opener.location.reload();
	location.replace('/admin/shopmaster/itemcouponmasterreg.asp?itemcouponidx=<%= itemcouponidx %>');
	</script>
<% elseif mode="reservecoupon" then %>
	<script language='javascript'>
	alert('������ ������ ���� �Ǿ����ϴ�. ���� 0�ÿ� ����˴ϴ�.');
	opener.location.reload();
	location.replace('/admin/shopmaster/itemcouponmasterreg.asp?itemcouponidx=<%= itemcouponidx %>');
	</script>
<% elseif mode="closecoupon" then %>
	<script language='javascript'>
	alert('������ ���� �Ǿ����ϴ�.');
	opener.location.reload();
	location.replace('/admin/shopmaster/itemcouponmasterreg.asp?itemcouponidx=<%= itemcouponidx %>');
	</script>
<% elseif mode="imageupload" then %>
	<script language='javascript'>
	alert('���� �Ǿ����ϴ�.');
	location.replace('/admin/shopmaster/itemcouponmasterreg.asp?itemcouponidx=<%= itemcouponidx %>');
	</script>
<% end if %>
<%= "mode=" + mode %>
<!-- #include virtual="/lib/db/dblogicsclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->