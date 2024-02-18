<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ��ǰ ����
' History : 2010.09.30 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
function AplyItemCountUpdate(itemcouponidx)
	dim sqlStr
	
	''�����ǰ���� ������Ʈ
	sqlStr = "update [db_academy].dbo.tbl_diy_item_coupon_master" + VbCrlf
	sqlStr = sqlStr + " set applyitemcount=IsNULL(T.cnt,0)" + VbCrlf
	sqlStr = sqlStr + " from (" + VbCrlf
	sqlStr = sqlStr + "     select count(*) as cnt from [db_academy].dbo.tbl_diy_item_coupon_detail where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
	sqlStr = sqlStr + " ) as T" + VbCrlf
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
	
	'response.write sqlStr & "<br>"
	dbacademyget.Execute sqlStr
end function

function AplyToItem(itemcouponidx)
	dim sqlStr
	dim ocouponGubun, oitemcoupontype, oitemcouponvalue, oitemcouponstartdate, oitemcouponexpiredate, openstate, currdatetime
	dim couponExpired
	dim resultCnt

	applyitemcount = 0
	couponExpired = false

	sqlStr = "select top 1 couponGubun, margintype, itemcoupontype, itemcouponvalue, openstate, applyitemcount,"
	sqlStr = sqlStr + " convert(varchar(19),itemcouponstartdate,21) as itemcouponstartdate,"
	sqlStr = sqlStr + " convert(varchar(19),itemcouponexpiredate,21) as itemcouponexpiredate,"
	sqlStr = sqlStr + " (case when (itemcouponstartdate>getdate()) or (itemcouponexpiredate<getdate()) then 'Y' else 'N' end ) as couponexpired, "
	sqlStr = sqlStr + " convert(varchar(19),getdate()) as currdatetime"
	sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_coupon_master" + VbCrlf
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx)

	rsacademyget.Open sqlStr,dbacademyget,1
	if Not rsacademyget.Eof then
	    ocouponGubun   = rsacademyget("couponGubun")
		itemcoupontype = rsacademyget("itemcoupontype")
		itemcouponvalue = rsacademyget("itemcouponvalue")
		itemcouponstartdate = rsacademyget("itemcouponstartdate")
		itemcouponexpiredate = rsacademyget("itemcouponexpiredate")
		openstate = rsacademyget("openstate")
		applyitemcount = rsacademyget("applyitemcount")
		currdatetime = rsacademyget("currdatetime")

		couponExpired = rsacademyget("couponexpired")

		response.write "couponExpired :" + CStr(couponExpired) + "<br>"
	end if
	rsacademyget.Close

    ''Ÿ������, ���������ΰ�� ��ŵ.
    if (ocouponGubun<>"C") then exit function

	''�߱޴�����̰ų� �߱޿������ ��ŵ.
	if ((openstate<>"7") and (openstate<>"9")) then exit function

	''�߱� ����� �����ΰ�� -> N�� ����
	if (openstate="9") or (couponExpired="Y") then
		sqlStr = "update [db_academy].dbo.tbl_diy_item"
		sqlStr = sqlStr + " set itemcouponyn='N'"
		sqlStr = sqlStr + " ,itemcoupontype='1'"
		sqlStr = sqlStr + " ,itemcouponvalue=0"
		sqlStr = sqlStr + " ,curritemcouponidx=NULL"
		sqlStr = sqlStr + " ,lastupdate=getdate()"
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_coupon_detail"
		sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx)
		sqlStr = sqlStr + " and [db_academy].dbo.tbl_diy_item.itemid=[db_academy].dbo.tbl_diy_item_coupon_detail.itemid"

		'response.write sqlStr + "<br>"
		dbacademyget.Execute sqlStr
	end if

	''��ǰ�� �����Ȱ�� -> N�� ����
	sqlStr = "update [db_academy].dbo.tbl_diy_item"
	sqlStr = sqlStr + " set itemcouponyn='N'"
	sqlStr = sqlStr + " ,itemcoupontype='1'"
	sqlStr = sqlStr + " ,itemcouponvalue=0"
	sqlStr = sqlStr + " ,curritemcouponidx=NULL"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select i.itemid  "
	sqlStr = sqlStr + " 	from [db_academy].dbo.tbl_diy_item i"
	sqlStr = sqlStr + " 	left join [db_academy].dbo.tbl_diy_item_coupon_detail d"
	sqlStr = sqlStr + " 	on d.itemcouponidx=" + CStr(itemcouponidx) + " and i.itemid=d.itemid "
	sqlStr = sqlStr + " 	where i.curritemcouponidx=" + CStr(itemcouponidx)
	sqlStr = sqlStr + " 	and d.itemcouponidx is null"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_academy].dbo.tbl_diy_item.itemid=T.itemid"

	'response.write sqlStr + "<br>"
	dbacademyget.Execute sqlStr, resultCnt
	response.write "�����Ǽ�=" + CStr(resultCnt) + "<br>"

	''itemcouponidx�� ��ϵ� ��ǰ�� ��� �������� ������ Update
	sqlStr = "update [db_academy].dbo.tbl_diy_item"
	sqlStr = sqlStr + " set itemcouponyn='Y'"
	sqlStr = sqlStr + " ,itemcoupontype=T.itemcoupontype"
	sqlStr = sqlStr + " ,itemcouponvalue=T.itemcouponvalue"
	sqlStr = sqlStr + " ,curritemcouponidx=T.itemcouponidx"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select m.itemcouponidx, m.itemcoupontype, m.itemcouponvalue, d.itemid "
	sqlStr = sqlStr + " 	from [db_academy].dbo.tbl_diy_item_coupon_master m,"
	sqlStr = sqlStr + " 	[db_academy].dbo.tbl_diy_item_coupon_detail d"
	sqlStr = sqlStr + " 	where m.itemcouponidx=d.itemcouponidx"
	sqlStr = sqlStr + " 	and m.openstate='7'"
	sqlStr = sqlStr + " 	and d.itemcouponidx=" + CStr(itemcouponidx)
	sqlStr = sqlStr + " 	and m.itemcouponstartdate<=getdate()"
	sqlStr = sqlStr + " 	and m.itemcouponexpiredate>=getdate()"
	sqlStr = sqlStr + " ) T "
	sqlStr = sqlStr + " where [db_academy].dbo.tbl_diy_item.itemid=T.itemid"
	sqlStr = sqlStr + " and Not ("
	sqlStr = sqlStr + " 		 	[db_academy].dbo.tbl_diy_item.itemcouponyn='Y'"
	sqlStr = sqlStr + " 		and [db_academy].dbo.tbl_diy_item.itemcoupontype=T.itemcoupontype"
	sqlStr = sqlStr + " 		and [db_academy].dbo.tbl_diy_item.itemcouponvalue=T.itemcouponvalue"
	sqlStr = sqlStr + " 		and [db_academy].dbo.tbl_diy_item.curritemcouponidx=T.itemcouponidx"
	sqlStr = sqlStr + "			)"

	'response.write sqlStr + "<br>"
	dbacademyget.Execute sqlStr, resultCnt

    response.write "�����Ǽ�=" + CStr(resultCnt)
end function

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

dim itemcouponvalue ,itemcouponstartdate ,itemcoupontype ,couponGubun ,itemcouponidx
dim openstate ,margintype ,applyitemcount ,itemcouponexplain ,itemcouponimage ,itemcouponname ,itemcouponexpiredate
dim itemidarr, couponbuypricearr, makerid, sailyn ,ErrStr ,buf ,sqlstr,i ,IsEditMode ,mode ,defaultmargin
dim sType, addSql, itemid, itemname, sellyn, usingyn, danjongyn ,limityn, mwdiv, cdl, cdm, cds, deliverytype
	itemcouponidx      	= requestCheckVar(request("itemcouponidx"),9)
	couponGubun         = requestCheckVar(request("couponGubun"),9)
	itemcoupontype      = requestCheckVar(request("itemcoupontype"),9)
	itemcouponvalue     = requestCheckVar(request("itemcouponvalue"),9)
	itemcouponstartdate = RequestCheckvar(request("itemcouponstartdate"),10) + " " + RequestCheckvar(request("itemcouponstartdate2"),10)
	itemcouponexpiredate= RequestCheckvar(request("itemcouponexpiredate"),10) + " " + RequestCheckvar(request("itemcouponexpiredate2"),10)
	itemcouponname      = html2Db(request("itemcouponname"))
	itemcouponimage     = request("itemcouponimage")
	applyitemcount      = RequestCheckvar(request("applyitemcount"),10)
	openstate         	= RequestCheckvar(request("openstate"),10)
	margintype          = RequestCheckvar(request("margintype"),3)
	defaultmargin		= RequestCheckvar(request("defaultmargin"),10)
	mode 				= RequestCheckvar(request("mode"),16)
	itemidarr			= request("itemidarr")
	couponbuypricearr	= request("couponbuypricearr")
	itemcouponexplain	= html2Db(request("itemcouponexplain"))	
	makerid				= RequestCheckvar(request("makerid"),32)
	sailyn				= RequestCheckvar(request("sailyn"),1)
	sType               = RequestCheckvar(request("sType"),10)	
	addSql              = request("addSql")
	itemid              = request("itemid")
	itemname            = RequestCheckvar(request("itemname"),32)
	sellyn              = RequestCheckvar(request("sellyn"),1)
	usingyn             = RequestCheckvar(request("usingyn"),1)
	danjongyn           = RequestCheckvar(request("danjongyn"),1)
	limityn             = RequestCheckvar(request("limityn"),1)
	mwdiv               = RequestCheckvar(request("mwdiv"),2)
	cdl                 = RequestCheckvar(request("cdl"),3)
	cdm                 = RequestCheckvar(request("cdm"),3)
	cds                 = RequestCheckvar(request("cds"),3)
	deliverytype        = RequestCheckvar(request("deliverytype"),1)
	'response.write mode
	'response.end
  	if itemcouponname <> "" then
		if checkNotValidHTML(itemcouponname) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
  	if itemidarr <> "" then
		if checkNotValidHTML(itemidarr) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
  	if couponbuypricearr <> "" then
		if checkNotValidHTML(couponbuypricearr) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
	if itemcouponexplain <> "" then
		if checkNotValidHTML(itemcouponexplain) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
	if addSql <> "" then
		if checkNotValidHTML(addSql) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
	if itemid <> "" then
		if checkNotValidHTML(itemid) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
	if itemcouponidx="" then itemcouponidx="0"
	if defaultmargin="" then defaultmargin=0
	if (itemcouponidx<>"0") then
		IsEditMode = true
	else
		IsEditMode = false
	end if

'/���� ���
if mode="couponmaster" then
	
	on Error Resume Next
		buf = CDate(itemcouponstartdate)
		if Err then
			response.Write "<script>alert('�߱޽����� ����-" + Err.Description + "')</script>"
			response.Write "<script>history.back()</script>"
			dbacademyget.close()	:	response.End
		end if
	on Error Goto 0

	on Error Resume Next
		buf = CDate(itemcouponexpiredate)
		if Err then
			response.Write "<script>alert('�߱������� ����-" + Err.Description + "')</script>"
			response.Write "<script>history.back()</script>"
			dbacademyget.close()	:	response.End
		end if
	on Error Goto 0

	if (itemcoupontype="1") then
		if (itemcouponvalue>=100) or (itemcouponvalue<1) then
			response.Write "<script>alert('���������� 1~99% ���� ���� �����մϴ�.')</script>"
			response.Write "<script>history.back()</script>"
			dbacademyget.close()	:	response.End
		end if
	elseif (itemcoupontype="2") then
		if (itemcouponvalue<100) or (itemcouponvalue>=100000) then
			response.Write "<script>alert('���������� 1~100000 ���� ���� �����մϴ�.')</script>"
			response.Write "<script>history.back()</script>"
			dbacademyget.close()	:	response.End
		end if
	elseif (itemcoupontype="3") then
		if (itemcouponvalue<>2000) then
			response.Write "<script>alert('������ ���������� 2000 ���� �����մϴ�.')</script>"
			response.Write "<script>history.back()</script>"
			dbacademyget.close()	:	response.End
		end if
	else
		response.Write "<script>alert('����Ÿ���� �������� �ʾҽ��ϴ�.')</script>"
		response.Write "<script>history.back()</script>"
		dbacademyget.close()	:	response.End
	end if

	'/����
	if (IsEditMode) then		
		dim orgDefaultMargin ,orgDefaultMargintype
		
		sqlstr = "SELECT defaultmargin,margintype FROM db_academy.dbo.tbl_diy_item_coupon_master "
		sqlstr = sqlstr + " where itemcouponidx=" + CStr(itemcouponidx)

		'response.write sqlStr &"<Br>"
		rsacademyget.open sqlstr ,dbacademyget ,2

		IF not rsacademyget.eof Then
			orgDefaultMargin = rsacademyget("defaultmargin")
			orgDefaultMargintype = rsacademyget("margintype")
		End IF
		
		rsacademyget.close

		sqlstr = "update [db_academy].dbo.tbl_diy_item_coupon_master" + VbCrlf
		sqlstr = sqlstr + " set itemcoupontype='" + itemcoupontype + "'" + VbCrlf
		sqlstr = sqlstr + " ,couponGubun='" + couponGubun + "'" + VbCrlf
		sqlstr = sqlstr + " ,itemcouponvalue=" + CStr(itemcouponvalue) + VbCrlf
		sqlstr = sqlstr + " ,itemcouponstartdate='" + itemcouponstartdate + "'" + VbCrlf
		sqlstr = sqlstr + " ,itemcouponexpiredate='" + itemcouponexpiredate + "'" + VbCrlf
		sqlstr = sqlstr + " ,itemcouponname='" + itemcouponname + "'" + VbCrlf
		sqlstr = sqlstr + " ,itemcouponexplain='" + itemcouponexplain + "'" + VbCrlf
		sqlstr = sqlstr + " ,margintype='" + margintype + "'" + VbCrlf
		sqlstr = sqlstr + " ,defaultmargin='" + defaultmargin + "'" + VbCrlf
		sqlstr = sqlstr + " where itemcouponidx=" + CStr(itemcouponidx)

		'response.write sqlStr &"<Br>"
		dbacademyget.Execute sqlStr

		'���� ���� ����� ��� ��ǰ ��ü ����
		IF (Cint(orgDefaultMargin) <> Cint(defaultmargin)) or (CStr(orgDefaultMargintype)<>CStr(margintype)) Then
				
			sqlStr =" UPDATE db_academy.dbo.tbl_diy_item_coupon_detail  "&_
					" SET couponbuyprice="
			
			SELECT Case margintype
				Case "00"  	''��ǰ�������� - ���԰� 0 �ΰ�� �����԰�
					sqlStr = sqlStr + " 0 " + VbCrlf
				Case "10"	''�ٹ����ٺδ� - �����԰�
					sqlStr = sqlStr + " 0 " + VbCrlf
				Case "20"	''�������� : �߰� [2008-09-23]
					if itemcoupontype="1" then			''������
						sqlStr = sqlStr & " convert(int,i.sellcash*"& Cstr((100-itemcouponvalue)/100) &"*"& Cstr((100-defaultmargin)/100) &")"
					elseif itemcoupontype="2" then   	''�ݾ�
						sqlStr = sqlStr + " convert(int,(i.sellcash-" & CStr(itemcouponvalue) + ")*"& Cstr((100-defaultmargin)/100) &")"
					else
						sqlStr = sqlStr + " 0 " + VbCrlf
					end if
				Case "30"	''���ϸ��� - ���縶�� : �߰� [2008-09-23]
					if itemcoupontype="1" then			''������
						sqlStr = sqlStr + " convert(int,i.sellcash*" + CStr((100-itemcouponvalue)/100) + "*i.buycash/i.sellcash)"
					elseif itemcoupontype="2" then   	''�ݾ�
						sqlStr = sqlStr + " convert(int,(i.sellcash-" + CStr(itemcouponvalue) + ")*i.buycash/i.sellcash)"
					else
						sqlStr = sqlStr + " 0 " + VbCrlf
					end if
				Case "50"	''�ݹݺδ�
					if itemcoupontype="1" then			''������
						sqlStr = sqlStr + " i.buycash - convert(int,i.sellcash*" + CStr(itemcouponvalue/100) + "*0.5)"
					elseif itemcoupontype="2" then   	''�ݾ�
						sqlStr = sqlStr + " i.buycash - convert(int," + CStr(itemcouponvalue) + "*0.5)"
					else
						sqlStr = sqlStr + " 0 " + VbCrlf
					end if
				Case "60"	''��ü�δ� - ���԰� ����
					if itemcoupontype="1" then			''������
						sqlStr = sqlStr + " i.buycash - convert(int,i.sellcash*" + CStr(itemcouponvalue/100) + ")"
					elseif itemcoupontype="2" then   	''�ݾ�
						sqlStr = sqlStr + " i.buycash - " + CStr(itemcouponvalue)
					else
						sqlStr = sqlStr + " 0 " + VbCrlf
					end if
		        Case "80"   ''���������� -500
		            sqlStr = sqlStr + " case when i.mwdiv='M' then 0 else i.buycash - 500 end "
				Case "90"	''20%��ü��� - �����ΰ�� �����԰�.
					if itemcoupontype="1" then			''������
						sqlStr = sqlStr + " case when i.mwdiv='M' then 0 else i.buycash - convert(int,i.sellcash*" + CStr(itemcouponvalue/100) + "*0.5) end "
					elseif itemcoupontype="2" then   	''�ݾ�
						sqlStr = sqlStr + " case when i.mwdiv='M' 0 else i.buycash - convert(int," + CStr(itemcouponvalue) + "*0.5)  end "
					else
						sqlStr = sqlStr + " 0 " + VbCrlf
					end if
				Case else
					sqlStr = sqlStr + " 0 " + VbCrlf
			End SELECT
			sqlStr = sqlStr & " FROM db_academy.dbo.tbl_diy_item_coupon_detail d "
			sqlStr = sqlStr & " JOIN db_academy.dbo.tbl_diy_item i "
			sqlStr = sqlStr & " 	on d.itemid = i.itemid "
			sqlStr = sqlStr & " WHERE d.itemcouponidx=" & CStr(itemcouponidx)
			
			'response.write sqlStr &"<Br>"
			dbacademyget.Execute sqlStr
		End IF

	''�ű� ���
	else		
		sqlStr = "select * from [db_academy].dbo.tbl_diy_item_coupon_master where 1=0"
		rsacademyget.Open sqlStr,dbacademyget,1,3
		rsacademyget.AddNew

		rsacademyget("itemcoupontype") = itemcoupontype
		rsacademyget("couponGubun") = couponGubun
		rsacademyget("itemcouponvalue") = itemcouponvalue
		rsacademyget("itemcouponstartdate") = itemcouponstartdate
		rsacademyget("itemcouponexpiredate") = itemcouponexpiredate
		rsacademyget("itemcouponname") = itemcouponname
		rsacademyget("itemcouponexplain") = itemcouponexplain
		rsacademyget("openstate") = "0"
		rsacademyget("margintype") = margintype
		rsacademyget("defaultmargin")	= defaultmargin
		rsacademyget("reguserid") = session("ssBctId")

		rsacademyget.update
			itemcouponidx = rsacademyget("itemcouponidx")
		rsacademyget.close
	end if
	
elseif mode="I" then
    '' �߰� �˾�â���� �Ѿ� �� ���.
	ErrStr = ""

	''����Ÿ�� ��������
	margintype = "00"

	sqlStr = "select top 1 margintype, itemcoupontype, itemcouponvalue,"
	sqlStr = sqlStr + " convert(varchar(19),itemcouponstartdate,21) as itemcouponstartdate,"
	sqlStr = sqlStr + " convert(varchar(19),itemcouponexpiredate,21) as itemcouponexpiredate"
	sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item_coupon_master" + VbCrlf
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx)
	
	'response.write sqlStr &"<Br>"
	rsacademyget.Open sqlStr,dbacademyget
	
	if Not rsacademyget.Eof then
		margintype = rsacademyget("margintype")
		itemcoupontype = rsacademyget("itemcoupontype")
		itemcouponvalue = rsacademyget("itemcouponvalue")
		itemcouponstartdate = rsacademyget("itemcouponstartdate")
		itemcouponexpiredate = rsacademyget("itemcouponexpiredate")
	end if
	
	rsacademyget.close

	itemidarr = trim(itemidarr)
	if Right(itemidarr,1)="," then itemidarr=Left(itemidarr,Len(itemidarr)-1)

	'' ������ �����ϰ��, ��ü��ǰ �� �ٹ蹫���� ���رݾ� �ʰ� ��ǰ �ȳ�
	if itemcoupontype=3 then
		sqlStr = "Select top 100 itemid, mwdiv, sellcash " &_
				" from db_academy.dbo.tbl_diy_item " &_
				" Where itemid in (" & itemidarr & ")"
		
		'response.write sqlStr &"<Br>"
		rsacademyget.Open sqlStr,dbacademyget
		
		if not rsacademyget.Eof then
			do until rsacademyget.Eof
				if rsacademyget("mwdiv")="U" then ErrStr = ErrStr + "-��ü��� ��ǰ (��ǰ��ȣ : " + CStr(rsacademyget("itemid")) + ") ��ϺҰ� \n"
				if rsacademyget("mwdiv")<>"U" and rsacademyget("sellcash")>=30000 then ErrStr = ErrStr + "- ������ ��ǰ (��ǰ��ȣ : " + CStr(rsacademyget("itemid")) + ") ��ϺҰ� \n"
				rsacademyget.moveNext
			loop

			if ErrStr<>"" then
				response.write "<script language=javascript>alert('��۷����� ��������\n\n" + ErrStr + "');</script>"
				response.End
			end if
		end if
		
		rsacademyget.close
	end if

    ''�˻��� ��ü ��ǰ�� ���.. �˻��� ��� ���� insert  ó��
    addSql = ""
    IF (sType="all") THEN

         '// �߰� ����
        if (makerid <> "") then
            addSql = addSql & " and i.makerid='" + makerid + "'"
        end if

        if (itemid <> "") then
            addSql = addSql & " and i.itemid in (" + itemid + ")"
        end if

        if (itemname <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(itemname) + "%'"
        end if

        if (sellyn <> "") then
            addSql = addSql & " and i.sellyn='" + sellyn + "'"
        end if

        if (usingyn <> "") then
            addSql = addSql & " and i.isusing='" + usingyn + "'"
        end if

        if danjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif danjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + danjongyn + "'"
        end if

		if limityn="Y0" then
            addSql = addSql + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif limityn<>"" then
            addSql = addSql + " and i.limityn='" + limityn + "'"
        end if

        if mwdiv="MW" then
            addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif mwdiv<>"" then
            addSql = addSql + " and i.mwdiv='" + mwdiv + "'"
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

        if sailyn<>"" then
            addSql = addSql + " and i.sailyn='" + sailyn + "'"
        end if

        if deliverytype <> "" then
        	addSql = addSql + " and i.deliverytype='" + deliverytype + "'"
        end if

        if (addSql="") then
            addSql = "select i.itemid from [db_academy].dbo.tbl_diy_item i where 1=0 "
        else
            addSql = "select i.itemid from [db_academy].dbo.tbl_diy_item i where 1=1 " & addSql
        end if
    ELSE
    	addSql = trim(itemidarr)
	END IF

	'' �ٸ� ������ ��ǰ�� ��ϵǾ� ������� üũ
	sqlStr = " select top 100 m.itemcouponidx, d.itemid from"
	sqlStr = sqlStr + " [db_academy].dbo.tbl_diy_item_coupon_master m,"
	sqlStr = sqlStr + " [db_academy].dbo.tbl_diy_item_coupon_detail d"
	sqlStr = sqlStr + " where m.itemcouponidx=d.itemcouponidx"
	sqlStr = sqlStr + " and m.itemcouponidx<>" + CStr(itemcouponidx)
	sqlStr = sqlStr + " and m.openstate<9"  ''�߱������ΰ� ����
	sqlStr = sqlStr + " and ( "
	sqlStr = sqlStr + " 	(m.itemcouponstartdate<='" + CStr(itemcouponstartdate) + "' and m.itemcouponexpiredate>'" + CStr(itemcouponstartdate) + "')"
	sqlStr = sqlStr + " 	or "
	sqlStr = sqlStr + " 	(m.itemcouponstartdate<='" + CStr(itemcouponexpiredate) + "' and m.itemcouponexpiredate>'" + CStr(itemcouponexpiredate) + "')"
	sqlStr = sqlStr + " 	)"
	sqlStr = sqlStr + " and d.itemid in (" + addSql + ")"  + VbCrlf

	'response.write sqlStr &"<Br>"
	rsacademyget.Open sqlStr,dbacademyget
	
	if not rsacademyget.Eof then
		do until rsacademyget.Eof
			ErrStr = ErrStr + "������ȣ : " + CStr(rsacademyget("itemcouponidx")) + " - ��ǰ��ȣ : " + CStr(rsacademyget("itemid")) + " ����� \n"
			rsacademyget.moveNext
		loop
	end if
	
	rsacademyget.close

	'' ���� ��Ͽ� ���� ��ǰ�� �߰�.
	sqlStr = "insert into [db_academy].dbo.tbl_diy_item_coupon_detail" + VbCrlf
	sqlStr = sqlStr + " (itemcouponidx, itemid, couponbuyprice)" + VbCrlf
	sqlStr = sqlStr + " select " + CStr(itemcouponidx) + "," + VbCrlf
	sqlStr = sqlStr + " i.itemid, " + VbCrlf
	
	Select Case margintype
		Case "00"  	''��ǰ�������� - ���԰� 0 �ΰ�� �����԰�
			sqlStr = sqlStr + " 0 " + VbCrlf
		'Case "10"	''�ΰŽ��δ� - ���԰� ����x
		'	if itemcoupontype="1" then			''������
		'		sqlStr = sqlStr + " i.buycash - convert(int,i.sellcash*" + CStr(itemcouponvalue/100) + ")"
		'	elseif itemcoupontype="2" then   	''�ݾ�
		'		sqlStr = sqlStr + " i.buycash - " + CStr(itemcouponvalue)
		'	else
		'		sqlStr = sqlStr + " 0 " + VbCrlf
		'	end if
		Case "10"	''�ΰŽ��δ� - �����԰�
			sqlStr = sqlStr + " 0 " + VbCrlf

		Case "20"	''�������� : �߰� [2008-09-23]
			if itemcoupontype="1" then			''������
				sqlStr = sqlStr & " convert(int,i.sellcash*"& Cstr((100-itemcouponvalue)/100) &"*"& Cstr((100-defaultmargin)/100) &")"
				'response.Write "<javascript language=javascript>alert(' convert(int,i.sellcash*"& Cstr((100-itemcouponvalue)/100) &"*"& Cstr((100-defaultmargin)/100) &") ')</script>"
				'response.end
			elseif itemcoupontype="2" then   	''�ݾ�
				sqlStr = sqlStr + " convert(int,(i.sellcash-" & CStr(itemcouponvalue) + ")*"& Cstr((100-defaultmargin)/100) &")"
			else
				sqlStr = sqlStr + " 0 " + VbCrlf
			end if
		Case "30"	''���ϸ��� - ���縶�� : �߰� [2008-09-23]
			if itemcoupontype="1" then			''������
				sqlStr = sqlStr + " convert(int,i.sellcash*" + CStr((100-itemcouponvalue)/100) + "*i.buycash/i.sellcash)"
			elseif itemcoupontype="2" then   	''�ݾ�
				sqlStr = sqlStr + " convert(int,(i.sellcash-" + CStr(itemcouponvalue) + ")*i.buycash/i.sellcash)"
			else
				sqlStr = sqlStr + " 0 " + VbCrlf
			end if
		Case "50"	''�ݹݺδ�
			if itemcoupontype="1" then			''������
				sqlStr = sqlStr + " i.buycash - convert(int,i.sellcash*" + CStr(itemcouponvalue/100) + "*0.5)"
			elseif itemcoupontype="2" then   	''�ݾ�
				sqlStr = sqlStr + " i.buycash - convert(int," + CStr(itemcouponvalue) + "*0.5)"
			else
				sqlStr = sqlStr + " 0 " + VbCrlf
			end if
		Case "60"	''��ü�δ� - ���԰� ����
			if itemcoupontype="1" then			''������
				sqlStr = sqlStr + " i.buycash - convert(int,i.sellcash*" + CStr(itemcouponvalue/100) + ")"
			elseif itemcoupontype="2" then   	''�ݾ�
				sqlStr = sqlStr + " i.buycash - " + CStr(itemcouponvalue)
			else
				sqlStr = sqlStr + " 0 " + VbCrlf
			end if
        Case "80"   ''���������� -500
            sqlStr = sqlStr + " case when i.mwdiv='M' then 0 else i.buycash - 500 end "
		Case "90"	''20%��ü��� - �����ΰ�� �����԰�.
			if itemcoupontype="1" then			''������
				sqlStr = sqlStr + " case when i.mwdiv='M' then 0 else i.buycash - convert(int,i.sellcash*" + CStr(itemcouponvalue/100) + "*0.5) end "
			elseif itemcoupontype="2" then   	''�ݾ�
				sqlStr = sqlStr + " case when i.mwdiv='M' 0 else i.buycash - convert(int," + CStr(itemcouponvalue) + "*0.5)  end "
			else
				sqlStr = sqlStr + " 0 " + VbCrlf
			end if
		Case else
			sqlStr = sqlStr + " 0 " + VbCrlf
	end Select

	sqlStr = sqlStr + " from [db_academy].dbo.tbl_diy_item i" + VbCrlf
	sqlStr = sqlStr + " left join [db_academy].dbo.tbl_diy_item_coupon_detail d" + VbCrlf
	sqlStr = sqlStr + " 	on d.itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
	sqlStr = sqlStr + " 	and d.itemid=i.itemid" + VbCrlf
	sqlStr = sqlStr + " where i.itemid in (" + addSql + ")"  + VbCrlf
	sqlStr = sqlStr + " and d.itemid is null"
	sqlStr = sqlStr + " and i.itemid not in ("
	sqlStr = sqlStr + " 	select distinct d.itemid from"
	sqlStr = sqlStr + " 	[db_academy].dbo.tbl_diy_item_coupon_master m,"
	sqlStr = sqlStr + " 	[db_academy].dbo.tbl_diy_item_coupon_detail d"
	sqlStr = sqlStr + " 	where m.itemcouponidx=d.itemcouponidx"
	sqlStr = sqlStr + " 	and m.itemcouponidx<>" + CStr(itemcouponidx)
	sqlStr = sqlStr + " 	and m.openstate<9"  ''�߱������ΰ� ����
	sqlStr = sqlStr + " 	and ( "
	sqlStr = sqlStr + " 		(m.itemcouponstartdate<='" + CStr(itemcouponstartdate) + "' and m.itemcouponexpiredate>'" + CStr(itemcouponstartdate) + "')"
	sqlStr = sqlStr + " 		or "
	sqlStr = sqlStr + " 		(m.itemcouponstartdate<='" + CStr(itemcouponexpiredate) + "' and m.itemcouponexpiredate>'" + CStr(itemcouponexpiredate) + "')"
	sqlStr = sqlStr + " 		)"
	sqlStr = sqlStr + " 	and d.itemid in (" + addSql + ")"  + VbCrlf
	sqlStr = sqlStr + " ) "

	'response.write sqlStr &"<Br>"
	rsacademyget.Open sqlStr,dbacademyget,1

	''�����ǰ��.
	AplyItemCountUpdate itemcouponidx
	AplyToItem itemcouponidx
	
elseif mode="delcouponitemarr" then
	itemidarr = trim(itemidarr)
	if Right(itemidarr,1)="," then itemidarr=Left(itemidarr,Len(itemidarr)-1)

	sqlStr = "delete from [db_academy].dbo.tbl_diy_item_coupon_detail" + VbCrlf
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
	sqlStr = sqlStr + " and itemid in (" + itemidarr + ")"  + VbCrlf
	
	'response.write sqlStr &"<Br>"
	rsacademyget.Open sqlStr,dbacademyget,1

	''�����ǰ��.
	AplyItemCountUpdate itemcouponidx

	''������ ���� ��ǰ���̺��� ���� ���� N �� ����
	AplyToItem itemcouponidx
	
elseif mode="modicouponitemarr" then
	itemidarr = trim(itemidarr)
	couponbuypricearr  = trim(couponbuypricearr)

	if Right(itemidarr,1)="," then itemidarr=Left(itemidarr,Len(itemidarr)-1)
	if Right(couponbuypricearr,1)="," then couponbuypricearr=Left(couponbuypricearr,Len(couponbuypricearr)-1)

	itemidarr = split(itemidarr,",")
	couponbuypricearr = split(couponbuypricearr,",")

	for i=LBound(itemidarr) to UBound(itemidarr)
		if trim(itemidarr(i))<>"" then
			sqlStr = "update [db_academy].dbo.tbl_diy_item_coupon_detail" + VbCrlf
			sqlStr = sqlStr + " set couponbuyprice=" + CStr(couponbuypricearr(i)) + VbCrlf
			sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
			sqlStr = sqlStr + " and itemid=" + CStr(itemidarr(i)) + VbCrlf
			
			'response.write sqlStr &"<Br>"
			rsacademyget.Open sqlStr,dbacademyget,1
		end if
	next

	''�����ǰ��.
	AplyItemCountUpdate itemcouponidx
	AplyToItem itemcouponidx
	
elseif mode="opencoupon" Then

	sqlStr = "update [db_academy].dbo.tbl_diy_item_coupon_master" + VbCrlf
	sqlStr = sqlStr + " set openstate='7'"
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf

	'response.write sqlStr &"<Br>"
	rsacademyget.Open sqlStr,dbacademyget,1

	AplyToItem(itemcouponidx)

elseif mode="reservecoupon" Then

	sqlStr = "update [db_academy].dbo.tbl_diy_item_coupon_master" + VbCrlf
	sqlStr = sqlStr + " set openstate='6'"
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
	
	'response.write sqlStr &"<Br>"
	rsacademyget.Open sqlStr,dbacademyget,1

elseif mode="closecoupon" Then

    dim MayExpireDt
    MayExpireDt = Left(CStr(DateAdd("d",-1,Now())),10) & " 23:59:59"

    ''response.write MayExpireDt

    ''�� �߱� �� ���� Expire
    sqlStr = "update [db_academy].dbo.tbl_user_diy_item_coupon" + VbCrlf
    sqlStr = sqlStr + " set itemcouponexpiredate='" & MayExpireDt & "'" + VbCrlf
    sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
    sqlStr = sqlStr + " and itemcouponexpiredate>'" & MayExpireDt & "'" + VbCrlf
    sqlStr = sqlStr + " and usedyn='N'" + VbCrlf
	
	'response.write sqlStr &"<Br>"
    dbacademyget.Execute sqlStr

	sqlStr = "update [db_academy].dbo.tbl_diy_item_coupon_master" + VbCrlf
	sqlStr = sqlStr + " set openstate='9'"
	sqlStr = sqlStr + " where itemcouponidx=" + CStr(itemcouponidx) + VbCrlf
	
	'response.write sqlStr &"<Br>"
	dbacademyget.Execute sqlStr

	AplyToItem(itemcouponidx)
end if
%>
<% if (mode="couponmaster") then %>
	<% if (IsEditMode) then %>
		<script language='javascript'>
			alert('���� �Ǿ����ϴ�.');
			location.replace('/academy/shopmaster/itemcouponmasterreg.asp?itemcouponidx=<%= itemcouponidx %>');
		</script>
	<% else %>
		<script language='javascript'>
			alert('���� �Ǿ����ϴ�. ��ǰ�� ��� �� �ּ���');
			opener.location.reload();
			window.close();	
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
		//location.replace('/academy/shopmaster/itemcouponitemlistedit.asp?itemcouponidx=<%= itemcouponidx %>&makerid=<%= makerid %>&sailyn=<%= sailyn %>');
	</script>
	
<% elseif mode="delcouponitemarr" then %>
	<script language='javascript'>
		alert('���� �Ǿ����ϴ�.');
		opener.location.reload();
		location.replace('/academy/shopmaster/itemcouponitemlistedit.asp?itemcouponidx=<%= itemcouponidx %>&makerid=<%= makerid %>&sailyn=<%= sailyn %>');
	</script>
	
<% elseif mode="modicouponitemarr" then %>
	<script language='javascript'>
		alert('���� �Ǿ����ϴ�.');
		opener.location.reload();
		location.replace('/academy/shopmaster/itemcouponitemlistedit.asp?itemcouponidx=<%= itemcouponidx %>&makerid=<%= makerid %>&sailyn=<%= sailyn %>');
	</script>
	
<% elseif mode="opencoupon" then %>
	<script language='javascript'>
		alert('������ ���� �Ǿ����ϴ�.');
		opener.location.reload();
		location.replace('/academy/shopmaster/itemcouponmasterreg.asp?itemcouponidx=<%= itemcouponidx %>');
	</script>
	
<% elseif mode="reservecoupon" then %>
	<script language='javascript'>
		alert('������ ������ ���� �Ǿ����ϴ�. ���� 0�ÿ� ����˴ϴ�.');
		opener.location.reload();
		location.replace('/academy/shopmaster/itemcouponmasterreg.asp?itemcouponidx=<%= itemcouponidx %>');
	</script>
	
<% elseif mode="closecoupon" then %>
	<script language='javascript'>
		alert('������ ���� �Ǿ����ϴ�.');
		opener.location.reload();
		location.replace('/academy/shopmaster/itemcouponmasterreg.asp?itemcouponidx=<%= itemcouponidx %>');
		self.close();
	</script>
<% end if %>

<%= "mode=" + mode %>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->