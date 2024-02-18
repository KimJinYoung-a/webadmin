<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���� ����
' History : 2010.10.11 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
function AplyItemCountUpdate(lecturercouponidx)
	dim sqlStr
	
	''�����ǰ���� ������Ʈ
	sqlStr = "update [db_academy].dbo.tbl_lecturer_coupon_master" + VbCrlf
	sqlStr = sqlStr + " set applyitemcount=IsNULL(T.cnt,0)" + VbCrlf
	sqlStr = sqlStr + " from (" + VbCrlf
	sqlStr = sqlStr + "     select count(*) as cnt from [db_academy].dbo.tbl_lecturer_coupon_detail where lecturercouponidx=" + CStr(lecturercouponidx) + VbCrlf
	sqlStr = sqlStr + " ) as T" + VbCrlf
	sqlStr = sqlStr + " where lecturercouponidx=" + CStr(lecturercouponidx) + VbCrlf
	
	'response.write sqlStr & "<br>"
	dbacademyget.Execute sqlStr
end function

function AplyToItem(lecturercouponidx)
	dim sqlStr
	dim ocouponGubun, olecturercoupontype, olecturercouponvalue, olecturercouponstartdate, olecturercouponexpiredate, openstate, currdatetime
	dim couponExpired
	dim resultCnt

	applyitemcount = 0
	couponExpired = false

	sqlStr = "select top 1 couponGubun, margintype, lecturercoupontype, lecturercouponvalue, openstate, applyitemcount,"
	sqlStr = sqlStr + " convert(varchar(19),lecturercouponstartdate,21) as lecturercouponstartdate,"
	sqlStr = sqlStr + " convert(varchar(19),lecturercouponexpiredate,21) as lecturercouponexpiredate,"
	sqlStr = sqlStr + " (case when (lecturercouponstartdate>getdate()) or (lecturercouponexpiredate<getdate()) then 'Y' else 'N' end ) as couponexpired, "
	sqlStr = sqlStr + " convert(varchar(19),getdate()) as currdatetime"
	sqlStr = sqlStr + " from [db_academy].dbo.tbl_lecturer_coupon_master" + VbCrlf
	sqlStr = sqlStr + " where lecturercouponidx=" + CStr(lecturercouponidx)

	rsacademyget.Open sqlStr,dbacademyget,1
	if Not rsacademyget.Eof then
	    ocouponGubun   = rsacademyget("couponGubun")
		lecturercoupontype = rsacademyget("lecturercoupontype")
		lecturercouponvalue = rsacademyget("lecturercouponvalue")
		lecturercouponstartdate = rsacademyget("lecturercouponstartdate")
		lecturercouponexpiredate = rsacademyget("lecturercouponexpiredate")
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
		sqlStr = "update [db_academy].dbo.tbl_lec_item"
		sqlStr = sqlStr + " set lecturercouponyn='N'"
		sqlStr = sqlStr + " ,lecturercoupontype='1'"
		sqlStr = sqlStr + " ,lecturercouponvalue=0"
		sqlStr = sqlStr + " ,currlecturercouponidx=NULL"
		sqlStr = sqlStr + " ,lastupdate=getdate()"
		sqlStr = sqlStr + " from [db_academy].dbo.tbl_lecturer_coupon_detail"
		sqlStr = sqlStr + " where lecturercouponidx=" + CStr(lecturercouponidx)
		sqlStr = sqlStr + " and [db_academy].dbo.tbl_lec_item.idx=[db_academy].dbo.tbl_lecturer_coupon_detail.lectureridx"

		'response.write sqlStr + "<br>"
		dbacademyget.Execute sqlStr
	end if

	''��ǰ�� �����Ȱ�� -> N�� ����
	sqlStr = "update [db_academy].dbo.tbl_lec_item"
	sqlStr = sqlStr + " set lecturercouponyn='N'"
	sqlStr = sqlStr + " ,lecturercoupontype='1'"
	sqlStr = sqlStr + " ,lecturercouponvalue=0"
	sqlStr = sqlStr + " ,currlecturercouponidx=NULL"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select i.idx  "
	sqlStr = sqlStr + " 	from [db_academy].dbo.tbl_lec_item i"
	sqlStr = sqlStr + " 	left join [db_academy].dbo.tbl_lecturer_coupon_detail d"
	sqlStr = sqlStr + " 	on d.lecturercouponidx=" + CStr(lecturercouponidx) + " and i.idx=d.lectureridx "
	sqlStr = sqlStr + " 	where i.currlecturercouponidx=" + CStr(lecturercouponidx)
	sqlStr = sqlStr + " 	and d.lecturercouponidx is null"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_academy].dbo.tbl_lec_item.idx=T.idx"

	'response.write sqlStr + "<br>"
	dbacademyget.Execute sqlStr, resultCnt
	response.write "�����Ǽ�=" + CStr(resultCnt) + "<br>"

	''lecturercouponidx�� ��ϵ� ��ǰ�� ��� �������� ������ Update
	sqlStr = "update [db_academy].dbo.tbl_lec_item"
	sqlStr = sqlStr + " set lecturercouponyn='Y'"
	sqlStr = sqlStr + " ,lecturercoupontype=T.lecturercoupontype"
	sqlStr = sqlStr + " ,lecturercouponvalue=T.lecturercouponvalue"
	sqlStr = sqlStr + " ,currlecturercouponidx=T.lecturercouponidx"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select m.lecturercouponidx, m.lecturercoupontype, m.lecturercouponvalue, d.lectureridx "
	sqlStr = sqlStr + " 	from [db_academy].dbo.tbl_lecturer_coupon_master m,"
	sqlStr = sqlStr + " 	[db_academy].dbo.tbl_lecturer_coupon_detail d"
	sqlStr = sqlStr + " 	where m.lecturercouponidx=d.lecturercouponidx"
	sqlStr = sqlStr + " 	and m.openstate='7'"
	sqlStr = sqlStr + " 	and d.lecturercouponidx=" + CStr(lecturercouponidx)
	sqlStr = sqlStr + " 	and m.lecturercouponstartdate<=getdate()"
	sqlStr = sqlStr + " 	and m.lecturercouponexpiredate>=getdate()"
	sqlStr = sqlStr + " ) T "
	sqlStr = sqlStr + " where [db_academy].dbo.tbl_lec_item.idx=T.lectureridx"
	sqlStr = sqlStr + " and Not ("
	sqlStr = sqlStr + " 		 	[db_academy].dbo.tbl_lec_item.lecturercouponyn='Y'"
	sqlStr = sqlStr + " 		and [db_academy].dbo.tbl_lec_item.lecturercoupontype=T.lecturercoupontype"
	sqlStr = sqlStr + " 		and [db_academy].dbo.tbl_lec_item.lecturercouponvalue=T.lecturercouponvalue"
	sqlStr = sqlStr + " 		and [db_academy].dbo.tbl_lec_item.currlecturercouponidx=T.lecturercouponidx"
	sqlStr = sqlStr + "			)"

	'response.write sqlStr + "<br>"
	dbacademyget.Execute sqlStr, resultCnt

    response.write "�����Ǽ�=" + CStr(resultCnt)
end function

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

dim lecturercouponvalue ,lecturercouponstartdate ,lecturercoupontype ,couponGubun ,lecturercouponidx
dim openstate ,margintype ,applyitemcount ,lecturercouponexplain ,lecturercouponimage ,lecturercouponname ,lecturercouponexpiredate
dim lectureridxarr, couponbuypricearr, makerid, sailyn ,ErrStr ,buf ,sqlstr,i ,IsEditMode ,mode ,defaultmargin
dim sType, addSql, itemid, itemname, sellyn, usingyn, danjongyn ,limityn, mwdiv, cdl, cdm, cds, deliverytype
	lecturercouponidx      	= requestCheckVar(request("lecturercouponidx"),9)
	couponGubun         = requestCheckVar(request("couponGubun"),9)
	lecturercoupontype      = requestCheckVar(request("lecturercoupontype"),9)
	lecturercouponvalue     = requestCheckVar(request("lecturercouponvalue"),9)
	lecturercouponstartdate = requestCheckVar(request("lecturercouponstartdate") + " " + request("lecturercouponstartdate2"),32)
	lecturercouponexpiredate= requestCheckVar(request("lecturercouponexpiredate") + " " + request("lecturercouponexpiredate2"),32)
	lecturercouponname      = html2Db(request("lecturercouponname"))
	lecturercouponimage     = requestCheckVar(request("lecturercouponimage"),16)
	applyitemcount      = requestCheckVar(request("applyitemcount"),10)
	openstate         	= requestCheckVar(request("openstate"),1)
	margintype          = requestCheckVar(request("margintype"),2)
	defaultmargin		= requestCheckVar(request("defaultmargin"),6)
	mode 				= requestCheckVar(request("mode"),16)
	lectureridxarr			= request("lectureridxarr")
	couponbuypricearr	= request("couponbuypricearr")
	lecturercouponexplain	= html2Db(request("lecturercouponexplain"))	
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
	mwdiv               = RequestCheckvar(request("mwdiv"),1)
	cdl                 = RequestCheckvar(request("cdl"),3)
	cdm                 = RequestCheckvar(request("cdm"),3)
	cds                 = RequestCheckvar(request("cds"),3)
	deliverytype        = RequestCheckvar(request("deliverytype"),1)
	'response.write mode
	'response.end
  	if lecturercouponname <> "" then
		if checkNotValidHTML(lecturercouponname) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
	if lecturercouponexplain <> "" then
		if checkNotValidHTML(lecturercouponexplain) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
	if lectureridxarr <> "" then
		if checkNotValidHTML(lectureridxarr) then
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
	if itemid <> "" then
		if checkNotValidHTML(itemid) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end If

	if lecturercouponidx="" then lecturercouponidx="0"
	if defaultmargin="" then defaultmargin=0
	if (lecturercouponidx<>"0") then
		IsEditMode = true
	else
		IsEditMode = false
	end if

'/���� ���
if mode="couponmaster" then
	
	on Error Resume Next
		buf = CDate(lecturercouponstartdate)
		if Err then
			response.Write "<script>alert('�߱޽����� ����-" + Err.Description + "')</script>"
			response.Write "<script>history.back()</script>"
			dbacademyget.close()	:	response.End
		end if
	on Error Goto 0

	on Error Resume Next
		buf = CDate(lecturercouponexpiredate)
		if Err then
			response.Write "<script>alert('�߱������� ����-" + Err.Description + "')</script>"
			response.Write "<script>history.back()</script>"
			dbacademyget.close()	:	response.End
		end if
	on Error Goto 0

	if (lecturercoupontype="1") then
		if (lecturercouponvalue>=100) or (lecturercouponvalue<1) then
			response.Write "<script>alert('���������� 1~99% ���� ���� �����մϴ�.')</script>"
			response.Write "<script>history.back()</script>"
			dbacademyget.close()	:	response.End
		end if
	elseif (lecturercoupontype="2") then
		if (lecturercouponvalue<100) or (lecturercouponvalue>=100000) then
			response.Write "<script>alert('���������� 1~100000 ���� ���� �����մϴ�.')</script>"
			response.Write "<script>history.back()</script>"
			dbacademyget.close()	:	response.End
		end if
	elseif (lecturercoupontype="3") then
		if (lecturercouponvalue<>2000) then
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
		
		sqlstr = "SELECT defaultmargin,margintype FROM db_academy.dbo.tbl_lecturer_coupon_master "
		sqlstr = sqlstr + " where lecturercouponidx=" + CStr(lecturercouponidx)

		'response.write sqlStr &"<Br>"
		rsacademyget.open sqlstr ,dbacademyget ,2

		IF not rsacademyget.eof Then
			orgDefaultMargin = rsacademyget("defaultmargin")
			orgDefaultMargintype = rsacademyget("margintype")
		End IF
		
		rsacademyget.close

		sqlstr = "update [db_academy].dbo.tbl_lecturer_coupon_master" + VbCrlf
		sqlstr = sqlstr + " set lecturercoupontype='" + lecturercoupontype + "'" + VbCrlf
		sqlstr = sqlstr + " ,couponGubun='" + couponGubun + "'" + VbCrlf
		sqlstr = sqlstr + " ,lecturercouponvalue=" + CStr(lecturercouponvalue) + VbCrlf
		sqlstr = sqlstr + " ,lecturercouponstartdate='" + lecturercouponstartdate + "'" + VbCrlf
		sqlstr = sqlstr + " ,lecturercouponexpiredate='" + lecturercouponexpiredate + "'" + VbCrlf
		sqlstr = sqlstr + " ,lecturercouponname='" + lecturercouponname + "'" + VbCrlf
		sqlstr = sqlstr + " ,lecturercouponexplain='" + lecturercouponexplain + "'" + VbCrlf
		sqlstr = sqlstr + " ,margintype='" + margintype + "'" + VbCrlf
		sqlstr = sqlstr + " ,defaultmargin='" + defaultmargin + "'" + VbCrlf
		sqlstr = sqlstr + " where lecturercouponidx=" + CStr(lecturercouponidx)

		'response.write sqlStr &"<Br>"
		dbacademyget.Execute sqlStr

		'���� ���� ����� ��� ��ǰ ��ü ����
		IF (Cint(orgDefaultMargin) <> Cint(defaultmargin)) or (CStr(orgDefaultMargintype)<>CStr(margintype)) Then
				
			sqlStr =" UPDATE db_academy.dbo.tbl_lecturer_coupon_detail  "&_
					" SET couponbuyprice="
			
			SELECT Case margintype
				Case "00"  	''��ǰ�������� - ���԰� 0 �ΰ�� �����԰�
					sqlStr = sqlStr + " 0 " + VbCrlf
				Case "10"	''�ΰŽ��δ� - �����԰�
					sqlStr = sqlStr + " 0 " + VbCrlf
				Case "20"	''�������� : �߰� [2008-09-23]
					if lecturercoupontype="1" then			''������
						sqlStr = sqlStr & " convert(int,i.sellcash*"& Cstr((100-lecturercouponvalue)/100) &"*"& Cstr((100-defaultmargin)/100) &")"
					elseif lecturercoupontype="2" then   	''�ݾ�
						sqlStr = sqlStr + " convert(int,(i.sellcash-" & CStr(lecturercouponvalue) + ")*"& Cstr((100-defaultmargin)/100) &")"
					else
						sqlStr = sqlStr + " 0 " + VbCrlf
					end if
				Case "30"	''���ϸ��� - ���縶�� : �߰� [2008-09-23]
					if lecturercoupontype="1" then			''������
						sqlStr = sqlStr + " convert(int,i.sellcash*" + CStr((100-lecturercouponvalue)/100) + "*i.buycash/i.sellcash)"
					elseif lecturercoupontype="2" then   	''�ݾ�
						sqlStr = sqlStr + " convert(int,(i.sellcash-" + CStr(lecturercouponvalue) + ")*i.buycash/i.sellcash)"
					else
						sqlStr = sqlStr + " 0 " + VbCrlf
					end if
				Case "50"	''�ݹݺδ�
					if lecturercoupontype="1" then			''������
						sqlStr = sqlStr + " i.buycash - convert(int,i.sellcash*" + CStr(lecturercouponvalue/100) + "*0.5)"
					elseif lecturercoupontype="2" then   	''�ݾ�
						sqlStr = sqlStr + " i.buycash - convert(int," + CStr(lecturercouponvalue) + "*0.5)"
					else
						sqlStr = sqlStr + " 0 " + VbCrlf
					end if
				Case "60"	''��ü�δ� - ���԰� ����
					if lecturercoupontype="1" then			''������
						sqlStr = sqlStr + " i.buycash - convert(int,i.sellcash*" + CStr(lecturercouponvalue/100) + ")"
					elseif lecturercoupontype="2" then   	''�ݾ�
						sqlStr = sqlStr + " i.buycash - " + CStr(lecturercouponvalue)
					else
						sqlStr = sqlStr + " 0 " + VbCrlf
					end if
		        Case "80"   ''���������� -500
		            sqlStr = sqlStr + " case when i.mwdiv='M' then 0 else i.buycash - 500 end "
				Case "90"	''20%��ü��� - �����ΰ�� �����԰�.
					if lecturercoupontype="1" then			''������
						sqlStr = sqlStr + " case when i.mwdiv='M' then 0 else i.buycash - convert(int,i.sellcash*" + CStr(lecturercouponvalue/100) + "*0.5) end "
					elseif lecturercoupontype="2" then   	''�ݾ�
						sqlStr = sqlStr + " case when i.mwdiv='M' 0 else i.buycash - convert(int," + CStr(lecturercouponvalue) + "*0.5)  end "
					else
						sqlStr = sqlStr + " 0 " + VbCrlf
					end if
				Case else
					sqlStr = sqlStr + " 0 " + VbCrlf
			End SELECT
			sqlStr = sqlStr & " FROM db_academy.dbo.tbl_lecturer_coupon_detail d "
			sqlStr = sqlStr & " JOIN db_academy.dbo.tbl_lec_item i "
			sqlStr = sqlStr & " 	on d.itemid = i.itemid "
			sqlStr = sqlStr & " WHERE d.lecturercouponidx=" & CStr(lecturercouponidx)
			
			'response.write sqlStr &"<Br>"
			dbacademyget.Execute sqlStr
		End IF

	''�ű� ���
	else		
		sqlStr = "select * from [db_academy].dbo.tbl_lecturer_coupon_master where 1=0"
		rsacademyget.Open sqlStr,dbacademyget,1,3
		rsacademyget.AddNew

		rsacademyget("lecturercoupontype") = lecturercoupontype
		rsacademyget("couponGubun") = couponGubun
		rsacademyget("lecturercouponvalue") = lecturercouponvalue
		rsacademyget("lecturercouponstartdate") = lecturercouponstartdate
		rsacademyget("lecturercouponexpiredate") = lecturercouponexpiredate
		rsacademyget("lecturercouponname") = lecturercouponname
		rsacademyget("lecturercouponexplain") = lecturercouponexplain
		rsacademyget("openstate") = "0"
		rsacademyget("margintype") = margintype
		rsacademyget("defaultmargin")	= defaultmargin
		rsacademyget("reguserid") = session("ssBctId")

		rsacademyget.update
			lecturercouponidx = rsacademyget("lecturercouponidx")
		rsacademyget.close
	end if
	
elseif mode="I" then
    '' �߰� �˾�â���� �Ѿ� �� ���.
	ErrStr = ""

	''����Ÿ�� ��������
	margintype = "00"

	sqlStr = "select top 1 margintype, lecturercoupontype, lecturercouponvalue,"
	sqlStr = sqlStr + " convert(varchar(19),lecturercouponstartdate,21) as lecturercouponstartdate,"
	sqlStr = sqlStr + " convert(varchar(19),lecturercouponexpiredate,21) as lecturercouponexpiredate"
	sqlStr = sqlStr + " from [db_academy].dbo.tbl_lecturer_coupon_master" + VbCrlf
	sqlStr = sqlStr + " where lecturercouponidx=" + CStr(lecturercouponidx)
	
	'response.write sqlStr &"<Br>"
	rsacademyget.Open sqlStr,dbacademyget
	
	if Not rsacademyget.Eof then
		margintype = rsacademyget("margintype")
		lecturercoupontype = rsacademyget("lecturercoupontype")
		lecturercouponvalue = rsacademyget("lecturercouponvalue")
		lecturercouponstartdate = rsacademyget("lecturercouponstartdate")
		lecturercouponexpiredate = rsacademyget("lecturercouponexpiredate")
	end if
	
	rsacademyget.close

	lectureridxarr = trim(lectureridxarr)
	if Right(lectureridxarr,1)="," then lectureridxarr=Left(lectureridxarr,Len(lectureridxarr)-1)

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
            addSql = "select i.itemid from [db_academy].dbo.tbl_lec_item i where 1=0 "
        else
            addSql = "select i.itemid from [db_academy].dbo.tbl_lec_item i where 1=1 " & addSql
        end if
    ELSE
    	addSql = trim(lectureridxarr)
	END IF

	'' �ٸ� ������ ��ǰ�� ��ϵǾ� ������� üũ
	sqlStr = " select top 100 m.lecturercouponidx, d.lectureridx from"
	sqlStr = sqlStr + " [db_academy].dbo.tbl_lecturer_coupon_master m,"
	sqlStr = sqlStr + " [db_academy].dbo.tbl_lecturer_coupon_detail d"
	sqlStr = sqlStr + " where m.lecturercouponidx=d.lecturercouponidx"
	sqlStr = sqlStr + " and m.lecturercouponidx<>" + CStr(lecturercouponidx)
	sqlStr = sqlStr + " and m.openstate<9"  ''�߱������ΰ� ����
	sqlStr = sqlStr + " and ( "
	sqlStr = sqlStr + " 	(m.lecturercouponstartdate<='" + CStr(lecturercouponstartdate) + "' and m.lecturercouponexpiredate>'" + CStr(lecturercouponstartdate) + "')"
	sqlStr = sqlStr + " 	or "
	sqlStr = sqlStr + " 	(m.lecturercouponstartdate<='" + CStr(lecturercouponexpiredate) + "' and m.lecturercouponexpiredate>'" + CStr(lecturercouponexpiredate) + "')"
	sqlStr = sqlStr + " 	)"
	sqlStr = sqlStr + " and d.lectureridx in (" + addSql + ")"  + VbCrlf

	'response.write sqlStr &"<Br>"
	rsacademyget.Open sqlStr,dbacademyget
	
	if not rsacademyget.Eof then
		do until rsacademyget.Eof
			ErrStr = ErrStr + "������ȣ : " + CStr(rsacademyget("lecturercouponidx")) + " - ���¹�ȣ : " + CStr(rsacademyget("lectureridx")) + " ����� \n"
			rsacademyget.moveNext
		loop
	end if
	
	rsacademyget.close

	'' ���� ��Ͽ� ���� ��ǰ�� �߰�.
	sqlStr = "insert into [db_academy].dbo.tbl_lecturer_coupon_detail" + VbCrlf
	sqlStr = sqlStr + " (lecturercouponidx, lectureridx, couponbuyprice)" + VbCrlf
	sqlStr = sqlStr + " select " + CStr(lecturercouponidx) + "," + VbCrlf
	sqlStr = sqlStr + " i.idx, " + VbCrlf
	
	Select Case margintype
		Case "00"  	''��ǰ�������� - ���԰� 0 �ΰ�� �����԰�
			sqlStr = sqlStr + " 0 " + VbCrlf
		'Case "10"	''�ΰŽ��δ� - ���԰� ����x
		'	if lecturercoupontype="1" then			''������
		'		sqlStr = sqlStr + " i.buycash - convert(int,i.sellcash*" + CStr(lecturercouponvalue/100) + ")"
		'	elseif lecturercoupontype="2" then   	''�ݾ�
		'		sqlStr = sqlStr + " i.buycash - " + CStr(lecturercouponvalue)
		'	else
		'		sqlStr = sqlStr + " 0 " + VbCrlf
		'	end if
		Case "10"	''�ΰŽ��δ� - �����԰�
			sqlStr = sqlStr + " 0 " + VbCrlf

		Case "20"	''�������� : �߰� [2008-09-23]
			if lecturercoupontype="1" then			''������
				sqlStr = sqlStr & " convert(int,i.lec_cost*"& Cstr((100-lecturercouponvalue)/100) &"*"& Cstr((100-defaultmargin)/100) &")"
				'response.Write "<javascript language=javascript>alert(' convert(int,i.sellcash*"& Cstr((100-lecturercouponvalue)/100) &"*"& Cstr((100-defaultmargin)/100) &") ')</script>"
				'response.end
			elseif lecturercoupontype="2" then   	''�ݾ�
				sqlStr = sqlStr + " convert(int,(i.lec_cost-" & CStr(lecturercouponvalue) + ")*"& Cstr((100-defaultmargin)/100) &")"
			else
				sqlStr = sqlStr + " 0 " + VbCrlf
			end if
		Case "30"	''���ϸ��� - ���縶�� : �߰� [2008-09-23]
			if lecturercoupontype="1" then			''������
				sqlStr = sqlStr + " convert(int,i.lec_cost*" + CStr((100-lecturercouponvalue)/100) + "*i.buying_cost/i.lec_cost)"
			elseif lecturercoupontype="2" then   	''�ݾ�
				sqlStr = sqlStr + " convert(int,(i.lec_cost-" + CStr(lecturercouponvalue) + ")*i.buying_cost/i.lec_cost)"
			else
				sqlStr = sqlStr + " 0 " + VbCrlf
			end if
		Case "50"	''�ݹݺδ�
			if lecturercoupontype="1" then			''������
				sqlStr = sqlStr + " i.buying_cost - convert(int,i.lec_cost*" + CStr(lecturercouponvalue/100) + "*0.5)"
			elseif lecturercoupontype="2" then   	''�ݾ�
				sqlStr = sqlStr + " i.buying_cost - convert(int," + CStr(lecturercouponvalue) + "*0.5)"
			else
				sqlStr = sqlStr + " 0 " + VbCrlf
			end if
		Case "60"	''��ü�δ� - ���԰� ����
			if lecturercoupontype="1" then			''������
				sqlStr = sqlStr + " i.buying_cost - convert(int,i.lec_cost*" + CStr(lecturercouponvalue/100) + ")"
			elseif lecturercoupontype="2" then   	''�ݾ�
				sqlStr = sqlStr + " i.buying_cost - " + CStr(lecturercouponvalue)
			else
				sqlStr = sqlStr + " 0 " + VbCrlf
			end if

		Case else
			sqlStr = sqlStr + " 0 " + VbCrlf
	end Select

	sqlStr = sqlStr + " from [db_academy].dbo.tbl_lec_item i" + VbCrlf
	sqlStr = sqlStr + " left join [db_academy].dbo.tbl_lecturer_coupon_detail d" + VbCrlf
	sqlStr = sqlStr + " 	on d.lecturercouponidx=" + CStr(lecturercouponidx) + VbCrlf
	sqlStr = sqlStr + " 	and d.lectureridx=i.idx" + VbCrlf
	sqlStr = sqlStr + " where i.idx in (" + addSql + ")"  + VbCrlf
	sqlStr = sqlStr + " and d.lectureridx is null"
	sqlStr = sqlStr + " and i.idx not in ("
	sqlStr = sqlStr + " 	select distinct d.lectureridx from"
	sqlStr = sqlStr + " 	[db_academy].dbo.tbl_lecturer_coupon_master m,"
	sqlStr = sqlStr + " 	[db_academy].dbo.tbl_lecturer_coupon_detail d"
	sqlStr = sqlStr + " 	where m.lecturercouponidx=d.lecturercouponidx"
	sqlStr = sqlStr + " 	and m.lecturercouponidx<>" + CStr(lecturercouponidx)
	sqlStr = sqlStr + " 	and m.openstate<9"  ''�߱������ΰ� ����
	sqlStr = sqlStr + " 	and ( "
	sqlStr = sqlStr + " 		(m.lecturercouponstartdate<='" + CStr(lecturercouponstartdate) + "' and m.lecturercouponexpiredate>'" + CStr(lecturercouponstartdate) + "')"
	sqlStr = sqlStr + " 		or "
	sqlStr = sqlStr + " 		(m.lecturercouponstartdate<='" + CStr(lecturercouponexpiredate) + "' and m.lecturercouponexpiredate>'" + CStr(lecturercouponexpiredate) + "')"
	sqlStr = sqlStr + " 		)"
	sqlStr = sqlStr + " 	and d.lectureridx in (" + addSql + ")"  + VbCrlf
	sqlStr = sqlStr + " ) "

	'response.write sqlStr &"<Br>"
	rsacademyget.Open sqlStr,dbacademyget,1

	''�����ǰ��.
	AplyItemCountUpdate lecturercouponidx
	AplyToItem lecturercouponidx
	
elseif mode="delcouponitemarr" then
	lectureridxarr = trim(lectureridxarr)
	if Right(lectureridxarr,1)="," then lectureridxarr=Left(lectureridxarr,Len(lectureridxarr)-1)

	sqlStr = "delete from [db_academy].dbo.tbl_lecturer_coupon_detail" + VbCrlf
	sqlStr = sqlStr + " where lecturercouponidx=" + CStr(lecturercouponidx) + VbCrlf
	sqlStr = sqlStr + " and lectureridx in (" + lectureridxarr + ")"  + VbCrlf
	
	'response.write sqlStr &"<Br>"
	rsacademyget.Open sqlStr,dbacademyget,1

	''�����ǰ��.
	AplyItemCountUpdate lecturercouponidx

	''������ ���� ��ǰ���̺��� ���� ���� N �� ����
	AplyToItem lecturercouponidx

'//������ ��ϵ� ���� ���� ����� ���԰� ����	
elseif mode="modicouponitemarr" then
	lectureridxarr = trim(lectureridxarr)
	couponbuypricearr  = trim(couponbuypricearr)

	if Right(lectureridxarr,1)="," then lectureridxarr=Left(lectureridxarr,Len(lectureridxarr)-1)
	if Right(couponbuypricearr,1)="," then couponbuypricearr=Left(couponbuypricearr,Len(couponbuypricearr)-1)

	lectureridxarr = split(lectureridxarr,",")
	couponbuypricearr = split(couponbuypricearr,",")

	for i=LBound(lectureridxarr) to UBound(lectureridxarr)
		if trim(lectureridxarr(i))<>"" then
			sqlStr = "update [db_academy].dbo.tbl_lecturer_coupon_detail" + VbCrlf
			sqlStr = sqlStr + " set couponbuyprice=" + CStr(couponbuypricearr(i)) + VbCrlf
			sqlStr = sqlStr + " where lecturercouponidx=" + CStr(lecturercouponidx) + VbCrlf
			sqlStr = sqlStr + " and lectureridx=" + CStr(lectureridxarr(i)) + VbCrlf
			
			'response.write sqlStr &"<Br>"
			rsacademyget.Open sqlStr,dbacademyget,1
		end if
	next

	''�����ǰ��.
	AplyItemCountUpdate lecturercouponidx
	AplyToItem lecturercouponidx

'//�������� ����	
elseif mode="opencoupon" Then

	sqlStr = "update [db_academy].dbo.tbl_lecturer_coupon_master" + VbCrlf
	sqlStr = sqlStr + " set openstate='7'"
	sqlStr = sqlStr + " where lecturercouponidx=" + CStr(lecturercouponidx) + VbCrlf

	'response.write sqlStr &"<Br>"
	rsacademyget.Open sqlStr,dbacademyget,1

	AplyToItem(lecturercouponidx)

elseif mode="reservecoupon" Then

	sqlStr = "update [db_academy].dbo.tbl_lecturer_coupon_master" + VbCrlf
	sqlStr = sqlStr + " set openstate='6'"
	sqlStr = sqlStr + " where lecturercouponidx=" + CStr(lecturercouponidx) + VbCrlf
	
	'response.write sqlStr &"<Br>"
	rsacademyget.Open sqlStr,dbacademyget,1

'//�������� �߱� ���� ����
elseif mode="closecoupon" Then

    dim MayExpireDt
    MayExpireDt = Left(CStr(DateAdd("d",-1,Now())),10) & " 23:59:59"

    ''response.write MayExpireDt

    ''������ �߱� �� ���� Expire
    sqlStr = "update [db_academy].dbo.tbl_user_lecturer_coupon" + VbCrlf
    sqlStr = sqlStr + " set lecturercouponexpiredate='" & MayExpireDt & "'" + VbCrlf
    sqlStr = sqlStr + " where lecturercouponidx=" + CStr(lecturercouponidx) + VbCrlf
    sqlStr = sqlStr + " and lecturercouponexpiredate>'" & MayExpireDt & "'" + VbCrlf
    sqlStr = sqlStr + " and usedyn='N'" + VbCrlf
	
	'response.write sqlStr &"<Br>"
    dbacademyget.Execute sqlStr

	sqlStr = "update [db_academy].dbo.tbl_lecturer_coupon_master" + VbCrlf
	sqlStr = sqlStr + " set openstate='9'"
	sqlStr = sqlStr + " where lecturercouponidx=" + CStr(lecturercouponidx) + VbCrlf
	
	'response.write sqlStr &"<Br>"
	dbacademyget.Execute sqlStr

	AplyToItem(lecturercouponidx)
end if
%>
<% if (mode="couponmaster") then %>
	<% if (IsEditMode) then %>
		<script language='javascript'>
			alert('���� �Ǿ����ϴ�.');
			location.replace('/academy/lecture/coupon/lecturercouponmasterreg.asp?lecturercouponidx=<%= lecturercouponidx %>');
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
		//location.replace('/academy/lecture/coupon/lecturercouponitemlistedit.asp?lecturercouponidx=<%= lecturercouponidx %>&makerid=<%= makerid %>&sailyn=<%= sailyn %>');
	</script>
	
<% elseif mode="delcouponitemarr" then %>
	<script language='javascript'>
		alert('���� �Ǿ����ϴ�.');
		opener.location.reload();
		location.replace('/academy/lecture/coupon/lecturercouponitemlistedit.asp?lecturercouponidx=<%= lecturercouponidx %>&makerid=<%= makerid %>&sailyn=<%= sailyn %>');
	</script>
	
<% elseif mode="modicouponitemarr" then %>
	<script language='javascript'>
		alert('���� �Ǿ����ϴ�.');
		opener.location.reload();
		location.replace('/academy/lecture/coupon/lecturercouponitemlistedit.asp?lecturercouponidx=<%= lecturercouponidx %>&makerid=<%= makerid %>&sailyn=<%= sailyn %>');
	</script>
	
<% elseif mode="opencoupon" then %>
	<script language='javascript'>
		alert('������ ���� �Ǿ����ϴ�.');
		opener.location.reload();
		location.replace('/academy/lecture/coupon/lecturercouponmasterreg.asp?lecturercouponidx=<%= lecturercouponidx %>');
	</script>
	
<% elseif mode="reservecoupon" then %>
	<script language='javascript'>
		alert('������ ������ ���� �Ǿ����ϴ�. ���� 0�ÿ� ����˴ϴ�.');
		opener.location.reload();
		location.replace('/academy/lecture/coupon/lecturercouponmasterreg.asp?lecturercouponidx=<%= lecturercouponidx %>');
	</script>
	
<% elseif mode="closecoupon" then %>
	<script language='javascript'>
		alert('������ ���� �Ǿ����ϴ�.');
		opener.location.reload();
		location.replace('/academy/lecture/coupon/lecturercouponmasterreg.asp?lecturercouponidx=<%= lecturercouponidx %>');
		self.close();
	</script>
<% end if %>

<%= "mode=" + mode %>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->