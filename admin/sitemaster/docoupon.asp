<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���ʽ� ����
' History : ������ ����
'			2022.07.04 �ѿ�� ����(isms���������)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/newcouponcls.asp" -->
<%
dim idx,couponname,couponvalue,coupontype ,sqlStr
dim openfinishdate,isusing, etcstr ,isopenlistcoupon ,minbuyprice,startdate,expiredate
dim targetitemlist,couponmeaipprice , isfreebeasongcoupon,isweekendcoupon ,validsitename
dim targetcpntype, targetcpnsourcebrand, targetcpnsourcecate, targetcpnsource, mxCpnDiscount, brandShareValue
dim couponimage
	idx			= requestCheckvar(request("idx"),10)
	couponname	= html2db(request("couponname"))
	couponvalue = requestCheckvar(request("couponvalue"),10)
	coupontype	= requestCheckvar(request("coupontype"),10)
	minbuyprice = requestCheckvar(request("minbuyprice"),10)
	startdate	= requestCheckvar(request("startdate"),19)
	expiredate	= requestCheckvar(request("expiredate"),19)
	openfinishdate	= requestCheckvar(request("openfinishdate"),19)
	isusing			= requestCheckvar(request("isusing"),10)
	etcstr	= html2db(request("etcstr"))
	isopenlistcoupon = requestCheckvar(request("isopenlistcoupon"),10)
	validsitename = request("validsitename")
	mxCpnDiscount = requestCheckvar(request("mxCpnDiscount"),10)
	brandShareValue = getNumeric(requestCheckvar(request("brandShareValue"),5))
	couponimage     = request("usercouponimage")

'targetitemlist = request("targetitemlist")
'couponmeaipprice = request("couponmeaipprice")
if (Not IsNumeric(couponmeaipprice)) or (couponmeaipprice="") then couponmeaipprice=0
if coupontype<>"1" then mxCpnDiscount=0

''�ָ���������, ���������� ���� �߰�
isfreebeasongcoupon = request("isfreebeasongcoupon")
isweekendcoupon = request("isweekendcoupon")


    targetcpntype = requestCheckvar(request("targetcpntype"),10)  '' "","B","C"   �Ϲ�,�귣��, ī�װ�
    targetcpnsourcebrand = requestCheckvar(request("targetcpnsourcebrand"),32)
    targetcpnsourcecate = requestCheckvar(request("targetcpnsourcecate"),32)
    
    if (targetcpntype="B") then
        targetcpnsource = Trim(targetcpnsourcebrand)
    end if
    
    if (targetcpntype="C") then
        targetcpnsource = Trim(targetcpnsourcecate)
    end if
    
    
    if (targetcpntype="B") and Len(targetcpnsourcebrand)<1 then
        response.write "�귣��ID ����"
        dbget.close() : response.end
    end if
    
    if (targetcpntype="C") and Len(targetcpnsourcecate)<1 then
        response.write "ī�װ� �ڵ� ����"
        dbget.close() : response.end
    end if
    
    if ((targetcpntype="B") or (targetcpntype="C")) and (isfreebeasongcoupon<>"") then
        response.write "�귣��,ī�װ� ������ ���������� ���Ұ�."
        dbget.close() : response.end
    end if
    
    if (targetcpntype="B") then
        '' check valid brandid
        if Not(checkValidBrandID(targetcpnsource)) then
            response.write "�ùٸ� �귣��ID�� �ƴմϴ�. - "&targetcpnsource
            dbget.close() : response.end
        end if

		if brandShareValue>50 then
			response.write "�귣�������� ��ü �д����� 50%�� ������ �����ϴ�."
			dbget.close() : response.end
		end if
    end if
    
    if (targetcpntype="C") then
        '' check valid categoryid
        if Not(checkValidDispCategoryID(targetcpnsource)) then
            response.write "�ùٸ� ī�װ��ڵ尡 �ƴմϴ�. - "&targetcpnsource
            dbget.close() : response.end
        end if
    end if
    
    ''ī�װ� depth
    ' if (targetcpntype="C") and Len(targetcpnsourcecate)<6 then
    '     response.write "ī�װ� ������ 2depth �̻� ���� �ϼ���."
    '     dbget.close() : response.end
    ' end if
	if (targetcpntype="C") and Len(targetcpnsourcecate)<3 then
        response.write "ī�װ� ������ 1depth �̻� ���� �ϼ���."
        dbget.close() : response.end
    end if
    if targetcpntype<>"B" or brandShareValue="" then brandShareValue=0

if isweekendcoupon<>"Y" then isweekendcoupon="N"
if isfreebeasongcoupon="Y" then 
    coupontype ="3"
    couponvalue = Cstr(getDefaultBeasongPayByDate(now()))
'   minbuyprice ="0"
    targetitemlist="0"
end if

if (idx<>"") then
	if couponname <> "" and not(isnull(couponname)) then
		couponname = ReplaceBracket(couponname)

		if checkNotValidHTML(couponname) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('�������� HTML�� ����Ͻ� �� �����ϴ�.');history.back();"
			response.write "</script>"
			response.End
		end if
	end If

	if etcstr <> "" and not(isnull(etcstr)) then
		etcstr = ReplaceBracket(etcstr)

		if checkNotValidHTML(etcstr) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('�ڸ�Ʈ���� HTML�� ����Ͻ� �� �����ϴ�.');history.back();"
			response.write "</script>"
			response.End
		end if
	end If

	sqlStr = "update [db_user].[dbo].tbl_user_coupon_master" + VBCrlf
	sqlStr = sqlStr + " set couponname='" + couponname + "'" + VBCrlf
	sqlStr = sqlStr + " ,couponvalue=" + couponvalue + "" + VBCrlf
	sqlStr = sqlStr + " ,coupontype='" + coupontype + "'" + VBCrlf
	sqlStr = sqlStr + " ,minbuyprice=" + minbuyprice + "" + VBCrlf
	sqlStr = sqlStr + " ,startdate='" + startdate + "'" + VBCrlf
	sqlStr = sqlStr + " ,expiredate='" + expiredate + "'" + VBCrlf
	sqlStr = sqlStr + " ,openfinishdate='" + openfinishdate + "'" + VBCrlf
	sqlStr = sqlStr + " ,isusing='" + isusing + "'" + VBCrlf
	sqlStr = sqlStr + " ,etcstr='" + etcstr + "'" + VBCrlf
	sqlStr = sqlStr + " ,isopenlistcoupon='" + isopenlistcoupon + "'" + VBCrlf
	sqlStr = sqlStr + " ,isweekendcoupon='" + isweekendcoupon + "'" + VBCrlf
	sqlStr = sqlStr + " ,targetitemlist='" + targetitemlist + "'" + VBCrlf
	sqlStr = sqlStr + " ,couponmeaipprice=" + CStr(couponmeaipprice) + "" + VBCrlf
	IF (validsitename<>"") then
	    sqlStr = sqlStr + " ,validsitename='" + CStr(validsitename) + "'" + VBCrlf
	ELSE
	    sqlStr = sqlStr + " ,validsitename=NULL" + VBCrlf
    END IF
    
    IF (targetcpntype<>"") then
	    sqlStr = sqlStr + " ,targetcpntype='" + CStr(targetcpntype) + "'" + VBCrlf
	ELSE
	    sqlStr = sqlStr + " ,targetcpntype=NULL" + VBCrlf
    END IF
    
    IF (targetcpnsource<>"") then
	    sqlStr = sqlStr + " ,targetcpnsource='" + CStr(targetcpnsource) + "'" + VBCrlf
	ELSE
	    sqlStr = sqlStr + " ,targetcpnsource=NULL" + VBCrlf
    END IF
	sqlStr = sqlStr + " ,mxCpnDiscount="&CStr(mxCpnDiscount)&VBCrlf
	sqlStr = sqlStr + " ,brandShareValue="&CStr(brandShareValue)&VBCrlf
	sqlStr = sqlStr + ", couponimage='" + couponimage + "'" + VbCrlf
	
	sqlStr = sqlStr + " where idx=" + CStr(idx)

	dbget.Execute sqlStr
else
	if couponname <> "" and not(isnull(couponname)) then
		couponname = ReplaceBracket(couponname)

		if checkNotValidHTML(couponname) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('�������� HTML�� ����Ͻ� �� �����ϴ�.');history.back();"
			response.write "</script>"
			response.End
		end if
	end If

	if etcstr <> "" and not(isnull(etcstr)) then
		etcstr = ReplaceBracket(etcstr)

		if checkNotValidHTML(etcstr) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('�ڸ�Ʈ���� HTML�� ����Ͻ� �� �����ϴ�.');history.back();"
			response.write "</script>"
			response.End
		end if
	end If

	sqlStr = "insert into [db_user].[dbo].tbl_user_coupon_master" + VBCrlf
	sqlStr = sqlStr + " (couponname,couponvalue,coupontype,minbuyprice" + VBCrlf
	sqlStr = sqlStr + " ,startdate,expiredate,openfinishdate,isusing,etcstr" + VBCrlf
	sqlStr = sqlStr + " ,isopenlistcoupon,isweekendcoupon,targetitemlist,couponmeaipprice, validsitename,targetcpntype,targetcpnsource,mxCpnDiscount,brandShareValue)" + VBCrlf
	sqlStr = sqlStr + " values("
	sqlStr = sqlStr + " '" + couponname + "'" + VBCrlf
	sqlStr = sqlStr + " ," + couponvalue + "" + VBCrlf
	sqlStr = sqlStr + " ,'" + coupontype + "'" + VBCrlf
	sqlStr = sqlStr + " ," + minbuyprice + "" + VBCrlf
	sqlStr = sqlStr + " ,'" + startdate + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + expiredate + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + openfinishdate + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + isusing + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + etcstr + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + isopenlistcoupon + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + isweekendcoupon + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + targetitemlist + "'" + VBCrlf
	sqlStr = sqlStr + " ," + Cstr(couponmeaipprice) + "" + VBCrlf
	IF (validsitename<>"") then
	    sqlStr = sqlStr + " ,'" + CStr(validsitename) + "'" + VBCrlf
	ELSE
	    sqlStr = sqlStr + " ,NULL" + VBCrlf
    END IF
    
    IF (targetcpntype<>"") then
        sqlStr = sqlStr + " ,'" + CStr(targetcpntype) + "'" + VBCrlf
	ELSE
	    sqlStr = sqlStr + " ,NULL" + VBCrlf
    END IF
    
    IF (targetcpnsource<>"") then
        sqlStr = sqlStr + " ,'" + CStr(targetcpnsource) + "'" + VBCrlf
	ELSE
	    sqlStr = sqlStr + " ,NULL" + VBCrlf
    END IF
	sqlStr = sqlStr + " ," + Cstr(mxCpnDiscount) + "" + VBCrlf
	sqlStr = sqlStr + " ," + Cstr(brandShareValue) + "" + VBCrlf
	
	sqlStr = sqlStr + " )"

	'response.write sqlStr
	dbget.Execute sqlStr
end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>

<script language="javascript">
	alert('�����Ǿ����ϴ�.');
	location.replace('/admin/sitemaster/couponlist.asp?menupos=<%=menupos%>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->