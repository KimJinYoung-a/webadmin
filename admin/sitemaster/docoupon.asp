<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  보너스 쿠폰
' History : 서동석 생성
'			2022.07.04 한용민 수정(isms취약점수정)
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

''주말쿠폰구분, 무료배송쿠폰 구분 추가
isfreebeasongcoupon = request("isfreebeasongcoupon")
isweekendcoupon = request("isweekendcoupon")


    targetcpntype = requestCheckvar(request("targetcpntype"),10)  '' "","B","C"   일반,브랜드, 카테고리
    targetcpnsourcebrand = requestCheckvar(request("targetcpnsourcebrand"),32)
    targetcpnsourcecate = requestCheckvar(request("targetcpnsourcecate"),32)
    
    if (targetcpntype="B") then
        targetcpnsource = Trim(targetcpnsourcebrand)
    end if
    
    if (targetcpntype="C") then
        targetcpnsource = Trim(targetcpnsourcecate)
    end if
    
    
    if (targetcpntype="B") and Len(targetcpnsourcebrand)<1 then
        response.write "브랜드ID 오류"
        dbget.close() : response.end
    end if
    
    if (targetcpntype="C") and Len(targetcpnsourcecate)<1 then
        response.write "카테고리 코드 오류"
        dbget.close() : response.end
    end if
    
    if ((targetcpntype="B") or (targetcpntype="C")) and (isfreebeasongcoupon<>"") then
        response.write "브랜드,카테고리 쿠폰은 무료배송쿠폰 사용불가."
        dbget.close() : response.end
    end if
    
    if (targetcpntype="B") then
        '' check valid brandid
        if Not(checkValidBrandID(targetcpnsource)) then
            response.write "올바른 브랜드ID가 아닙니다. - "&targetcpnsource
            dbget.close() : response.end
        end if

		if brandShareValue>50 then
			response.write "브랜드쿠폰의 업체 분담율은 50%를 넘을수 없습니다."
			dbget.close() : response.end
		end if
    end if
    
    if (targetcpntype="C") then
        '' check valid categoryid
        if Not(checkValidDispCategoryID(targetcpnsource)) then
            response.write "올바른 카테고리코드가 아닙니다. - "&targetcpnsource
            dbget.close() : response.end
        end if
    end if
    
    ''카테고리 depth
    ' if (targetcpntype="C") and Len(targetcpnsourcecate)<6 then
    '     response.write "카테고리 쿠폰은 2depth 이상 선택 하세요."
    '     dbget.close() : response.end
    ' end if
	if (targetcpntype="C") and Len(targetcpnsourcecate)<3 then
        response.write "카테고리 쿠폰은 1depth 이상 선택 하세요."
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
			response.write "	alert('쿠폰명에는 HTML을 사용하실 수 없습니다.');history.back();"
			response.write "</script>"
			response.End
		end if
	end If

	if etcstr <> "" and not(isnull(etcstr)) then
		etcstr = ReplaceBracket(etcstr)

		if checkNotValidHTML(etcstr) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('코멘트에는 HTML을 사용하실 수 없습니다.');history.back();"
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
			response.write "	alert('쿠폰명에는 HTML을 사용하실 수 없습니다.');history.back();"
			response.write "</script>"
			response.End
		end if
	end If

	if etcstr <> "" and not(isnull(etcstr)) then
		etcstr = ReplaceBracket(etcstr)

		if checkNotValidHTML(etcstr) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('코멘트에는 HTML을 사용하실 수 없습니다.');history.back();"
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
	alert('수정되었습니다.');
	location.replace('/admin/sitemaster/couponlist.asp?menupos=<%=menupos%>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->