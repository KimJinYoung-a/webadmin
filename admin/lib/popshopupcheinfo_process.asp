<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 매장 계약관리
' Hieditor : 2009.04.07 서동석 생성
'			 2011.01.21 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim shopid, designer, mode, cksel, defaultCenterMwdiv ,sqlStr, i, cnt, pp
dim adminopen, chargediv, comm_cd, defaultmargin, subtract,defaultsuplymargin, defaultbeasongdiv
dim etcjunsandetail, autojungsan, autojungsandiv ,itemregyn
	shopid      = RequestCheckVar(request.form("shopid"),1024)
	designer    = RequestCheckVar(request("designer"),32)
	mode        = RequestCheckVar(request("mode"),32)
	cksel       = RequestCheckVar(request("cksel"),1024)
	defaultCenterMwdiv = RequestCheckVar(request("defaultCenterMwdiv"),1024)
	comm_cd             = RequestCheckVar(request("comm_cd"),1024)
	defaultmargin       = RequestCheckVar(request("defaultmargin"),1024)
	defaultsuplymargin  = RequestCheckVar(request("defaultsuplymargin"),1024)
	defaultbeasongdiv  = RequestCheckVar(request("defaultbeasongdiv"),1024)
	etcjunsandetail     = RequestCheckVar(html2db(request("etcjunsandetail")),2048)
	itemregyn = RequestCheckVar(request("itemregyn"),1)

dim refer
	refer       = request.ServerVariables("HTTP_REFERER")

function ChkUpdateShopMagin(shopid, designer, comm_cd, defaultmargin, defaultsuplymargin, defaultbeasongdiv, etcjunsandetail)
    dim sqlStr
    dim protoShopID, resultCount
    dim old_comm_cd, old_defaultmargin, old_defaultsuplymargin, old_defaultbeasongdiv
    dim chargediv  ''예전 정산구분
    dim adminopen

	if (defaultbeasongdiv = "") or (Not IsNumeric(defaultbeasongdiv)) then
		defaultbeasongdiv = "0"
	end if

    resultCount = 0

    if (Left(shopid,11)="streetshop0") then
        protoShopID = "streetshop000"
    elseif (Left(shopid,12)="streetshop87") then
        protoShopID = "streetshop870"
    elseif (Left(shopid,11)="streetshop8") then
        protoShopID = "streetshop800"
    end if

    sqlStr = "select top 1 * from [db_shop].[dbo].tbl_shop_designer" + VbCrlf
    sqlStr = sqlStr + " where shopid='" + shopid + "'" + VbCrlf
    sqlStr = sqlStr + " and makerid='" + designer + "'" + VbCrlf

    'response.write sqlStr & "<Br>"
    rsget.Open sqlStr,dbget,1
        resultCount = rsget.RecordCount
        if Not rsget.Eof then
            old_comm_cd             = rsget("comm_cd")
            old_defaultmargin       = rsget("defaultmargin")
            old_defaultsuplymargin  = rsget("defaultsuplymargin")
            old_defaultbeasongdiv   = rsget("defaultbeasongdiv")

            if (IsNull(old_defaultbeasongdiv) = True) then
            	old_defaultbeasongdiv = "0"
            end if
        end if
    rsget.Close

    if (chargediv="") then
        if comm_cd="B011" then chargediv="2"                        ''텐위
        if comm_cd="B031" then chargediv="5"                        ''출고(매입)/오프매입
        if comm_cd="B021" then chargediv="5"                        ''출고(매입)/오프매입
        if comm_cd="B012" then chargediv="6"                        ''업위
        if comm_cd="B022" then chargediv="8"                        ''업매

        if comm_cd="B013" then chargediv="2"                        ''출고특정 =>2
    end if

    ''업체 특정/업체 매입/텐위 인경우 어드민 오픈.  ''차후에 다 오픈..(사용안함)
    if (comm_cd="B012") or (comm_cd="B022") or (comm_cd="B011") then
        adminopen = "Y"
    else
        adminopen = "N"
    end if

    if (resultCount>0) then
		sqlStr = "update [db_shop].[dbo].tbl_shop_designer" + VbCrlf
		sqlStr = sqlStr + " set chargediv='" + CStr(chargediv) + "'" + VbCrlf
		sqlStr = sqlStr + " ,comm_cd='" + CStr(comm_cd) + "'" + VbCrlf
		sqlStr = sqlStr + " ,defaultmargin=" + CStr(defaultmargin) + ""  + VbCrlf
		sqlStr = sqlStr + " ,defaultsuplymargin=" + CStr(defaultsuplymargin) + ""  + VbCrlf
		sqlStr = sqlStr + " ,defaultbeasongdiv=" + CStr(defaultbeasongdiv) + ""  + VbCrlf
		sqlStr = sqlStr + " ,subtract=0" + VbCrlf
		sqlStr = sqlStr + " ,adminopen='" + CStr(adminopen) + "'" + VbCrlf
		sqlStr = sqlStr + " ,autojungsan='Y'" + VbCrlf
		sqlStr = sqlStr + " ,autojungsandiv='S'" + VbCrlf
		sqlStr = sqlStr + " ,sdLastUpdate=getdate()" + VbCrlf
		''기타메모는 대표샵으로 통일
		''sqlStr = sqlStr + " etcjunsandetail='" + CStr(etcjunsandetail) + "'" + VbCrlf
		sqlStr = sqlStr + " where shopid='" + shopid + "'" + VbCrlf
		sqlStr = sqlStr + " and makerid='" + designer + "'" + VbCrlf

		'response.write sqlStr & "<Br>"
		dbget.Execute sqlStr

		sqlStr = "update [db_shop].[dbo].tbl_shop_designer" + VbCrlf
		sqlStr = sqlStr + " set etcjunsandetail='" + CStr(etcjunsandetail) + "'" + VbCrlf
		sqlStr = sqlStr + " where shopid='" + protoShopID + "'" + VbCrlf
		sqlStr = sqlStr + " and makerid='" + designer + "'" + VbCrlf

		'response.write sqlStr & "<Br>"
		dbget.Execute sqlStr

		''로그 입력
		if ((old_comm_cd<>comm_cd) or (CStr(old_defaultmargin)<>defaultmargin) or (CStr(old_defaultsuplymargin)<>defaultsuplymargin) or (CStr(old_defaultbeasongdiv)<>defaultbeasongdiv)) then
		    sqlStr = "insert into [db_shop].[dbo].tbl_shop_designer_Maginlog" + VbCrlf
		    sqlStr = sqlStr + " (shopid,makerid,comm_cd,defaultmargin,defaultsuplymargin,defaultbeasongdiv,actFlag,reguserid)" + VbCrlf
		    sqlStr = sqlStr + " values(" + VbCrlf
		    sqlStr = sqlStr + " '" + shopid + "'" + VbCrlf
		    sqlStr = sqlStr + " ,'" + designer + "'" + VbCrlf
            sqlStr = sqlStr + " ,'" + comm_cd + "'" + VbCrlf
            sqlStr = sqlStr + " ," + defaultmargin + "" + VbCrlf
            sqlStr = sqlStr + " ," + defaultsuplymargin + "" + VbCrlf
            sqlStr = sqlStr + " ," + defaultbeasongdiv + "" + VbCrlf
            sqlStr = sqlStr + " ,'M'" + VbCrlf
            sqlStr = sqlStr + " ,'" + session("ssBctID") + "')" + VbCrlf

			'response.write sqlStr & "<Br>"
            dbget.Execute sqlStr
		end if
	else
		sqlStr = "insert into [db_shop].[dbo].tbl_shop_designer" + VbCrlf
		sqlStr = sqlStr + " (shopid,makerid,chargediv,comm_cd,defaultmargin,defaultsuplymargin,defaultbeasongdiv,subtract,adminopen,autojungsan,autojungsandiv,etcjunsandetail)" + VbCrlf
		sqlStr = sqlStr + " values('" + shopid + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + designer + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + chargediv + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + comm_cd + "'" + VbCrlf
		sqlStr = sqlStr + " ," + defaultmargin + "" + VbCrlf
		sqlStr = sqlStr + " ," + defaultsuplymargin + "" + VbCrlf
		sqlStr = sqlStr + " ," + defaultbeasongdiv + "" + VbCrlf
		sqlStr = sqlStr + " ,0" + VbCrlf
		sqlStr = sqlStr + " ,'" + adminopen + "'" + VbCrlf
		sqlStr = sqlStr + " ,'Y'" + VbCrlf
		sqlStr = sqlStr + " ,'S'" + VbCrlf
		sqlStr = sqlStr + " ,'')" + VbCrlf

		'response.write sqlStr & "<Br>"
		dbget.Execute sqlStr

		sqlStr = "update [db_shop].[dbo].tbl_shop_designer" + VbCrlf
		sqlStr = sqlStr + " set etcjunsandetail='" + CStr(etcjunsandetail) + "'" + VbCrlf
		sqlStr = sqlStr + " where shopid='" + protoShopID + "'" + VbCrlf
		sqlStr = sqlStr + " and makerid='" + designer + "'" + VbCrlf

		'response.write sqlStr & "<Br>"
		dbget.Execute sqlStr

		''로그 입력
	    sqlStr = "insert into [db_shop].[dbo].tbl_shop_designer_Maginlog" + VbCrlf
	    sqlStr = sqlStr + " (shopid,makerid,comm_cd,defaultmargin,defaultsuplymargin,defaultbeasongdiv,actFlag,reguserid)" + VbCrlf
	    sqlStr = sqlStr + " values(" + VbCrlf
	    sqlStr = sqlStr + " '" + shopid + "'" + VbCrlf
	    sqlStr = sqlStr + " ,'" + designer + "'" + VbCrlf
        sqlStr = sqlStr + " ,'" + comm_cd + "'" + VbCrlf
        sqlStr = sqlStr + " ," + defaultmargin + "" + VbCrlf
        sqlStr = sqlStr + " ," + defaultsuplymargin + "" + VbCrlf
        sqlStr = sqlStr + " ," + defaultbeasongdiv + "" + VbCrlf
        sqlStr = sqlStr + " ,'I'" + VbCrlf
        sqlStr = sqlStr + " ,'" + session("ssBctID") + "')" + VbCrlf

		'response.write sqlStr & "<Br>"
        dbget.Execute sqlStr
	end if
end function

function ChkCenterMwDivChange(shopid, designer, Defaultcentermwdiv)
    dim sqlStr
    dim old_comm_cd,old_defaultmargin, old_defaultsuplymargin, old_defaultCenterMwDiv

    sqlStr = "select top 1 comm_cd, defaultmargin, defaultsuplymargin, IsNULL(defaultCenterMwDiv,'N') as defaultCenterMwDiv from [db_shop].[dbo].tbl_shop_designer" + VbCrlf
    sqlStr = sqlStr + " where shopid='" + shopid + "'" + VbCrlf
    sqlStr = sqlStr + " and makerid='" + designer + "'" + VbCrlf

    'response.write sqlStr & "<Br>"
    rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            old_comm_cd             = rsget("comm_cd")
            old_defaultmargin       = rsget("defaultmargin")
            old_defaultsuplymargin  = rsget("defaultsuplymargin")
            old_defaultCenterMwDiv   = Trim(rsget("defaultCenterMwDiv"))
        end if
    rsget.Close

	if (Defaultcentermwdiv = "") then
		Defaultcentermwdiv = "N"
	end if
    if (Defaultcentermwdiv<>old_defaultCenterMwDiv) then

        ''로그입력
        sqlStr = "update [db_shop].[dbo].tbl_shop_designer" + VbCrlf
		if Trim(defaultCenterMwDiv) = "N" then
			sqlStr = sqlStr + " set defaultCenterMwDiv=NULL" + VbCrlf
		else
			sqlStr = sqlStr + " set defaultCenterMwDiv='" + defaultCenterMwDiv + "'" + VbCrlf
		end if
        sqlStr = sqlStr + " ,sdLastUpdate=getdate()" + VbCrlf
        sqlStr = sqlStr + " where shopid='" + shopid + "'" + VbCrlf
    	sqlStr = sqlStr + " and makerid='" + designer + "'" + VbCrlf

    	'response.write sqlStr & "<Br>"
    	dbget.Execute sqlStr

    	sqlStr = "insert into [db_shop].[dbo].tbl_shop_designer_Maginlog" + VbCrlf
        sqlStr = sqlStr + " (shopid,makerid,comm_cd,defaultmargin,defaultsuplymargin,actFlag, defaultCenterMwDiv,reguserid)" + VbCrlf
        sqlStr = sqlStr + " values(" + VbCrlf
        sqlStr = sqlStr + " '" + shopid + "'" + VbCrlf
        sqlStr = sqlStr + " ,'" + designer + "'" + VbCrlf
        sqlStr = sqlStr + " ,'" + old_comm_cd + "'" + VbCrlf
        sqlStr = sqlStr + " ," + CStr(old_defaultmargin) + "" + VbCrlf
        sqlStr = sqlStr + " ," + CStr(old_defaultsuplymargin) + "" + VbCrlf
        sqlStr = sqlStr + " ,'M'" + VbCrlf
        sqlStr = sqlStr + " ,'" + CStr(Defaultcentermwdiv) + "'" + VbCrlf
        sqlStr = sqlStr + " ,'" + session("ssBctID") + "')" + VbCrlf

        'response.write sqlStr & "<Br>"
        dbget.Execute sqlStr
    end if
end function

function ChkDelShopMagin(shopid, designer)
    dim resultCount, old_comm_cd, old_defaultmargin, old_defaultsuplymargin
    resultCount = 0
    dim sqlStr

    sqlStr = "select top 1 * from [db_shop].[dbo].tbl_shop_designer" + VbCrlf
    sqlStr = sqlStr + " where shopid='" + shopid + "'" + VbCrlf
    sqlStr = sqlStr + " and makerid='" + designer + "'" + VbCrlf

    'response.write sqlStr & "<Br>"
    rsget.Open sqlStr,dbget,1
        resultCount = rsget.RecordCount
        if Not rsget.Eof then
            old_comm_cd             = rsget("comm_cd")
            old_defaultmargin       = rsget("defaultmargin")
            old_defaultsuplymargin  = rsget("defaultsuplymargin")
        end if
    rsget.Close

    if (resultCount>0) then
        ''삭제.
    	sqlStr = "delete from [db_shop].[dbo].tbl_shop_designer" + VbCrlf
    	sqlStr = sqlStr + " where shopid='" + shopid + "'" + VbCrlf
    	sqlStr = sqlStr + " and makerid='" + designer + "'" + VbCrlf

    	'response.write sqlStr & "<Br>"
    	dbget.Execute sqlStr

        ''로그입력
    	sqlStr = "insert into [db_shop].[dbo].tbl_shop_designer_Maginlog" + VbCrlf
        sqlStr = sqlStr + " (shopid,makerid,comm_cd,defaultmargin,defaultsuplymargin,actFlag,reguserid)" + VbCrlf
        sqlStr = sqlStr + " values(" + VbCrlf
        sqlStr = sqlStr + " '" + shopid + "'" + VbCrlf
        sqlStr = sqlStr + " ,'" + designer + "'" + VbCrlf
        sqlStr = sqlStr + " ,'" + old_comm_cd + "'" + VbCrlf
        sqlStr = sqlStr + " ," + CStr(old_defaultmargin) + "" + VbCrlf
        sqlStr = sqlStr + " ," + CStr(old_defaultsuplymargin) + "" + VbCrlf
        sqlStr = sqlStr + " ,'D'" + VbCrlf
        sqlStr = sqlStr + " ,'" + session("ssBctID") + "')" + VbCrlf

        'response.write sqlStr & "<Br>"
        dbget.Execute sqlStr

    	ChkDelShopMagin = true
	else
	    ChkDelShopMagin = false
	end if
end function

if (mode="edit") then

    Call ChkUpdateShopMagin(shopid, designer, comm_cd, defaultmargin, defaultsuplymargin, defaultbeasongdiv, etcjunsandetail)

    response.write "<script>alert('수정 되었습니다.');</script>"
    response.write "<script>location.replace('" + refer + "');</script>"

    ''//그룹 전체 설정.
''	if (groupflag="on") then
''		''없는 샾 추가.
''		sqlStr = "insert into [db_shop].[dbo].tbl_shop_designer" + VbCrlf
''		sqlStr = sqlStr + " (shopid,makerid,chargediv,comm_cd,defaultmargin,defaultsuplymargin,subtract,adminopen,autojungsan,autojungsandiv,etcjunsandetail)" + VbCrlf
''		sqlStr = sqlStr + " select s.userid" + VbCrlf
''		sqlStr = sqlStr + " ,'" + designer + "'" + VbCrlf
''		sqlStr = sqlStr + " ,'" + chargediv + "'" + VbCrlf
''		sqlStr = sqlStr + " ,'" + comm_cd + "'" + VbCrlf
''		sqlStr = sqlStr + " ," + defaultmargin + "" + VbCrlf
''		sqlStr = sqlStr + " ," + defaultsuplymargin + "" + VbCrlf
''		sqlStr = sqlStr + " ," + subtract + "" + VbCrlf
''		sqlStr = sqlStr + " ,'" + adminopen + "'" + VbCrlf
''		sqlStr = sqlStr + " ,'" + autojungsan + "'" + VbCrlf
''		sqlStr = sqlStr + " ,'" + autojungsandiv + "'" + VbCrlf
''		sqlStr = sqlStr + " ,'" + etcjunsandetail + "'" + VbCrlf
''
''		sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_user s" + VbCrlf
''		sqlStr = sqlStr + " left join [db_shop].[dbo].tbl_shop_designer d" + VbCrlf
''		sqlStr = sqlStr + " on d.shopid=s.userid and d.makerid='" + designer + "'" + VbCrlf
''		sqlStr = sqlStr + " where s.userid<>'" + shopid + "'" + VbCrlf
''		sqlStr = sqlStr + " and Left(s.userid,11)='" + Left(shopid,11) + "'" + VbCrlf
''		sqlStr = sqlStr + " and d.chargediv is null" + VbCrlf
''
''		rsget.Open sqlStr,dbget,1
''
''		''같은 정산 조건의 샾 업데이트
''		sqlStr = "update [db_shop].[dbo].tbl_shop_designer" + VbCrlf
''		sqlStr = sqlStr + " set defaultmargin=" + CStr(defaultmargin) + ","  + VbCrlf
''		sqlStr = sqlStr + " defaultsuplymargin=" + CStr(defaultsuplymargin) + ","  + VbCrlf
''		sqlStr = sqlStr + " subtract=" + CStr(subtract) + "," + VbCrlf
''		sqlStr = sqlStr + " adminopen='" + CStr(adminopen) + "'," + VbCrlf
''		sqlStr = sqlStr + " autojungsan='" + CStr(autojungsan) + "'," + VbCrlf
''		sqlStr = sqlStr + " autojungsandiv='" + CStr(autojungsandiv) + "'," + VbCrlf
''        sqlStr = sqlStr + " comm_cd='" + CStr(comm_cd) + "'" + VbCrlf
''
''		sqlStr = sqlStr + " where Left(shopid,11)='" + Left(shopid,11) + "'" + VbCrlf
''		sqlStr = sqlStr + " and makerid='" + designer + "'" + VbCrlf
''		sqlStr = sqlStr + " and chargediv='" + CStr(chargediv) + "'" + VbCrlf
''
''		rsget.Open sqlStr,dbget,1
''	end if

elseif (mode="del") then

    if (ChkDelShopMagin(shopid,designer)) then
    	response.write "<script>alert('삭제 되었습니다.');</script>"
    	response.write "<script>location.replace('" + refer + "');</script>"
	else
	    response.write "<script>alert('계약된 내역이 없습니다.');</script>"
    	response.write "<script>location.replace('" + refer + "');</script>"
	end if
elseif (mode="delArr") then
    dim delCnt
    delCnt =0
    cksel = split(cksel,",")
    shopid = split(shopid,",")

    cnt = UBound(cksel)

    for i=0 to cnt
        pp = cksel(i)
        if (ChkDelShopMagin(Trim(shopid(pp)), Trim(designer))) then
            delCnt = delCnt + 1
        end if
    next

    response.write "<script>alert('" + CStr(delCnt) + "건 삭제 되었습니다.');</script>"
    response.write "<script>location.replace('" + refer + "');</script>"
elseif (mode="arredit") then

    cksel = split(cksel,",")
    shopid = split(shopid,",")
    comm_cd = split(comm_cd,",")
    defaultmargin = split(defaultmargin,",")
    defaultsuplymargin = split(defaultsuplymargin,",")
    defaultbeasongdiv = split(defaultbeasongdiv,",")

    cnt = UBound(cksel)

    for i=0 to cnt
        pp = cksel(i)
        Call ChkUpdateShopMagin(Trim(shopid(pp)), Trim(designer), Trim(comm_cd(pp)), Trim(defaultmargin(pp)), Trim(defaultsuplymargin(pp)), Trim(defaultbeasongdiv(pp)), etcjunsandetail)
    next

    response.write "<script>alert('수정 되었습니다.');</script>"
    response.write "<script>location.replace('" + refer + "');</script>"
elseif (mode="defaultCenterMwdivChange") then


    Call ChkCenterMwDivChange(shopid,designer,defaultCenterMwdiv)

    response.write "<script>alert('수정 되었습니다.');</script>"
    response.write "<script>location.replace('" + refer + "');</script>"

elseif (mode="offadminopen") then
    dim offadminopen : offadminopen = requestCheckvar(request("offadminopen"),10)

    sqlStr = "update [db_shop].[dbo].tbl_shop_designer" + VbCrlf
    sqlStr = sqlStr + " set adminopen='" + offadminopen + "'" + VbCrlf
    sqlStr = sqlStr + " , itemregyn='" + itemregyn + "'" + VbCrlf
    sqlStr = sqlStr + " ,sdLastUpdate=getdate()" + VbCrlf
    sqlStr = sqlStr + " where shopid='" + shopid + "'" + VbCrlf
	sqlStr = sqlStr + " and makerid='" + designer + "'" + VbCrlf

	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr

    response.write "<script>alert('수정 되었습니다.');</script>"
    response.write "<script>location.replace('" + refer + "');</script>"
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
