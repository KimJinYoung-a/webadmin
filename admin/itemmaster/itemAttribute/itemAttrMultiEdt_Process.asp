<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<%

dim refer, i, j
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr
dim mode
dim changeList

changeList = request.Form("changeList")

mode = requestCheckVar(request.Form("mode"), 32)
dim ChkIxCnt : ChkIxCnt = request.Form("chkix").count

dim chkidx,itemid, attrCd, checked

dim objCmd, returnValue, retErrText, retText, item
dim catecodelist, optionlistArr


if (mode = "itemattrmultiset") then
    ''EXEC sp_Ten_category_attrib_item_set @itemid=1472345
    '',@catecode='101102101101,101102101102,101102101103,101102101104,101102103106,101102103107'
    '',@optionlist='301001,1|301002,1|301003,1'

    catecodelist = request.Form("catecodelist")
    if (ChkIxCnt>0) then
		 for i=1 to ChkIxCnt
		 	chkidx			= request.Form("chkix")(i)
            itemid	        = Trim(request.Form("itemid")(chkidx+1))

            optionlistArr   = ""

            if (request.Form("chkAttrCd"&chkidx).count>0) then
                ''optionlistArr   = replace(request.Form("chkAttrCd"&chkidx),", ",",1|")
                ''if (optionlistArr<>"") then optionlistArr=optionlistArr&",1"

                for j=1 to request.form("chkAttrCd"&chkidx).count
                    optionlistArr = optionlistArr & request.form("chkAttrCd"&chkidx)(j)&",1|"
                next

                if Right(optionlistArr,1)="|" then optionlistArr=LEFT(optionlistArr,LEN(optionlistArr)-1)
            end if

            sqlStr = " exec [db_item].[dbo].[sp_Ten_category_attrib_item_set] '"&itemid&"','"&catecodelist&"','"&optionlistArr&"'"
	        dbget.Execute sqlStr

		 next
	else
		response.write "ERR:"
        dbget.close() : response.end
	end if

    response.write	"<script language='javascript'>"
	response.write	"	alert('저장되었습니다.'); "
	response.write	"	location.replace('" + CStr(refer) + "'); "
	response.write	"</script>"

elseif (mode="savechanged") then

    changeList = Split(changeList, "|")
    for i = 0 to UBound(changeList)
        if (changeList(i) <> "") then
            item = Split(changeList(i), ",")
            itemid = item(0)
            attrCd = item(1)
            checked = item(2)
            if checked = "true" then
                checked = "A"
            else
                checked = "D"
            end if

            sqlStr = " exec [db_item].[dbo].[usp_Ten_category_attrib_item_set] '" & itemid & "','" & attrCd & "','" & checked & "'"
	        dbget.Execute sqlStr
        end if
    next

    response.write	"<script language='javascript'>"
	response.write	"	alert('저장되었습니다.'); "
	response.write	"	location.replace('" + CStr(refer) + "'); "
	response.write	"</script>"

elseif (mode="TTTTTT") then
    response.write "잘못된 접근입니다."

    ''EXEC sp_Ten_category_attrib_item_set @itemid=1472345
    '',@catecode='101102101101,101102101102,101102101103,101102101104,101102103106,101102103107'
    '',@optionlist='301001,1|301002,1|301003,1'

	' Set objCmd = Server.CreateObject("ADODB.COMMAND")
	' 	With objCmd
	' 		.ActiveConnection = dbget
	' 		.CommandType = adCmdStoredProc
	' 		.CommandText = "db_item.[dbo].[sp_Ten_category_attrib_item_set]"
	' 		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
    '         .Parameters.Append .CreateParameter("@itemid", adInteger, adParamInput, , itemid)
    '         .Parameters.Append .CreateParameter("@catecode", adVarchar, adParamInput, 1000, exceptmakerid)
    '         .Parameters.Append .CreateParameter("@optionlist", adVarchar, adParamInput, 8000, exceptmakerid)

	' 		.Execute, , adExecuteNoRecords
	' 		End With
	' 	    returnValue = objCmd.Parameters("RETURN_VALUE").Value
	' Set objCmd = nothing

	' retText = "처리되었습니다."

	' if returnValue<0 then
	' 	retText = retErrText
	' end if
else

	response.write "잘못된 접근입니다."

end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
