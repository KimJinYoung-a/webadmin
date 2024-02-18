<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%

dim ogroup

function checkidexist(userid)
        dim sql

	sql = "select top 1 * from db_shop.dbo.tbl_shop_user where userid = '"&userid&"'"
        rsget.Open sql,dbget,1
                checkidexist = (not rsget.EOF)
        rsget.close
        if checkidexist then exit Function

	sql = "select top 1 * from [db_partner].[dbo].tbl_partner where id = '"&userid&"'"
        rsget.Open sql,dbget,1
                checkidexist = (not rsget.EOF)
        rsget.close
        if checkidexist then exit Function

        sql = "select top 1 userid from [db_user].[dbo].tbl_logindata where userid = '" + userid + "'"
        rsget.Open sql,dbget,1
                checkidexist = (not rsget.EOF)
        rsget.close
        if checkidexist then exit Function

        sql = "select userid from [db_user].[dbo].tbl_deluser where userid = '" + userid + "'"
        rsget.Open sql, dbget, 1
                checkidexist = (Not rsget.Eof)
        rsget.Close
        if checkidexist then exit Function
end function

function checksocnoexist(socno)
        dim sql

        sql = "select top 1 userid from [db_user].[dbo].tbl_user_c where socno = '" + socno + "'"
        rsget.Open sql,dbget,1

        checksocnoexist = (not rsget.EOF)

        rsget.close
end function


function checkspecialpass(target)
        dim buf, result, index

        index = 1
        do until index > len(target)
                buf = mid(target, index, cint(1))
                if (buf="'") or (buf="`") then
                        checkspecialpass = true
                        exit function
                else
                        result = false
                end if
                index = index + 1
        loop
        checkspecialpass = false
end function

function checkspecialchar(target)
        dim buf, result, index

        index = 1
        do until index > len(target)
                buf = mid(target, index, cint(1))
                if (lcase(buf) >= "a" and lcase(buf) <= "z") then
                        result = false
                elseif (buf >= "0" and buf <= "9") then
                        result = false
                ' 10x10_cs 때문에 추가
                elseif (buf = "_") then
                        result = false
                else
                        checkspecialchar = true
                        exit function
                end if
                index = index + 1
        loop
        checkspecialchar = false
end function

dim mode, uid, password, socno, pcuserdiv
mode = request("mode")
uid = request("uid")
password = request("password")
socno = request("socno")
pcuserdiv = request("pcuserdiv")

if (mode = "") then
	mode = "checkidpassword"
end if

if (mode = "checkidpassword") then
    if (chkPasswordComplex(uid,password)<>"") then
            response.write "<script>alert('" & chkPasswordComplex(uid,password) & "\n비밀번호를 확인후 다시 시도해주세요.')</script>"
			dbget.close()	:	response.End
	end if

	if (checkidexist(uid)) then
			response.write "<script>alert('이미 사용중이거나, 사용 할 수 없는 아이디입니다.')</script>"
			dbget.close()	:	response.End
	end if

	if pcuserdiv<>"900_21" then
		if (checkspecialchar(uid)) then
			response.write "<script>alert('해당ID로는 신청하실수가 없습니다.\r\n특수문자나 한글,한문을 사용하신 경우 변경하십시요.')</script>"
			dbget.close()	:	response.End
		end if
	end if

	if (checkspecialpass(password)) then
		response.write "<script>alert('해당 password로는 신청하실수가 없습니다.\r\n특수문자을 사용하신 경우 변경하십시요.')</script>"
		dbget.close()	:	response.End
	end if
elseif (mode = "CheckSocno") then
	set ogroup = new CPartnerGroup

	ogroup.FPageSize = 20
	ogroup.FCurrPage = 1
	ogroup.FRectsocno = socno

	ogroup.GetGroupInfoList

	if (ogroup.FResultCount > 0) then
			response.write "<script>alert('" & ogroup.FItemList(0).Fcompany_name & "(" & socno & ") : 이미 등록된 업체입니다.\n\n등록할 수 없습니다.')</script>"
			dbget.close()	:	response.End
	end if
elseif (mode = "CheckSocnoOnSave") then
	set ogroup = new CPartnerGroup

	ogroup.FPageSize = 20
	ogroup.FCurrPage = 1
	ogroup.FRectsocno = socno

	ogroup.GetGroupInfoList

	if (ogroup.FResultCount > 0) then
			response.write "<script>alert('" & ogroup.FItemList(0).Fcompany_name & "(" & socno & ") : 이미 등록된 업체입니다.\n\n등록할 수 없습니다.')</script>"
			dbget.close()	:	response.End
	end if
end if

%>
<script language="javascript">
try {
	// alert("mode : <%= mode %>");
	parent.AddProc("<%= mode %>");
} catch (err) {
	alert(err.message);
}
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
