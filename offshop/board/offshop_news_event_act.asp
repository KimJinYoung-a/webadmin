<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionoffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<%

'// 사용안함, 2015-07-07, skyer9
dbget.close()
response.end

function FormatStr(n,orgData)
		dim tmp
		if (n-Len(CStr(orgData))) < 0 then
			FormatStr = CStr(orgData)
			Exit Function
		end if

		tmp = String(n-Len(CStr(orgData)), "0") & CStr(orgData)
		FormatStr = tmp
end Function

function CheckFiles(ifilesys, ifile, ifilesize)
	dim file1_size, file1_name
	dim extension

	if (ifile="") then
		CheckFiles =0
		exit function
	end if

	file1_size = CLng(ifilesize)
    file1_name = ifilesys.GetFileName(ifile)
    extension = LCase(Right(file1_name,3))

    if (file1_size>1024000) then
    	response.write "<script language='javascript'>alert('파일사이즈 1M까지 지원됩니다.'); history.go(-1);</script>"
        dbget.close()	:	response.End
    	exit function
    end if

    CheckFiles =0
end function

Function getFileExt(str)
	dim sp
	sp = split(str,".")
	getFileExt = sp(UBound(sp))
End Function

dim uploadForm,objFSO
Set uploadForm = Server.CreateObject("DEXT.FileUpload")
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

dim idx, mode, shopid, gubun, file1, dl_file1
dim userid,title,contents,enddate,isusing

idx = uploadForm("idx")
mode = uploadForm("mode")
shopid = uploadForm("shopid")
gubun = uploadForm("gubun")
userid = uploadForm("userid")
title = html2db(uploadForm("title"))
contents = html2db(uploadForm("contents"))
enddate = uploadForm("enddate")
isusing = uploadForm("isusing")
file1 = uploadForm("file1")
dl_file1 = uploadForm("dl_file1")

dim file1size
file1size =  uploadForm("file1").FileLen

dim ckret
ckret = CheckFiles(objFSO,file1,file1size)

dim updir
updir = replace(Server.MapPath("\uploadfile\staff\"),"webadmin","webadmin")

dim iid
dim sql

if (mode = "add") then

		sql = " insert into [db_board].[dbo].tbl_offshop_news_event(shopid, gubun, userid, title, contents, enddate) "
		sql = sql + " values('" + shopid + "','" + gubun + "','" + userid + "', '" + title + "', '" + contents + "','" + enddate + "') "
		rsget.Open sql, dbget, 1


		sqlStr = " select IDENT_CURRENT('[db_board].[dbo].tbl_offshop_news_event') as currid"
		rsget.Open sqlStr,dbget,1
		iid = rsget("currid")
		rsget.close

elseif  (mode = "edit") then

		sql = "update [db_board].[dbo].tbl_offshop_news_event " + VbCRlf
		sql = sql + " set shopid = '" + shopid + "'," + VbCRlf
		sql = sql + " gubun = '" + gubun + "'," + VbCRlf
		sql = sql + " title = '" + title + "'," + VbCRlf
		sql = sql + " contents = '" + contents + "', " + VbCRlf
		sql = sql + " enddate = '" + enddate + "', " + VbCRlf
		sql = sql + " isusing = '" + isusing + "' " + VbCRlf
		sql = sql + " where (idx = " + idx + ") " + VbCRlf
		rsget.Open sql, dbget, 1
end if


function DelProc(fieldname, idx)
	sql = " update [db_shop].[dbo].tbl_offshop_staff"
	sql = sql + " set " + fieldname + "=NULL"
	sql = sql + " where idx=" + CStr(idx)
	rsget.Open sql,dbget,1
end function

dim filenameolny, svrname_img

if (file1<>"") and (dl_file1<>"on") then
	filenameolny = "staff_img" & FormatStr(7,iid) & "." & getFileExt(file1)
	svrname_img = updir & "\" & filenameolny

	uploadForm("file1").saveas(svrname_img)


	sql = " update [db_shop].[dbo].tbl_offshop_staff"
	sql = sql + " set  file1 ='" + filenameolny + "'"
	sql = sql + " where idx=" + CStr(idx)
	rsget.Open sql,dbget,1

end if

if (dl_file1="on") then
	DelProc "file1", iid
end if

Set uploadForm = Nothing
Set objFSO = Nothing

response.write "<script>alert('적용 되었습니다.')</script>"
response.write "<script>location.replace('offshop_news_event_list.asp')</script>"

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
