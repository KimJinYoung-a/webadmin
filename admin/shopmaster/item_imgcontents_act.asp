<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

'// 사용안함, 2015-07-07, skyer9
dbget.close()
response.end

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

    if (file1_size>800000) then
    	response.write "<script language='javascript'>alert('파일사이즈 800,000Byte 까지 지원됩니다.'); history.go(-1);</script>"
        dbget.close()	:	response.End
    	exit function
    end if

    if ((extension <> "gif") and (extension <> "jpg") and (extension <> "bmp") and (extension <> "png")) then
    	response.write "<script language='javascript'>alert('이미지(gif,jpg,bmp,png) 화일만 지원됩니다.'); history.go(-1);</script>"
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

dim itemid,mode, img1, img2, imginfo

dim uploadForm,objFSO
Set uploadForm = Server.CreateObject("DEXT.FileUpload")
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

mode = uploadForm("mode")
itemid = uploadForm("itemid")
img1 = uploadForm("img1")
img2 = uploadForm("img2")

dim img1size,img2size
dim dl_img1,dl_img2
dim img1name,img2name

img1size = uploadForm("img1").FileLen
img2size = uploadForm("img2").FileLen

dl_img1 = uploadForm("dl_img1")
dl_img2 = uploadForm("dl_img2")

'response.write contents1
'dbget.close()	:	response.End

dim ckret
ckret = CheckFiles(objFSO,img1,img1size)
ckret = CheckFiles(objFSO,img2,img2size)

dim updir
updir = replace(Server.MapPath("\image\info\"),"partner","webimage")

dim sqlStr,Fimginfo,simginfo

if mode = "edit" then
'데이터 가져오기
	sqlStr = " select top 1 imginfo from [db_item].[dbo].tbl_item_image"
	sqlStr = sqlStr + " where itemid='" + CStr(itemid) + "'"
	rsget.Open sqlStr,dbget,1
	if  not rsget.EOF  then
		Fimginfo = rsget("imginfo")
	end if
	rsget.Close

	simginfo = split(Fimginfo,",")

	img1name = simginfo(0)
	img2name = simginfo(1)
end if
'response.write filename1 & "<br>"
'response.write filename2
'dbget.close()	:	response.End

'하다 말았음... --;
function DelProc(fieldname, gubun, itemid)
	dim sqlStr,delname
	if gubun = "img1" then
		delname = "," &  fieldname
	else
		delname = fieldname & ","
	end if

'response.write delname
'dbget.close()	:	response.End

	sqlStr = " update [db_item].[dbo].tbl_item_image"
	sqlStr = sqlStr + " set imginfo='" + Cstr(delname) + "'"
	sqlStr = sqlStr + " where itemid='" + CStr(itemid) + "'"
	rsget.Open sqlStr,dbget,1
end function

dim filenameolny, svrname_img

if (img1<>"" and dl_img1 <> "on") then

	filenameolny =  "imginfo1_" + Cstr(itemid) + "." + getFileExt(img1)
	img1name = "imginfo1_" + Cstr(itemid) + "." + getFileExt(img1)
	svrname_img = updir & "\" & filenameolny

	uploadForm("img1").saveas(svrname_img)
end if

if (img2<>"" and dl_img2 <> "on") then

	filenameolny =  "imginfo2_" + Cstr(itemid) + "." + getFileExt(img2)
	img2name = "imginfo2_" + Cstr(itemid) + "." + getFileExt(img2)
	svrname_img = updir & "\" & filenameolny

	uploadForm("img2").saveas(svrname_img)
end if

if mode = "add" then

	imginfo = img1name & "," & img2name

	sqlStr = " update [db_item].[dbo].tbl_item_image"
	sqlStr = sqlStr + " set imginfo ='" + imginfo + "'"
	sqlStr = sqlStr + " where itemid=" + CStr(itemid)
	rsget.Open sqlStr,dbget,1

else

	if (dl_img1="on") or (dl_img2="on") then
		if (dl_img1="on") then
			DelProc img2name,"img1", itemid
		end if
		if (dl_img2="on") then
			DelProc img1name,"img2", itemid
		end if
	else

	imginfo = img1name & "," & img2name

	sqlStr = " update [db_item].[dbo].tbl_item_image"
	sqlStr = sqlStr + " set imginfo ='" + imginfo + "'"
	sqlStr = sqlStr + " where itemid=" + CStr(itemid)
	rsget.Open sqlStr,dbget,1

	end if

end if

Set uploadForm = Nothing
Set objFSO = Nothing

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>

<script language="JavaScript">
<!--
alert("데이터를 저장하였습니다.");
location.replace("item_imgcontents_list.asp");
//-->
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
