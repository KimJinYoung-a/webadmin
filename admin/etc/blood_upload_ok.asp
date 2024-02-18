<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
Function getFileExt(str)
	dim sp
	sp = split(str,".")
	getFileExt = sp(UBound(sp))

	if (Lcase(getFileExt)<>"gif") and (Lcase(getFileExt)<>"jpg") then
		response.write "<script>alert('gif나 jpg 파일만 업로드 가능');</script>"
		response.write "<script>history.back()</scirpt>"
		dbget.close()	:	response.End
	end if
End Function

dim uploadForm,objFSO
Set uploadForm = Server.CreateObject("SiteGalaxyUpload.Form")
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

dim id
dim mode

dim title ,sex ,isusing

id = uploadForm.item("eid")
mode = uploadForm.item("mode")

title = uploadForm.item("title")
sex = uploadForm.item("sex")
isusing = uploadForm.item("isusing")

dim imageAmain,A_main
dim imageAlist,A_list
dim imageBmain,B_main
dim imageBlist,B_list
dim imageOmain,O_main
dim imageOlist,O_list
dim imageABmain,AB_main
dim imageABlist,AB_list

imageAmain = uploadForm.item("imageAmain")
A_main = uploadForm.item("A_main")
imageAlist = uploadForm.item("imageAlist")
A_list = uploadForm.item("A_list")
imageBmain = uploadForm.item("imageBmain")
B_main = uploadForm.item("B_main")
imageBlist = uploadForm.item("imageBlist")
B_list = uploadForm.item("B_list")
imageOmain = uploadForm.item("imageOmain")
O_main = uploadForm.item("O_main")
imageOlist = uploadForm.item("imageOlist")
O_list = uploadForm.item("O_list")
imageABmain = uploadForm.item("imageABmain")
AB_main = uploadForm.item("AB_main")
imageABlist = uploadForm.item("imageABlist")
AB_list = uploadForm.item("AB_list")

dim updir_title
updir_title = replace(Server.MapPath("\image\blood\"),"partner","www")

dim sqlStr
if (mode="add") then
	sqlStr = "insert into [db_contents].[dbo].tbl_blood_master(title,sex,isusing)" + vbcrlf
	sqlStr = sqlStr + " values(" + vbcrlf
	sqlStr = sqlStr + " '" + Cstr(title) + "'," + vbcrlf
	sqlStr = sqlStr + " '" + Cstr(sex) + "'," + vbcrlf
	sqlStr = sqlStr + " '" + Cstr(isusing) + "'" + vbcrlf
	sqlStr = sqlStr + ")"

elseif mode="edit" then
	sqlStr = "update [db_contents].[dbo].tbl_blood_master" + vbcrlf
	sqlStr = sqlStr + " set title ='" + Cstr(title) + "'," + vbcrlf
	sqlStr = sqlStr + " sex='" + Cstr(sex) + "'," + vbcrlf
	sqlStr = sqlStr + " isusing='" + Cstr(isusing) + "'" + vbcrlf
	sqlStr = sqlStr + " where idx=" + CStr(id)
end if
'response.write sqlstr
rsget.Open sqlStr,dbget,1

'#################### 파일올리기
dim filenameolny,svrname
'----------------------------------------------------------------------------------
if (imageAmain<>"") and (A_main<>"on") then
	filenameolny =  "blood" + Format00(6,id) + "Amain" + "." + getFileExt(imageAmain)
	svrname = updir_title & "\" & filenameolny

	uploadform("imageAmain").saveas(svrname)

	sqlStr = " update [db_contents].[dbo].[tbl_blood_master]" + VbCrlf
	sqlStr = sqlStr + " set Amain='" + filenameolny + "'" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(id) + "" + VbCrlf

	rsget.Open sqlStr,dbget,1
end if

if (A_main="on") then
	sqlStr = " update [db_contents].[dbo].[tbl_blood_master]" + VbCrlf
	sqlStr = sqlStr + " set Amain=NULL" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(id) + "" + VbCrlf

	rsget.Open sqlStr,dbget,1
end if

if (imageAlist<>"") and (A_list<>"on") then
	filenameolny =  "blood" + Format00(6,id) + "Alist" + "." + getFileExt(imageAlist)
	svrname = updir_title & "\" & filenameolny

	uploadform("imageAlist").saveas(svrname)

	sqlStr = " update [db_contents].[dbo].[tbl_blood_master]" + VbCrlf
	sqlStr = sqlStr + " set Alist='" + filenameolny + "'" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(id) + "" + VbCrlf

	rsget.Open sqlStr,dbget,1
end if

if (A_list="on") then
	sqlStr = " update [db_contents].[dbo].[tbl_blood_master]" + VbCrlf
	sqlStr = sqlStr + " set Alist=NULL" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(id) + "" + VbCrlf

	rsget.Open sqlStr,dbget,1
end if
'----------------------------------------------------------------------------------
if (imageBmain<>"") and (B_main<>"on") then
	filenameolny =  "blood" + Format00(6,id) + "Bmain" + "." + getFileExt(imageBmain)
	svrname = updir_title & "\" & filenameolny

	uploadform("imageBmain").saveas(svrname)

	sqlStr = " update [db_contents].[dbo].[tbl_blood_master]" + VbCrlf
	sqlStr = sqlStr + " set Bmain='" + filenameolny + "'" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(id) + "" + VbCrlf

	rsget.Open sqlStr,dbget,1
end if

if (B_main="on") then
	sqlStr = " update [db_contents].[dbo].[tbl_blood_master]" + VbCrlf
	sqlStr = sqlStr + " set Bmain=NULL" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(id) + "" + VbCrlf

	rsget.Open sqlStr,dbget,1
end if

if (imageBlist<>"") and (B_list<>"on") then
	filenameolny =  "blood" + Format00(6,id) + "Blist" + "." + getFileExt(imageBlist)
	svrname = updir_title & "\" & filenameolny

	uploadform("imageBlist").saveas(svrname)

	sqlStr = " update [db_contents].[dbo].[tbl_blood_master]" + VbCrlf
	sqlStr = sqlStr + " set Blist='" + filenameolny + "'" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(id) + "" + VbCrlf

	rsget.Open sqlStr,dbget,1
end if

if (B_list="on") then
	sqlStr = " update [db_contents].[dbo].[tbl_blood_master]" + VbCrlf
	sqlStr = sqlStr + " set Blist=NULL" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(id) + "" + VbCrlf

	rsget.Open sqlStr,dbget,1
end if
'----------------------------------------------------------------------------------
if (imageOmain<>"") and (O_main<>"on") then
	filenameolny =  "blood" + Format00(6,id) + "Omain" + "." + getFileExt(imageBmain)
	svrname = updir_title & "\" & filenameolny

	uploadform("imageOmain").saveas(svrname)

	sqlStr = " update [db_contents].[dbo].[tbl_blood_master]" + VbCrlf
	sqlStr = sqlStr + " set Omain='" + filenameolny + "'" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(id) + "" + VbCrlf

	rsget.Open sqlStr,dbget,1
end if

if (O_main="on") then
	sqlStr = " update [db_contents].[dbo].[tbl_blood_master]" + VbCrlf
	sqlStr = sqlStr + " set Omain=NULL" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(id) + "" + VbCrlf

	rsget.Open sqlStr,dbget,1
end if

if (imageOlist<>"") and (O_list<>"on") then
	filenameolny =  "blood" + Format00(6,id) + "Olist" + "." + getFileExt(imageBlist)
	svrname = updir_title & "\" & filenameolny

	uploadform("imageOlist").saveas(svrname)

	sqlStr = " update [db_contents].[dbo].[tbl_blood_master]" + VbCrlf
	sqlStr = sqlStr + " set Olist='" + filenameolny + "'" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(id) + "" + VbCrlf

	rsget.Open sqlStr,dbget,1
end if

if (O_list="on") then
	sqlStr = " update [db_contents].[dbo].[tbl_blood_master]" + VbCrlf
	sqlStr = sqlStr + " set Olist=NULL" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(id) + "" + VbCrlf

	rsget.Open sqlStr,dbget,1
end if
'----------------------------------------------------------------------------------
if (imageABmain<>"") and (AB_main<>"on") then
	filenameolny =  "blood" + Format00(6,id) + "ABmain" + "." + getFileExt(imageBmain)
	svrname = updir_title & "\" & filenameolny

	uploadform("imageABmain").saveas(svrname)

	sqlStr = " update [db_contents].[dbo].[tbl_blood_master]" + VbCrlf
	sqlStr = sqlStr + " set ABmain='" + filenameolny + "'" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(id) + "" + VbCrlf

	rsget.Open sqlStr,dbget,1
end if

if (AB_main="on") then
	sqlStr = " update [db_contents].[dbo].[tbl_blood_master]" + VbCrlf
	sqlStr = sqlStr + " set ABmain=NULL" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(id) + "" + VbCrlf

	rsget.Open sqlStr,dbget,1
end if

if (imageABlist<>"") and (AB_list<>"on") then
	filenameolny =  "blood" + Format00(6,id) + "ABlist" + "." + getFileExt(imageBlist)
	svrname = updir_title & "\" & filenameolny

	uploadform("imageABlist").saveas(svrname)

	sqlStr = " update [db_contents].[dbo].[tbl_blood_master]" + VbCrlf
	sqlStr = sqlStr + " set ABlist='" + filenameolny + "'" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(id) + "" + VbCrlf

	rsget.Open sqlStr,dbget,1
end if

if (AB_list="on") then
	sqlStr = " update [db_contents].[dbo].[tbl_blood_master]" + VbCrlf
	sqlStr = sqlStr + " set ABlist=NULL" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(id) + "" + VbCrlf

	rsget.Open sqlStr,dbget,1
end if
'----------------------------------------------------------------------------------

Set uploadForm = Nothing
Set objFSO = Nothing

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language='javascript'>
alert('저장되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->