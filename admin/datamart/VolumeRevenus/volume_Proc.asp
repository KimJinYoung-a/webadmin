<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
Dim idx , yyyy1 , mm1 , cdl , cdm , cds , wid , uid , regdate , lastupdate , volume , revenus , lastvolume , lastrevenus , updatecnt , mode ,yyyymm

idx = request("idx")
yyyy1 = request("yyyy1")
mm1 = request("mm1")
cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")
wid = request("wid")
uid = request("uid")
regdate = Now()
lastupdate = request("lastupdate")
volume = Replace(request("volume"),",","")
revenus = Replace(request("revenus"),",","")
lastvolume = request("lastvolume")
lastrevenus = request("lastrevenus")
updatecnt = request("updatecnt")

yyyymm = yyyy1&"-"&mm1

'rw idx & " : idx " 
'rw yyyy1 & " : yyyy1 "
'rw mm1 & " : mm1 "
'rw cdl & " : cdl "
'rw cdm & " : cdm "
'rw wid & " : wid "
'rw uid & " : uid "
'rw regdate & " : regdate "
'rw lastupdate & " : lastupdate "
'rw volume & " : volume "
'rw revenus & " : revenus "
'rw lastvolume & " : lastvolume "
'rw lastrevenus & " : lastrevenus "
'rw updatecnt & " : updatecnt "
'response.end

If updatecnt = "" Then updatecnt = 0 '수정

if idx = "" then
	idx = 0
end if

if idx = 0 then
	mode = "add"
Else
	mode = "edit"
end if

dim sqlStr

if (mode = "add") Then

	sqlStr = "select top 1 idx "  
	sqlStr = sqlStr + " from db_datamart.[dbo].tbl_mkt_monthly_volume_revenus" 
	sqlStr = sqlStr + " where yyyymm = '"&yyyymm&"' and cdl = '"&cdl&"' "
	If cdl = "110" then
	sqlStr = sqlStr + " and cdm = '"&cdm&"' "
	End If 
	db3_rsget.Open sqlStr, db3_dbget, 1
	If Not db3_rsget.Eof then
		idx = db3_rsget("idx")
	end If
	db3_rsget.close

	If idx = "" Then 
		sqlStr = " insert into db_datamart.dbo.tbl_mkt_monthly_volume_revenus" + VbCrlf
		sqlStr = sqlStr + " (yyyymm , cdl , cdm , cds , wid ,  volume , revenus)" + VbCrlf
		sqlStr = sqlStr + " values(" + VbCrlf
		sqlStr = sqlStr + " '" + yyyymm + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + cdl + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + cdm + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + cds + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + wid + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + volume + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + revenus + "'" + VbCrlf

		sqlStr = sqlStr + " )"

		db3_dbget.Execute sqlStr
		

		sqlStr = "select IDENT_CURRENT('db_datamart.[dbo].tbl_mkt_monthly_volume_revenus') as idx"
		db3_rsget.Open sqlStr, db3_dbget, 1
		If Not db3_rsget.Eof then
			idx = db3_rsget("idx")
		end if
		db3_rsget.close
	Else
	response.write "<script>alert('이미 등록 되어 있습니다. 수정을 해주세요');self.close();</script>"
	response.end
	End If 

elseif mode = "edit" then
   sqlStr = " update db_datamart.dbo.tbl_mkt_monthly_volume_revenus " + VbCrlf
   sqlStr = sqlStr + " set " + VbCrlf
   sqlStr = sqlStr + " yyyymm ='" + yyyymm + "'" + VbCrlf
   sqlStr = sqlStr + " ,cdl ='" + cdl + "'" + VbCrlf
   sqlStr = sqlStr + " ,cdm ='" + cdm + "'" + VbCrlf
   sqlStr = sqlStr + " ,cds ='" + cds + "'" + VbCrlf
   sqlStr = sqlStr + " ,wid ='" + wid + "'" + VbCrlf
   sqlStr = sqlStr + " ,uid ='" + uid + "'" + VbCrlf
   sqlStr = sqlStr + " ,volume ='" + volume + "'" + VbCrlf
   sqlStr = sqlStr + " ,revenus ='" + revenus + "'" + VbCrlf
   sqlStr = sqlStr + " ,lastvolume ='" + lastvolume + "'" + VbCrlf
   sqlStr = sqlStr + " ,lastrevenus ='" + lastrevenus + "'" + VbCrlf
   sqlStr = sqlStr + " ,lastupdate ='" + lastupdate + "'" + VbCrlf
   sqlStr = sqlStr + " ,updatecnt ='" + updatecnt + "'" + VbCrlf


   sqlStr = sqlStr + " where idx=" + CStr(idx)
	'response.write sqlStr
   db3_dbget.Execute sqlStr

end if


dim referer
referer = request.ServerVariables("HTTP_REFERER")
response.write "<script>alert('저장되었습니다.');self.close();</script>"
response.write "<script>location.href='/admin/datamart/volumerevenus/pop_write.asp?idx=" + Cstr(idx) + "&reload=on'</script>"
%>
<!-- #include virtual="/lib/db/db3close.asp" -->
