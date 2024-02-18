<%@ codepage="65001" language=vbscript %>
<% option explicit %>
<%
response.Charset="UTF-8"
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/apps/academy/lib/inc_const.asp" -->
<!-- #include virtual="/apps/academy/lib/smsLib.asp"-->
<!-- #include virtual="/apps/academy/lib/maillib.asp"-->
<!-- #include virtual="/apps/academy/lib/mailLib2.asp"-->
<!-- #include virtual="/apps/academy/lib/mailFunc_Designer.asp"-->
<%
'개발서버인 경우 메일/SMS 발송이 안되도록 되어 있다.
Dim SENDMAIL_YN
if (application("Svr_Info")	= "Dev") then
SENDMAIL_YN = "N"		'Y 인 경우 개발서버에서도 이메일을 발송하게 한다.
Else
SENDMAIL_YN = "Y"
End If

Dim sqlStr,ix, i, mibeasongSoldOutExists, AssignedRow, GetOrderStateNum
Dim MakerID, FailRow, mode
dim iMailOrderserialArr : iMailOrderserialArr=""
mode=requestCheckVar(Request.Form("mode"),12)
MakerID = requestCheckVar(request.cookies("partner")("userid"),32)
If mode="edit" Then SENDMAIL_YN="N" '수정일 경우 메일 발송 안함
'배열로 처리
ReDim DetailIDX(Request.Form("detailidx").Count)
ReDim OrderSerial(Request.Form("orderserial").Count)
ReDim SongjangDIV(Request.Form("songjangdiv").Count)
ReDim SongjangNO(Request.Form("songjangno").Count)
For ix=1 To Request.Form("detailidx").Count
	DetailIDX(ix) = Request.Form("detailidx")(ix)
	OrderSerial(ix) = Request.Form("orderserial")(ix)
	SongjangDIV(ix) = Request.Form("songjangdiv")(ix)
	SongjangNO(ix) = Request.Form("songjangno")(ix)
Next
FailRow=0

If MakerID<>"" Then
	'DB에 처리
	For ix=1 to Request.Form("detailidx").Count
		If DetailIDX(ix)<>"" And OrderSerial(ix)<>"" And SongjangDIV(ix)<>"" And SongjangNO(ix)<>"" Then
		   ''품절출고 불가 등록된경우 SKIP
			mibeasongSoldOutExists = false
			sqlStr = "select count(*) as CNT from db_academy.dbo.tbl_academy_mibeasong_list" & VbCRLF
			sqlStr = sqlStr + " where detailidx=" & Trim(DetailIDX(ix))  & VbCRLF
			sqlStr = sqlStr + " and orderserial='" & Trim(OrderSerial(ix)) & "'" & VbCRLF
			sqlStr = sqlStr + " and code='05'" & VbCRLF
			rsACADEMYget.CursorLocation = adUseClient
			rsACADEMYget.Open sqlStr, dbACADEMYget, adOpenForwardOnly
			if Not rsACADEMYget.Eof then
				mibeasongSoldOutExists = rsACADEMYget("CNT")>0
			end if
			rsACADEMYget.close

			if (mibeasongSoldOutExists) then
				FailRow = FailRow + 1
			ELSE
				''중복메일 방지용.
				sqlStr = "select d.orderserial from [db_academy].[dbo].tbl_academy_order_detail d"
				sqlStr = sqlStr + "     Join [db_academy].[dbo].tbl_academy_order_master m"
				sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
				sqlStr = sqlStr + " where d.orderserial='" & Trim(OrderSerial(ix)) & "'" & VbCRLF
				sqlStr = sqlStr + " and d.detailidx =" & Trim(DetailIDX(ix))  & VbCRLF
				sqlStr = sqlStr + " and d.makerid='" & Makerid & "'"
				sqlStr = sqlStr + " and d.cancelyn<>'Y'"
				sqlStr = sqlStr + " and m.cancelyn='N'"      '''취소 이전내역만.
				rsACADEMYget.CursorLocation = adUseClient
				rsACADEMYget.Open sqlStr, dbACADEMYget, adOpenForwardOnly
				if Not rsACADEMYget.Eof then
					if Not (InStr(iMailOrderserialArr,rsACADEMYget("orderserial") + ",")>0) then
						iMailOrderserialArr = iMailOrderserialArr + rsACADEMYget("orderserial") + ","
					end if
				end if
				rsACADEMYget.close
				
				sqlStr = "update D" & VbCRLF
				sqlStr = sqlStr + " set currstate='7'" & VbCRLF
				sqlStr = sqlStr + " ,songjangno='" & Trim(SongjangNO(ix)) & "'" & VbCRLF
				sqlStr = sqlStr + " ,songjangdiv='" & Trim(SongjangDIV(ix)) & "'" & VbCRLF
				sqlStr = sqlStr + " ,beasongdate=IsNULL(beasongdate,getdate())" & VbCRLF
				sqlStr = sqlStr + " ,passday=IsNULL(db_academy.dbo.fn_Ten_NetWorkDays((select convert(varchar(10),baljudate,21) from db_academy.dbo.tbl_academy_order_master where orderserial='" & Trim(OrderSerial(ix)) & "'),IsNULL(convert(varchar(10),beasongdate,21),convert(varchar(10),getdate(),21))),0)"& VbCRLF
				sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_detail D"
				sqlStr = sqlStr + "     Join [db_academy].[dbo].tbl_academy_order_master m"
				sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
				sqlStr = sqlStr + " where d.orderserial='" & Trim(OrderSerial(ix)) & "'" & VbCRLF
				sqlStr = sqlStr + " and d.detailidx =" & Trim(DetailIDX(ix))  & VbCRLF
				sqlStr = sqlStr + " and d.makerid='" & MakerID & "'"
				sqlStr = sqlStr + " and d.cancelyn<>'Y'"
				sqlStr = sqlStr + " and m.cancelyn='N'"      '''취소 이전내역만.
				dbACADEMYget.Execute sqlStr
			END IF
		End If
	Next
	'' ipkumdiv 8 추가
	sqlStr = "update [db_academy].[dbo].tbl_academy_order_master" & VbCRLF
	sqlStr = sqlStr + " set  ipkumdiv='7'" & VbCRLF
	sqlStr = sqlStr + " , beadaldate=getdate()" & VbCRLF
	sqlStr = sqlStr + " where orderserial in (" & VbCRLF
	sqlStr = sqlStr + "     select m.orderserial" & VbCRLF
	sqlStr = sqlStr + "     from [db_academy].[dbo].tbl_academy_order_master m" & VbCRLF
	sqlStr = sqlStr + "         left join [db_academy].[dbo].tbl_academy_order_detail d" & VbCRLF
	sqlStr = sqlStr + "         on m.orderserial=d.orderserial" & VbCRLF
	sqlStr = sqlStr + "     where m.orderserial='" & Trim(Request.Form("orderserial")) & "'" & VbCRLF
	sqlStr = sqlStr + "     and m.cancelyn='N'" & VbCRLF
	sqlStr = sqlStr + "     and jumundiv<>9" & VbCRLF
	sqlStr = sqlStr + "     and d.itemid<>0" & VbCRLF
	sqlStr = sqlStr + "     group by m.orderserial" & VbCRLF
	sqlStr = sqlStr + "     having sum(case when IsNull(d.currstate,'0')<>'7' then 1 else 0 end )>0" & VbCRLF
	sqlStr = sqlStr + " ) "
	dbACADEMYget.Execute sqlStr

	sqlStr = "update [db_academy].[dbo].tbl_academy_order_master" & VbCRLF
	sqlStr = sqlStr + " set  ipkumdiv='8'" & VbCRLF
	sqlStr = sqlStr + " , beadaldate=getdate()" & VbCRLF
	sqlStr = sqlStr + " where orderserial in (" & VbCRLF
	sqlStr = sqlStr + "     select m.orderserial" & VbCRLF
	sqlStr = sqlStr + "     from [db_academy].[dbo].tbl_academy_order_master m" & VbCRLF
	sqlStr = sqlStr + "         left join [db_academy].[dbo].tbl_academy_order_detail d" & VbCRLF
	sqlStr = sqlStr + "         on m.orderserial=d.orderserial" & VbCRLF
	sqlStr = sqlStr + "     where m.orderserial='" & Trim(Request.Form("orderserial")) & "'" & VbCRLF
	sqlStr = sqlStr + "     and m.cancelyn='N'" & VbCRLF
	sqlStr = sqlStr + "     and m.jumundiv<>9" & VbCRLF
	sqlStr = sqlStr + "     and d.itemid<>0" & VbCRLF
	sqlStr = sqlStr + "     and d.cancelyn<>'Y'" & VbCRLF
	sqlStr = sqlStr + "     group by m.orderserial" & VbCRLF
	sqlStr = sqlStr + "     having sum(case when IsNull(d.currstate,'0')<>'7' then 1 else 0 end )=0"
	sqlStr = sqlStr + " ) "
	dbACADEMYget.Execute sqlStr, AssignedRow

	If AssignedRow>0 Then
		sqlStr = sqlStr + "update [db_academy].[dbo].[tbl_academy_app_iconbadge_count]" + vbCrlf
		sqlStr = sqlStr + "	set ordercnt=ordercnt-"+Cstr(AssignedRow)+ vbCrlf
		sqlStr = sqlStr + "	where makerid='" + CStr(MakerID) + "'" + vbCrlf
		dbACADEMYget.Execute sqlStr
	End If

	sqlStr = "select mibaljucnt, ordercnt from [db_academy].[dbo].[tbl_academy_app_iconbadge_count] where makerid='" + CStr(MakerID) + "'" + vbCrlf
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	if not rsACADEMYget.EOF then
		GetOrderStateNum=rsACADEMYget("mibaljucnt")+rsACADEMYget("ordercnt")
	Else
		GetOrderStateNum=0
	End If
	rsACADEMYget.close
    ''#################################################
    ''메일보내기 루프
    ''#################################################
    iMailOrderserialArr = split(iMailOrderserialArr,",")
    if IsArray(iMailOrderserialArr) then
        for i=LBound(iMailOrderserialArr) to UBound(iMailOrderserialArr)
            if Trim(iMailOrderserialArr(i))<>"" then
                On Error resume Next
                if (SENDMAIL_YN = "Y") then
                    call fcSendMail_UpcheSendItem(iMailOrderserialArr(i), MakerID)
                end if
                on Error Goto 0
            end if
        next
    end if

End If
%>
<script>
<!--
parent.fnSongjangInputEnd(<%=GetOrderStateNum%>);
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->