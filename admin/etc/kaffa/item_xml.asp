<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 900
Response.CharSet = "utf-8"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/kaffa/kaffaCls.asp"-->
<%
dim oKaffatotalpage, oKaffaitem,i, j, k, buf, optbuf, optstr, vTemp, arrList, intLoop, vBody
dim keywordsStr, keywordsBuf, itemid

itemid = Request("itemid")

If itemid = "" Then
	dbget.close()
	Response.End
End IF

If isNumeric(itemid) = False Then
	dbget.close()
	Response.End
End IF

dim totalpage
dim maxpage
dim fso, FileName,tFile,appPath
dim readtextfile
Dim imakerid,kaffaPrice

dim sqlStr,ref
ref = Left(request.ServerVariables("REMOTE_ADDR"),250)

dim IsTheLastOption

sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('" + "ATS1-" + ref + "')"
'dbget.execute sqlStr

if (TRUE) then
	set oKaffaitem = new cKaffaItem
	oKaffaitem.FRectItemID = itemid
	oKaffaitem.GetAllKaffaItemList

	vBody = vBody & "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
	vBody = vBody & "<data>" & vbCrLf
	vBody = vBody & "	<api_key>$2a$08$ik.RQbF9tGCZibk7JnPueuG/8AIeuTDd.lgCP/fYuuZX7dnNuJRe6</api_key>" & vbCrLf

	For i=0 To oKaffaitem.FResultCount-1
	    imakerid    = oKaffaitem.FItemList(i).Fmakerid          '' 2013/06/28 수정
	    kaffaPrice  = oKaffaitem.FItemList(i).Forgsellcash      '' Fsellcash => Forgsellcash 2013/07/04 수정 소비가로 연동 요청 (고상미)

        ''해외상품가 기준으로 수정 필요.
	    if (LCASE(imakerid)="ithinkso") or (LCASE(imakerid)="antennashop") then
		    kaffaPrice = CLNG(kaffaPrice * 1.5)
		end if

		vBody = vBody & "	<product_id>" & oKaffaitem.FItemList(i).Fitemid & "</product_id>" & vbCrLf
		vBody = vBody & "	<language_code>ko</language_code>" & vbCrLf
		vBody = vBody & "	<product_code>" & oKaffaitem.FItemList(i).Fitemid & "</product_code>" & vbCrLf
		vBody = vBody & "	<product_name><![CDATA[" & fnReplaceTag(Trim(oKaffaitem.FItemList(i).Fitemname)) & "]]></product_name>" & vbCrLf
		vBody = vBody & "	<supply_price>0</supply_price>" & vbCrLf
		vBody = vBody & "	<consumer_price>" & kaffaPrice & "</consumer_price>" & vbCrLf
		vBody = vBody & "	<sale_price>" & kaffaPrice & "</sale_price>" & vbCrLf
		vBody = vBody & "	<minimum>1</minimum>" & vbCrLf
		vBody = vBody & "	<maximum>" & CHKIIF(oKaffaitem.FItemList(i).Fstockqty>30 and oKaffaitem.FItemList(i).Flimityn="Y",30,oKaffaitem.FItemList(i).Fstockqty) & "</maximum>" & vbCrLf
		vBody = vBody & "	<weight>" & (oKaffaitem.FItemList(i).Fitemweight*0.001) & "</weight>" & vbCrLf
		vBody = vBody & "	<hs_code></hs_code>" & vbCrLf
		vBody = vBody & "	<volume_x></volume_x>" & vbCrLf
		vBody = vBody & "	<volume_y></volume_y>" & vbCrLf
		vBody = vBody & "	<volume_h></volume_h>" & vbCrLf
		vBody = vBody & "	<reg_datetime>" & oKaffaitem.FItemList(i).FRegdate & "</reg_datetime>" & vbCrLf
		vBody = vBody & "	<edit_datetime>" & oKaffaitem.FItemList(i).FLastUpdate & "</edit_datetime>" & vbCrLf
		vBody = vBody & "	<production_date></production_date>" & vbCrLf
		vBody = vBody & "	<limit_date></limit_date>" & vbCrLf
		vBody = vBody & "	<discount>" & vbCrLf
		vBody = vBody & "		<start></start>" & vbCrLf
		vBody = vBody & "		<end></end>" & vbCrLf
		vBody = vBody & "		<price></price>" & vbCrLf
		vBody = vBody & "		<rate></rate>" & vbCrLf
		vBody = vBody & "	</discount>" & vbCrLf
		vBody = vBody & "	<brand_id></brand_id>" & vbCrLf
		vBody = vBody & "	<brand_code>" & oKaffaitem.FItemList(i).FMakerid & "</brand_code>" & vbCrLf
		vBody = vBody & "	<category>" & oKaffaitem.FItemList(i).Fkaffacate1 & "," & oKaffaitem.FItemList(i).Fkaffacate2 & "," & oKaffaitem.FItemList(i).Fkaffacate3 & "</category>" & vbCrLf
		vBody = vBody & "	<tags>" & vbCrLf
		If oKaffaitem.FItemList(i).Fkeywords <> "" Then
			For j = LBound(Split(oKaffaitem.FItemList(i).Fkeywords,",")) To UBound(Split(oKaffaitem.FItemList(i).Fkeywords,","))
				vBody = vBody & "		<tag>" & Split(oKaffaitem.FItemList(i).Fkeywords,",")(j) & "</tag>" & vbCrLf
			Next
		End If
		vBody = vBody & "	</tags>" & vbCrLf
		vBody = vBody & "	<images>" & vbCrLf
		vBody = vBody & "		<img>" & oKaffaitem.FItemList(i).Ficon1image & "</img>" & vbCrLf
		'vBody = vBody & "		<img>" & oKaffaitem.FItemList(i).Flistimage & "</img>" & vbCrLf
		vBody = vBody & "		<img>" & oKaffaitem.FItemList(i).Fbasicimage & "</img>" & vbCrLf
		If oKaffaitem.FItemList(i).Faddimage <> "" Then
			For j = LBound(Split(oKaffaitem.FItemList(i).Faddimage,",")) To UBound(Split(oKaffaitem.FItemList(i).Faddimage,","))
				vBody = vBody & "		<img>" & Split(oKaffaitem.FItemList(i).Faddimage,",")(j) & "</img>" & vbCrLf
			Next
		End If
		vBody = vBody & "	</images>" & vbCrLf
		vBody = vBody & "	<desc><![CDATA[" & fnReplaceTag(oKaffaitem.FItemList(i).FItemContent) & fnAddImageHTML(oKaffaitem.FItemList(i).Fitemid,oKaffaitem.FItemList(i).Fmainimage,oKaffaitem.FItemList(i).Fmainimage2) & "]]></desc>" & vbCrLf
		vBody = vBody & "	<additional>" & vbCrLf
		vBody = vBody & "		<made_in><![CDATA[" & fnReplaceTag(oKaffaitem.FItemList(i).Fsourcearea) & "]]></made_in>" & vbCrLf
		vBody = vBody & "		<manufacturer><![CDATA[" & fnReplaceTag(oKaffaitem.FItemList(i).FMakerName) & "]]></manufacturer>" & vbCrLf
		vBody = vBody & "		<stuff><![CDATA[" & fnReplaceTag(oKaffaitem.FItemList(i).Fitemsource) & "]]></stuff>" & vbCrLf
		vBody = vBody & "		<size_text><![CDATA[" & fnReplaceTag(oKaffaitem.FItemList(i).Fitemsize) & "]]></size_text>" & vbCrLf
		vBody = vBody & "	</additional>" & vbCrLf

		IF oKaffaitem.FItemList(i).Fitemoption = "0000" Then	'### 옵션없을때.
			vBody = vBody & "	<product_options>" & vbCrLf
			vBody = vBody & "		<product_option_name>" & vbCrLf
			vBody = vBody & "			<value>옵션</value>" & vbCrLf
			vBody = vBody & "		</product_option_name>" & vbCrLf
			vBody = vBody & "		<product_option_values>" & vbCrLf
			vBody = vBody & "			<group>" & vbCrLf
			vBody = vBody & "				<name>옵션</name>" & vbCrLf
			vBody = vBody & "				<value>단품</value>" & vbCrLf
			vBody = vBody & "			</group>" & vbCrLf
			vBody = vBody & "		</product_option_values>" & vbCrLf
			vBody = vBody & "	</product_options>" & vbCrLf
			vBody = vBody & "	<product_items>" & vbCrLf
			vBody = vBody & "		<item>" & vbCrLf
			'vBody = vBody & "			<is_manage_stock>" & CHKIIF(oKaffaitem.FItemList(i).Flimityn="Y","1","0") & "</is_manage_stock>" & vbCrLf
			vBody = vBody & "			<is_manage_stock>1</is_manage_stock>" & vbCrLf
			vBody = vBody & "			<price_sign></price_sign>" & vbCrLf
			vBody = vBody & "			<price></price>" & vbCrLf
			vBody = vBody & "			<point_sign></point_sign>" & vbCrLf
			vBody = vBody & "			<point></point>" & vbCrLf
			vBody = vBody & "			<product_item_code>" & oKaffaitem.FItemList(i).Fitemid & "-0000</product_item_code>" & vbCrLf
			vBody = vBody & "			<barcode></barcode>" & vbCrLf
			vBody = vBody & "			<qrcode_image></qrcode_image>" & vbCrLf
			vBody = vBody & "			<stock_cnt>" & oKaffaitem.FItemList(i).Fstockqty & "</stock_cnt>" & vbCrLf
			vBody = vBody & "			<item_options>" & vbCrLf
			vBody = vBody & "				<group>" & vbCrLf
			vBody = vBody & "					<name>옵션</name>" & vbCrLf
			vBody = vBody & "					<value>단품</value>" & vbCrLf
			vBody = vBody & "				</group>" & vbCrLf
			vBody = vBody & "			</item_options>" & vbCrLf
			vBody = vBody & "		</item>" & vbCrLf
			vBody = vBody & "	</product_items>" & vbCrLf
		Else
			'### '|option|' + cast(o.itemoption) + '|||' + o.optionname + '|||' +  cast(o.optaddprice) + '|||' + case when o.optionTypeName = '' then '선택' else o.optionTypeName + ' 선택' end as optionTypeName
			'### + '|||' + o.optlimityn + '|||' + cast(o.optlimitno as varchar(10)) + '|||' + cast(o.optlimitsold as varchar(10))
			'###
			'### 배열번호 -> itemoption : 0, optionname : 1, optaddprice : 2, optionTypeName : 3, optlimityn : 4, optlimitno : 5, optlimitsold : 6
			arrList = oKaffaitem.FItemList(i).Fitemoption

			vBody = vBody & "	<product_options>" & vbCrLf
			vBody = vBody & "		<product_option_name>" & vbCrLf

			If Split(Split(arrList,"|option|")(0),"|||")(3) = "" Then
				vBody = vBody & "			<value>선택</value>" & vbCrLf
			Else
				vBody = vBody & "			<value>" & Trim(Split(Split(arrList,"|option|")(0),"|||")(3)) & "</value>" & vbCrLf
			End If

			vBody = vBody & "		</product_option_name>" & vbCrLf
			vBody = vBody & "		<product_option_values>" & vbCrLf
			vBody = vBody & "			<group>" & vbCrLf

			If Split(Split(arrList,"|option|")(0),"|||")(3) = "" Then
				vBody = vBody & "				<name>선택</name>" & vbCrLf
			Else
				vBody = vBody & "				<name>" & Trim(Split(Split(arrList,"|option|")(0),"|||")(3)) & "</name>" & vbCrLf
			End If

			For j = LBound(Split(arrList,"|option|")) To UBound(Split(arrList,"|option|"))
				vBody = vBody & "				<value>" & Split(Split(arrList,"|option|")(j),"|||")(1) & "</value>" & vbCrLf
			Next

			vBody = vBody & "			</group>" & vbCrLf
			vBody = vBody & "		</product_option_values>" & vbCrLf
			vBody = vBody & "	</product_options>" & vbCrLf
			vBody = vBody & "	<product_items>" & vbCrLf

			For j = LBound(Split(arrList,"|option|")) To UBound(Split(arrList,"|option|"))
			vBody = vBody & "		<item>" & vbCrLf
			'vBody = vBody & "			<is_manage_stock>" & CHKIIF(Split(Split(arrList,"|option|")(j),"|||")(4)="Y","1","0") & "</is_manage_stock>" & vbCrLf
			vBody = vBody & "			<is_manage_stock>1</is_manage_stock>" & vbCrLf
			vBody = vBody & "			<price_sign></price_sign>" & vbCrLf
			vBody = vBody & "			<price>" & Split(Split(arrList,"|option|")(j),"|||")(2) & "</price>" & vbCrLf
			vBody = vBody & "			<point_sign></point_sign>" & vbCrLf
			vBody = vBody & "			<point></point>" & vbCrLf
			vBody = vBody & "			<product_item_code>" & oKaffaitem.FItemList(i).Fitemid & "-" & Replace(Split(Split(arrList,"|option|")(j),"|||")(0),"option|","") & "</product_item_code>" & vbCrLf
			vBody = vBody & "			<barcode></barcode>" & vbCrLf
			vBody = vBody & "			<qrcode_image></qrcode_image>" & vbCrLf
			vBody = vBody & "			<stock_cnt>" & fnChangeMinusJaeGo((CHKIIF(Split(Split(arrList,"|option|")(j),"|||")(4)="N",999,Split(Split(arrList,"|option|")(j),"|||")(5))) - (Split(Split(arrList,"|option|")(j),"|||")(6))) & "</stock_cnt>" & vbCrLf
			vBody = vBody & "			<item_options>" & vbCrLf
			vBody = vBody & "				<group>" & vbCrLf
			If Split(Split(arrList,"|option|")(j),"|||")(3) = "" Then
				vBody = vBody & "				<name>선택</name>" & vbCrLf
			Else
				vBody = vBody & "				<name>" & Trim(Split(Split(arrList,"|option|")(j),"|||")(3)) & "</name>" & vbCrLf
			End If
			vBody = vBody & "					<value>" & Split(Split(arrList,"|option|")(j),"|||")(1) & "</value>" & vbCrLf
			vBody = vBody & "				</group>" & vbCrLf
			vBody = vBody & "			</item_options>" & vbCrLf
			vBody = vBody & "		</item>" & vbCrLf
			Next

			vBody = vBody & "	</product_items>" & vbCrLf

			arrList = ""
		End If
	Next

	vBody = vBody & "</data>"

	set oKaffaitem = Nothing
end if

sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('" + "ATS2-" + ref + "')"
'dbget.execute sqlStr


'####### 뿌려지고나면 [db_item].[dbo].[tbl_kaffa_reg_item].[useyn] 값을 y로 바꿔줌.
sqlStr = "update db_item.dbo.tbl_kaffa_reg_item set useyn = 'y' where itemid = '" & itemid & "'"
dbget.execute sqlStr


Response.Write vBody


Function fnReplaceTag(v)
	Dim vTemp
	vTemp = ""
	vTemp = v
	vTemp = Replace(vTemp,"[","")
	vTemp = Replace(vTemp,"]","")
	vTemp = Replace(vTemp,"'","")
	vTemp = Replace(vTemp,chr(34),"")
	fnReplaceTag = vTemp
End Function

Function fnChangeMinusJaeGo(v)
	If v < 1 Then
		fnChangeMinusJaeGo = 0
	Else
		fnChangeMinusJaeGo = v
	End If
End Function

Function fnAddImageHTML(byval itemid, imagemain1, imagemain2)
		dim strSQL,ArrRows,i, FResultCount, vAddImage, vBody
		vBody = ""

		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid = " & CStr(itemid)
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.Locktype = adLockReadOnly
		rsget.Open strSQL, dbget

		If Not rsget.EOF Then
			ArrRows 	= rsget.GetRows
		End if
		rsget.close

		if isArray(ArrRows) then

			FResultCount = Ubound(ArrRows,2) + 1

			vBody = vBody & "<br /><br /><table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf

			'설명 이미지(추가)
			IF FResultCount > 0 THEN
				FOR i= 0 to FResultCount-1

					IF ArrRows(1,i)	= "1" Then
						vAddImage = "http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(itemid) & "/" & ArrRows(2,i)
					Else
						vAddImage = "http://webimage.10x10.co.kr/image/add" & Cstr(ArrRows(0,i)) & "/" & GetImageSubFolderByItemid(itemid) & "/" & ArrRows(2,i)
					End IF

					IF ArrRows(1,i) = 1 THEN
						vBody = vBody & "<tr><td align=""center"">" & vbCrLf
						vBody = vBody & "<img src=""" & vAddImage & """ border=""0"" style=""max-width:960px;"" />" & vbCrLf
						vBody = vBody & "</td></tr>" & vbCrLf
					End IF
				NEXT
			END IF

			'설명 이미지(기본)
			if ImageExists(imagemain1) then
				vBody = vBody & "<tr><td align=""center"">" & vbCrLf
				vBody = vBody & "<img src=""" & imagemain1 & """ border=""0"" id=""filemain"" style=""max-width:960px;"" />" & vbCrLf
				vBody = vBody & "</td></tr>" & vbCrLf
			end if
			if ImageExists(imagemain2) then
				vBody = vBody & "<tr><td align=""center"">" & vbCrLf
				vBody = vBody & "<img src=""" & imagemain2 & """ border=""0"" id=""filemain2"" style=""max-width:960px;"" />" & vbCrLf
				vBody = vBody & "</td></tr>" & vbCrLf
			end if

			vBody = vBody & "</table>"

		end if
	fnAddImageHTML = vBody
End Function

function ImageExists(byval iimg)
	if (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") then
		ImageExists = false
	else
		ImageExists = true
	end if
end function
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->