<%
'//��ǰ�ڵ� �˻��� �޸�����ó��
Function fnForDBItmeid(ByVal Itemid) 
  dim iA ,arrTemp,arrItemid
  	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
	
	fnForDBItmeid = itemid
End Function

'//�Ǹ�
Sub sbOptSellYN(selectedId)
   dim tmp_str,query1
%>  
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >�Ǹ�</option>
   <option value="S" <% if selectedId="S" then response.write "selected" %> >�Ͻ�ǰ��</option>
   <option value="N" <% if selectedId="N" then response.write "selected" %> >ǰ��</option>
   <option value="YS" <% if selectedId="YS" then response.write "selected" %> >�Ǹ�+�Ͻ�ǰ��</option> 
<%
End Sub

'//����
Sub sbOptDanjongYN(selectedId)
   dim tmp_str,query1
   %> 
   <option value="N" <% if selectedId="N" then response.write "selected" %> >������</option>
   <option value="S" <% if selectedId="S" then response.write "selected" %> >������</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >����</option>
   <option value="M" <% if selectedId="M" then response.write "selected" %> >MDǰ��</option>
   <option value="YM" <% if selectedId="YM" then response.write "selected" %> >����+MDǰ��</option>
   <option value="SN" <% if selectedId="SN" then response.write "selected" %> >�����ƴ�</option> 
   <%
End Sub

'//���� 
Sub sbOptLimitYN(selectedId)
   dim tmp_str,query1
   %> 
   <option value="N" <% if selectedId="N" then response.write "selected" %> >������</option>
   <option value="Y" <% if selectedId="Y" then response.write "selected" %> >����</option>
   <option value="Y0" <% if selectedId="Y0" then response.write "selected" %> >����(0)</option> 
   <%
End Sub

'//�ŷ�����
Sub sbOptMWU(selectedId)
   dim tmp_str,query1
   %> 
   <option value="MW" <% if selectedId="MW" then response.write "selected" %> >����+Ư��</option>
   <option value="W" <% if selectedId="W" then response.write "selected" %> >Ư��</option>
   <option value="M" <% if selectedId="M" then response.write "selected" %> >����</option>
   <option value="U" <% if selectedId="U" then response.write "selected" %> >��ü</option> 
   <%
End Sub

'// ���п� ���� ���ڿ� ���� ����
function fnSetColor(str, div)
	Select Case div
		Case "yn"
			if str<>"Y" or isNull(str) then
				fnSetColor = "<span class=""cRd1"">" & str & "</span>"
			else
				fnSetColor = "<span class=""cBl1"">" & str & "</span>"
			end if
		Case "mw"
			Select Case str
				Case "M"
					fnSetColor = "<span class=""cRd1"">����</span>"
				Case "W"
					fnSetColor = "<span class=""cGn1"">Ư��</span>"
				Case "U"
					fnSetColor = "<span class=""cBl1"">��ü</span>"
			end Select
		Case "tx"
			if str="Y" then
				fnSetColor = "<Font color=#808080>����</font>"
			elseif str="N" then
				fnSetColor = "<Font color=#F08050>�鼼</font>"
			else
			    fnSetColor = str
			end if
		Case "dj"
			if str="Y" then
				fnSetColor = "<Font color=#33CC33>����</font>"
			elseif str="S" then
				fnSetColor = "<Font color=#3333CC>������</font>"
			elseif str="M" then
				fnSetColor = "<Font color=#CC3333>MDǰ��</font>"
			end if
		Case "delivery"
			IF str THEN
				fnSetColor = "<Font color=#F08050>��ü</font>"
			ELSE
				fnSetColor = "<Font color=#5080F0>10x10</font>"
			end IF
		Case "sellyn"
			IF str="Y" THEN
				fnSetColor = "<span class=""cBk1"">�Ǹ�</span>"
			elseif str="S" then
			    fnSetColor = "<span class=""cBl1"">�Ͻ�ǰ��</span>"
			else 
				fnSetColor = "<span class=""cRd1"">ǰ��</span>"    
			end IF
		Case "cancelyn"
			IF str="N" THEN
				fnSetColor = "<Font color=#000000>����</font>"
			elseif str="D" then
			    fnSetColor = "<Font color=#FF0000>����</font>"
			elseif str="Y" then
			    fnSetColor = "<Font color=#FF0000>���</font>"
			elseif str="A" then
			    fnSetColor = "<Font color=#FF0000>�߰�</font>"
			end IF
	end Select
end Function


'//�������� ��� 
function fnLimitEa(ByVal LimitNo,ByVal LimitSold)
    if (LimitNo-LimitSold<1) then
        fnLimitEa = 0
    else
        fnLimitEa = LimitNo-LimitSold
    end if
end function

'//�ŷ�����
public function fnmwdivName(ByVal v) 
	if v="M" then
		fnmwdivName = "����"
	elseif v="U" then
		fnmwdivName = "��ü"
	elseif v="W" then
		fnmwdivName = "Ư��" 
	end if
end function
 

'//��۱���
public function fnDeliveryName(ByVal v) 
	If v = "1" Then
		fnDeliveryName ="�ٹ����ٹ��"
	ElseIf v = "2" Then
		fnDeliveryName ="��ü(����)���"
	ElseIf v = "4" Then
		fnDeliveryName ="�ٹ����ٹ�����"
	ElseIf v = "9" Then
		fnDeliveryName ="��ü���ǹ��(���� ��ۺ�ΰ�)"
	ElseIf v = "7" Then
		fnDeliveryName ="��ü���ҹ��"
	End If
end function




'// �̹��� Ǯ��� ��������(GetImageSubFolderByItemid -> /lib/util/htmllib.asp ����)
Function fnGetSmallImage(ByVal ImgName, ByVal Itemid)
	fnGetSmallImage = webImgUrl + "/image/small/" + GetImageSubFolderByItemid(Itemid) + "/" + ImgName
End Function


'//��ǰ �ǸŰ���ǥ��
Sub SbDispItemSellPrice(ByVal sailyn, ByVal orgPrice,  ByVal sailprice, ByVal itemCouponyn, ByVal itemcoupontype, ByVal itemcouponvalue)
%>
 <%=formatnumber(orgPrice,0)%><%'���ΰ�
	if sailyn="Y" then
%><br /><span class="cOr1">(<%=CLng((orgPrice-sailprice)/orgPrice*100)%>%��)<%=FormatNumber(sailprice,0) %></span>
<%		
	end if
	'������ 
	Dim discountprice
	if itemCouponyn="Y" then
		IF itemcoupontype = "1" or itemcoupontype ="2" THEN 
			IF itemcoupontype = "1" THEN ''% ����
				discountprice = CLng(orgPrice*itemcouponvalue/100)
			ELSE ''�� ����
				discountprice = itemcouponvalue
			END IF		
%><br /><span class="cBl1">(��)<%=FormatNumber(orgPrice-discountprice ,0)%></span>
<%
		END IF 
	end if 
End Sub


'//��ǰ ���ް�ǥ��
Sub SbDispItemSupplyPrice(ByVal sailyn, ByVal orgsuplycash, ByVal sailsuplycash, ByVal itemCouponyn, ByVal itemcoupontype, ByVal couponbuyprice)
%>
	<%=formatnumber(orgsuplycash,0)%><%	'����
	 if sailyn="Y" then 
%><br /><span class="cOr1"><%=FormatNumber(sailsuplycash,0) %></span>
<%	end if  
	'������
    if itemCouponyn="Y" then
		IF itemcoupontype = "1" or itemcoupontype ="2" THEN 
			if couponbuyprice=0 or isNull(couponbuyprice) then
		%><br /><span class="cBl1"><%=FormatNumber(orgsuplycash,0)%></span>
		<%	else %><br /><span class="cBl1"><%=FormatNumber(couponbuyprice,0)%></span>
		<%	end if
		END IF 
	END IF 
End Sub 

'//��ǰ �ǸŰ���ǥ��
Sub SbDispItemSellSalePrice(ByVal sailyn, ByVal orgPrice,  ByVal sailprice, ByVal itemCouponyn, ByVal itemcoupontype, ByVal itemcouponvalue)
 '���ΰ�
	if sailyn="Y" then
%><br /><span class="cOr1">(<%=CLng((orgPrice-sailprice)/orgPrice*100)%>%��)<%=FormatNumber(sailprice,0) %></span>
<%		
	end if
	'������ 
	Dim discountprice
	if itemCouponyn="Y" then
		IF itemcoupontype = "1" or itemcoupontype ="2" THEN 
			IF itemcoupontype = "1" THEN ''% ����
				discountprice = CLng(orgPrice*itemcouponvalue/100)
			ELSE ''�� ����
				discountprice = itemcouponvalue
			END IF		
%><br /><span class="cBl1">(��)<%=FormatNumber(orgPrice-discountprice ,0)%></span>
<%
		END IF 
	end if 
End Sub


'//��ǰ ���ް�ǥ��
Sub SbDispItemSupplySalePrice(ByVal sailyn, ByVal orgsuplycash, ByVal sailsuplycash, ByVal itemCouponyn, ByVal itemcoupontype, ByVal couponbuyprice)
 	'���ΰ�
 	if sailyn="Y" then 
%><br /><span class="cOr1"><%=FormatNumber(sailsuplycash,0) %></span>
<%	end if  
	'������
    if itemCouponyn="Y" then
		IF itemcoupontype = "1" or itemcoupontype ="2" THEN 
			if couponbuyprice=0 or isNull(couponbuyprice) then
		%><br /><span class="cBl1"><%=FormatNumber(orgsuplycash,0)%></span>
		<%	else %><br /><span class="cBl1"><%=FormatNumber(couponbuyprice,0)%></span>
		<%	end if
		END IF 
	END IF 
End Sub 


'
'--2016 version ==================================================================================
'// ���� ī�װ�(��ϻ�ǰ) - �б⸸ �����Ұ��� //
public function getDispOnlyCategory2016(iid)
	dim SQL, i, strPrt

	SQL = "select d.catecode, i.isDefault, i.depth " &_
		"	,db_item.dbo.getCateCodeFullDepthName(d.catecode) as catename " &_
		"from db_item.dbo.tbl_display_cate as d " &_
		"	join db_item.dbo.tbl_display_cate_item as i " &_
		"		on d.catecode=i.catecode " &_
		"where i.itemid=" & iid & " " &_
		"order by i.isDefault desc, d.sortNo, i.sortNo"

	rsget.Open SQL,dbget,1

	strPrt = "<div id='dDispCate' style='margin-bottom:5px;display:;'><table id='tbl_DispCate' class='tbType1 listTb'>"
	strPrt = strPrt &"<tr>"
  strPrt = strPrt &"<th><div>�⺻</div></th>"
	strPrt = strPrt &"<th><div>ī�װ�</div></th>"
	strPrt = strPrt &"</tr>"
	if Not(rsget.EOf or rsget.BOf) then
		i = 0
		Do Until rsget.EOF
			strPrt = strPrt & "<tr onMouseOver='tbl_DispCate.clickedRowIndex=this.rowIndex'>"
			strPrt = strPrt & "<td><input type='radio' name='isDefault' value='"&rsget(0) &"' " 
			if UCase(rsget(1))="Y" then
				 strPrt = strPrt & "  checked"  
			 end if 
				strPrt = strPrt & " disabled></td>" 
			strPrt = strPrt &_
				"<td class='lt'>" & Replace(rsget(3),"^^"," >> ") &_
					"<input type='hidden' name='catecode' value='" & rsget(0) & "'>" &_
					"<input type='hidden' name='catedepth' value='" & rsget(2) & "'>" &_
				"</td>" &_
			"</tr>"
			i = i + 1
		rsget.MoveNext
		Loop
	end if
	strPrt = strPrt & "</table></div>"

	'����� ��ȯ
	getDispOnlyCategory2016 = strPrt

	rsget.Close
end Function
'// ���� ī�װ� ���� ����(��ϻ�ǰ) //
public function getDispCategory2016(iid)
	dim SQL, i, strPrt

	SQL = "select d.catecode, i.isDefault, i.depth " &_
		"	,db_item.dbo.getCateCodeFullDepthName(d.catecode) as catename " &_
		"from db_item.dbo.tbl_display_cate as d " &_
		"	join db_item.dbo.tbl_display_cate_item as i " &_
		"		on d.catecode=i.catecode " &_
		"where i.itemid=" & iid & " " &_
		"order by i.isDefault desc, d.sortNo, i.sortNo"

	rsget.Open SQL,dbget,1

	strPrt = "<div id='dDispCate' style='margin-bottom:5px;display:;'><table id='tbl_DispCate' class='tbType1 listTb'>"
	strPrt = strPrt &"<thead><tr>"
  strPrt = strPrt &"<th><div>�⺻</div></th>"
	strPrt = strPrt &"<th><div>ī�װ�</div></th>"
	strPrt = strPrt &"<th><div>����</div></th>"
	strPrt = strPrt &"</tr></thead><tbody>"
	if Not(rsget.EOf or rsget.BOf) then
		i = 0
		Do Until rsget.EOF
			strPrt = strPrt & "<tr onMouseOver='tbl_DispCate.clickedRowIndex=this.rowIndex'>"
			strPrt = strPrt & "<td><input type='radio' name='isDefault' value='"&rsget(0) &"' " 
			if UCase(rsget(1))="Y" then
				 strPrt = strPrt & "  checked" 
			 end if 
				strPrt = strPrt & " >&nbsp;&nbsp;</td>" 
			strPrt = strPrt &_
				"<td class='lt'>" & Replace(rsget(3),"^^"," >> ") &_
					"<input type='hidden' name='catecode' value='" & rsget(0) & "'>" &_
					"<input type='hidden' name='catedepth' value='" & rsget(2) & "'>" &_
				"</td>" &_
				"<td><input type='button'  value='&times' class='btn3 btnSmall'  onClick='delDispCateItem()'>&nbsp;&nbsp;</td>" &_
			"</tr>"
			i = i + 1
		rsget.MoveNext
		Loop
	end if
	strPrt = strPrt & "</tbody></table></div>"

	'����� ��ȯ
	getDispCategory2016 = strPrt

	rsget.Close
end Function

'// ���� ī�װ� ���� ����(����ǰ) //
public function getDispCategoryWait2016(iid)
	dim SQL, i, strPrt

	SQL = "select d.catecode, i.isDefault, i.depth " &_
		"	, isNull(db_item.dbo.getCateCodeFullDepthName(d.catecode),'') as catename " &_
		"from db_item.dbo.tbl_display_cate as d " &_
		"	join db_temp.dbo.tbl_display_cate_waitItem as i " &_
		"		on d.catecode=i.catecode " &_
		"where i.itemid=" & iid & " " &_
		"order by i.isDefault desc, d.sortNo, i.sortNo"

	rsget.Open SQL,dbget,1

	strPrt = "<div id='dDispCate' style='margin-bottom:5px;display:;'><table id='tbl_DispCate' class='tbType1 listTb'>"
	strPrt = strPrt &"<tr>"
  strPrt = strPrt &"<th><div>�⺻</div></th>"
	strPrt = strPrt &"<th><div>ī�װ�</div></th>"
	strPrt = strPrt &"<th><div>����</div></th>"
	strPrt = strPrt &"</tr>"
	if Not(rsget.EOf or rsget.BOf) then
		i = 0
		Do Until rsget.EOF
			strPrt = strPrt & "<tr onMouseOver='tbl_DispCate.clickedRowIndex=this.rowIndex'>"
			
				strPrt = strPrt & "<td><input type='radio' name='isDefault'  value='"&rsget(0) &"' " 
				if UCase(rsget(1))="Y" then
				 strPrt = strPrt & "  checked " 
			 end if 
				strPrt = strPrt & ">&nbsp;&nbsp;</td>" 
				 
			strPrt = strPrt &_
				"<td class='lt'>" & Replace(rsget(3),"^^"," >> ") &_
					"<input type='hidden' name='catecode' value='" & rsget(0) & "'>" &_
					"<input type='hidden' name='catedepth' value='" & rsget(2) & "'>" &_
				"</td>" &_
				"<td><input type='button'  value='&times' class='btn3 btnSmall'  onClick='delDispCateItem()'>&nbsp;&nbsp;</td>" &_
			"</tr>" 
			i = i + 1
		rsget.MoveNext
		Loop
	end if
	strPrt = strPrt & "</table></div>"

	'����� ��ȯ
	getDispCategoryWait2016 = strPrt

	rsget.Close
end Function
'--==================================================================================


'// �÷�Ĩ ���ù� �����Լ�
Function FnPASelectColorBar(icd,colSize)
	Dim oClr, tmpStr, lineCr, lp
	set oClr = new CItemColor
	oClr.FPageSize = 31
	oClr.FRectUsing = "Y"
	oClr.GetColorList
%> 
	<ul class="colorChip">
		<li id="cline0" <%if cStr(icd)="" then%>class="selected"<%end if%>> 
		<a href="javascript:selColorChip('')" onfocus="this.blur()"><img src="<%=fixImgUrl%>/web2009/common/color01_n00.gif" alt="��ü" /></a>
		</li>
<%		
	if oClr.FResultCount>0 then
		for lp=0 to oClr.FResultCount-1 
%>		
		<li id="cline<%=oClr.FItemList(lp).FcolorCode%>" <%if cStr(icd)=cStr(oClr.FItemList(lp).FcolorCode) then%>class="selected"<%end if%>> 
		<a href="javascript:selColorChip('<%=oClr.FItemList(lp).FcolorCode%>');" onfocus="this.blur()"><img src="<%=oClr.FItemList(lp).FcolorIcon%>" alt="<%=oClr.FItemList(lp).FcolorName%>" /></a> 
		</li>
<%		
			'//�౸��
'			if ((lp+1) mod colSize)=(colSize-1) and (lp+1)<oClr.FResultCount then
'				tmpStr = tmpStr & ""
'			end if
		next
	end if
	%>
	 </ul>
	 <%
	set oClr = Nothing

	'FnPASelectColorBar = tmpStr
End Function
%>