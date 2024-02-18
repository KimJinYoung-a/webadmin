<%
Set cMain = New cDispCateMain
cMain.FRectCateCode = vCateCode
cMain.FRectPage = vPage
cMain.FRectStartDate = vStartDate

cMain.GetDispCateMainComment
vIdx = cMain.Fidx
vWorkComment = cMain.Fworkcomment
vRegUserID = cMain.Freguserid

vArr = cMain.GetDispCateMainDetailList
Set cMain = Nothing
IF isArray(vArr) THEN
	vMultiImg1 = vArr(5,0)
	vMultiLink1 = vArr(6,0)
	vMultiWorker = "<br>마지막작업자:" & vArr(9,0)
	vMultiImg2 = vArr(5,1)
	vMultiLink2 = vArr(6,1)
	vMultiImg3 = vArr(5,2)
	vMultiLink3 = vArr(6,2)
	vItemID1 = vArr(2,3)
	vItemImg1 = vArr(5,3)
	vItem1Worker = "<br>마지막작업자:" & vArr(9,3)
	vItemID2 = vArr(2,4)
	vItemImg2 = vArr(5,4)
	vItem2Worker = "<br>마지막작업자:" & vArr(9,4)
	vItemID3 = vArr(2,5)
	vItemImg3 = vArr(5,5)
	vItem3Worker = "<br>마지막작업자:" & vArr(9,5)
	vItemID4 = vArr(2,6)
	vItemImg4 = vArr(5,6)
	vItem4Worker = "<br>마지막작업자:" & vArr(9,6)
	vItemID5 = vArr(2,7)
	vItemImg5 = vArr(5,7)
	vItem5Worker = "<br>마지막작업자:" & vArr(9,7)
	vItemID6 = vArr(2,8)
	vItemImg6 = vArr(5,8)
	vItem6Worker = "<br>마지막작업자:" & vArr(9,8)
	vItemID7 = vArr(2,9)
	vItemImg7 = vArr(5,9)
	vItem7Worker = "<br>마지막작업자:" & vArr(9,9)
	vItemID8 = vArr(2,10)
	vItemImg8 = vArr(5,10)
	vItem8Worker = "<br>마지막작업자:" & vArr(9,10)
	vItemID9 = vArr(2,11)
	vItemImg9 = vArr(5,11)
	vItem9Worker = "<br>마지막작업자:" & vArr(9,11)
	vItemID10 = vArr(2,12)
	vItemImg10 = vArr(5,12)
	vItem10Worker = "<br>마지막작업자:" & vArr(9,12)
	vItemID11 = vArr(2,13)
	vItemImg11 = vArr(5,13)
	vItem11Worker = "<br>마지막작업자:" & vArr(9,13)
	vItemID12 = vArr(2,14)
	vItemImg12 = vArr(5,14)
	vItem12Worker = "<br>마지막작업자:" & vArr(9,14)
	vEventID1 = vArr(2,15)
	vEventImg1 = vArr(5,15)
	vEventHtml1 = chrbyte(db2html(vArr(3,15)),48,"Y") & "<br>"
	vEventHtml1 = vEventHtml1 & chrbyte(db2html(vArr(4,15)),156,"Y")
	'fnGetEventIconName(vArr(7,15))
	vEventHtml1 = Replace(vEventHtml1,chr(34),"'")
	vEvent1Worker = "<br>마지막작업자:" & vArr(9,15)
	vEventID2 = vArr(2,16)
	vEventImg2 = vArr(5,16)
	vEventHtml2 = chrbyte(db2html(vArr(3,16)),48,"Y") & "<br>"
	vEventHtml2 = vEventHtml2 & chrbyte(db2html(vArr(4,16)),156,"Y")
	'fnGetEventIconName(vArr(7,16))
	vEventHtml2 = Replace(vEventHtml2,chr(34),"'")
	vEvent2Worker = "<br>마지막작업자:" & vArr(9,16)
	vEventID3 = vArr(2,17)
	vEventImg3 = vArr(5,17)
	vEventHtml3 = chrbyte(db2html(vArr(3,17)),48,"Y") & "<br>"
	vEventHtml3 = vEventHtml3 & chrbyte(db2html(vArr(4,17)),156,"Y")
	'fnGetEventIconName(vArr(7,17))
	vEventHtml3 = Replace(vEventHtml3,chr(34),"'")
	vEvent3Worker = "<br>마지막작업자:" & vArr(9,17)
	vEventID4 = vArr(2,18)
	vEventImg4 = vArr(5,18)
	vEventHtml4 = chrbyte(db2html(vArr(3,18)),48,"Y") & "<br>"
	vEventHtml4 = vEventHtml4 & chrbyte(db2html(vArr(4,18)),156,"Y")
	'fnGetEventIconName(vArr(7,18))
	vEventHtml4 = Replace(vEventHtml4,chr(34),"'")
	vEvent4Worker = "<br>마지막작업자:" & vArr(9,18)
	vBookImg = vArr(5,19)
	vBookLink = vArr(6,19)
	vBookWorker = "<br>마지막작업자:" & vArr(9,19)
	vRecipeImg = vArr(5,20)
	vRecipeWorker = "<br>마지막작업자:" & vArr(9,20)
End If
%>