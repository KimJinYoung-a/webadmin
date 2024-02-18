<%
Class CBoardItem
	private Fid
	private Fuserid
	private Fuseremail
	private Ftitle
	private Flinkurl
	private Fcomment
	private Fimage1
	private Fimage2
	private Fimage3
	private Fregdate
	private Fhitcount
	private Fdeleteyn
	private FIIcon
	private FImgExplain1
	private FImgExplain2
	private FImgExplain3
	private FCntCount
	private FPoints
	
    '==========================================================================
	Property Get id()
		id = Fid
	end Property

	Property Get userid()
		userid = Fuserid
	end Property

	Property Get useremail()
		useremail = Fuseremail
	end Property

	Property Get title()
		title = Ftitle
	end Property

	Property Get linkurl()
		linkurl = Flinkurl
	end Property

	Property Get comment()
		comment = Fcomment
	end Property

	Property Get image1()
		image1 = Trim(Fimage1)
	end Property

	Property Get image2()
		image2 = Trim(Fimage2)
	end Property

	Property Get image3()
		image3 = Trim(Fimage3)
	end Property

	Property Get regdate()
		regdate = Fregdate
	end Property

	Property Get hitcount()
		hitcount = Fhitcount
	end Property

	Property Get deleteyn()
		deleteyn = Fdeleteyn
	end Property

	Property Get IIcon()
		IIcon = FIIcon
	end Property
	
	Property Get ImgExplain1()
		if isnull(FImgExplain1) then
			ImgExplain1 = ""
		else
			ImgExplain1 = FImgExplain1
		end if
	end Property
	
	Property Get ImgExplain2()
		if isnull(FImgExplain2) then
			ImgExplain2 = ""
		else
			ImgExplain2 = FImgExplain2
		end if
	end Property
	
	Property Get ImgExplain3()
		if isnull(FImgExplain3) then
			ImgExplain3 = ""
		else
			ImgExplain3 = FImgExplain3
		end if
	end Property
	
	Property Get CommentCount()
		if IsNumeric(FCntCount) then
			CommentCount = FCntCount
		else
			CommentCount = 0
		end if
	end Property
	
	Property Get Points()
		if IsNull(FPoints) then
			Points = 0
		else
			Points = FPoints
		end if
	end Property

    '==========================================================================
	Property Let id(byVal v)
		Fid = v
	end Property

	Property Let userid(byVal v)
		Fuserid = v
	end Property

	Property Let useremail(byVal v)
		Fuseremail = v
	end Property

	Property Let title(byVal v)
		Ftitle = v
	end Property

	Property Let linkurl(byVal v)
		Flinkurl = v
	end Property

	Property Let comment(byVal v)
		Fcomment = v
	end Property

	Property Let image1(byVal v)
		Fimage1 = v
	end Property

	Property Let image2(byVal v)
		Fimage2 = v
	end Property

	Property Let image3(byVal v)
		Fimage3 = v
	end Property

	Property Let regdate(byVal v)
		Fregdate = v
	end Property

	Property Let hitcount(byVal v)
		Fhitcount = v
	end Property

	Property Let deleteyn(byVal v)
		Fdeleteyn = v
	end Property
	
	Property Let IIcon(byVal v)
		FIIcon = v
	end Property
	
	Property Let ImgExplain1(byVal v)
		FImgExplain1 = v
	end Property
	
	Property Let ImgExplain2(byVal v)
		FImgExplain2 = v
	end Property
	
	Property Let ImgExplain3(byVal v)
		FImgExplain3 = v
	end Property
	
	Property Let CommentCount(byVal v)
		FCntCount = v
	end Property
	
	Property Let Points(byval v)
		FPoints = v
	end Property
	
    '==========================================================================
    
    public function IsImageExists()
    	IsImageExists = (image1<>"") or (image2<>"") or (image3<>"")
	end function

		
	public function GetImageCount()
		dim cnt 
		cnt=0
		if image1<>"" then cnt = cnt +1
		if image2<>"" then cnt = cnt +1
		if image3<>"" then cnt = cnt +1
		GetImageCount = cnt
	end function
	'==========================================================================
	
	Private Sub Class_Initialize()
		'
	End Sub


	Private Sub Class_Terminate()
        '
	End Sub
end Class
%>