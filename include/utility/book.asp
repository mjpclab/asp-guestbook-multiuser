<%
Dim BannerUrl,LogoUrl,BreadcrumbItems()
Sub InitHeaderData(PageName)
	if LogoBannerMode then
		BannerUrl = HomeLogo
	else
		LogoUrl = HomeLogo
	end if

	if PageName="" then
		Redim BreadcrumbItems(1,1)
	else
		Redim BreadcrumbItems(2,1)
	end if

	BreadcrumbItems(0,0)=HomeAddr
	BreadcrumbItems(0,1)=HomeName
	BreadcrumbItems(1,0)="index.asp?user=" & ruser
	BreadcrumbItems(1,1)="留言本"
	if PageName<>"" then
		BreadcrumbItems(2,0)=""
		BreadcrumbItems(2,1)=PageName
	end if
End Sub

Sub WebInitHeaderData(PageUrl1, PageName1, PageUrl2, PageName2)
	if PageName2<>"" then
		Redim BreadcrumbItems(3,1)
	elseif PageName1<>"" then
		Redim BreadcrumbItems(2,1)
	else
		Redim BreadcrumbItems(1,1)
	end if

	BreadcrumbItems(0,0)=""
	BreadcrumbItems(0,1)=web_BookName
	BreadcrumbItems(1,0)="face.asp"
	BreadcrumbItems(1,1)="首页"
	if PageName1<>"" then
		BreadcrumbItems(2,0)=PageUrl1
		BreadcrumbItems(2,1)=PageName1
	end if
	if PageName2<>"" then
		BreadcrumbItems(3,0)=PageUrl2
		BreadcrumbItems(3,1)=PageName2
	end if
End Sub
%>