<%
'Instructions For Use...........
'---------------------------------------
'If xImage.ErrNumber<> 0 Then Response.Write xImage.ErrDesc & "<BR>"

'First dim your variable to use as the DBImage.cImage object
'---------------------------------------
Dim xImage

'Create your object
'---------------------------------------
Set xImage=Server.CreateObject("DBImage.cImage")



'Setup xImage's properties if using any of the database functions...
'---------------------------------------
'Setup the connection string, user id, password if you want to use the GetImage, GetASPImage, SaveImage or ClearImage Functions

'<-Notice you can use a DSN like so
'  xImage.ConnectionString="DSN=Northwind"

'<-But I would recommend using OLEDB like below
'<-It is also a good idea to set these variables in your GLOBAL.ASA file to Application Variables and call as such...
'---------------------------------------
xImage.ConnectionString="Provider=SQLOLEDB; Server=10.0.0.1; Initial Catalog=Northwind;"
xImage.UID="sa"
xImage.Password=""



'Example of make picproxy method-Used to display a picture with an asp file and your session variable(s)
'    Only really need to make once per site unless you delete it each time
'    This will create an asp file that can act as a picture with an image tag
'    View the example below of getting the picture with GetASPImage
'---------------------------------------
xImage.MakeASPProxy Server.MapPath(".") & "\picproxy.asp", "Picture"



'Example of delete file method - I know you can already do it, but this was quick and easy...
'---------------------------------------
'xImage.DeleteFile Server.MapPath(".") & "\picproxy.asp"



'Example of get image method from database field and save to disk...
'---------------------------------------
'xImage.GetImage "Categories", "Picture", "CategoryID=9", Server.MapPath(".") & "\picture9.jpg"



'Example of clear image method-Sets field to null, can be used on any field really though!
'---------------------------------------
'xImage.ClearImage "Categories", "Picture", "CategoryID=9"
'If xImage.ErrNumber<> 0 Then Response.Write xImage.ErrDesc & "<BR>"



'Example of clear internal error number
'---------------------------------------
'xImage.ErrClear



'Example of save picture to database method
'---------------------------------------
'xImage.SaveImage "Categories", "Picture", "CategoryID=9", Server.MapPath(".") & "\wmcomp.gif"



'Example of create pixel.gif function
'    You may ask what this is for, but I find it useful to size tables with a single pixel
'    and sometimes make borders etc...  OK, you can use size and style formats in your HTML but
'    not all browswers read them correctly, they WILL however read an image correctly.
'    This will create a single 1x1 pixel which you can resize till your hearts content in
'    your img tag...which will allow you to size columns, rows, tables etc or just make
'    nice little colored borders around your table items.  Don't use it if you don't like it:)
'---------------------------------------
'xImage.MakePixel Server.MapPath(".") & "\pixel.gif", "#FFFFFF", "False"



'Example of getting picture from database and using the proxy asp to view it...
'    This will set a session variable equal to the bytes needed to be a picture.
'    When used with the MakeASPProxy function it will work just fine!  You have to
'    have already created a file with the MakeASPProxy function or make the ASP file
'    on your own before you use this if you expect it to work.
'---------------------------------------
Session("Picture")=xImage.GetASPImage("Categories","Picture","CategoryID=9")
'Setup a boolean type variable (I know all variants, but we intend it to be boolean!)
'    You can use this to see if there was any errors when you run the GetASPImage function
'    or to see if the field was NULL so you display a different image if necessary!
Dim blnViewPic

If xImage.ErrNumber<> 0 Then
	Response.Write xImage.ErrDesc & "<BR>"
	blnViewPic=False
Else
	blnViewPic=True
End If

'//Set it to nothing to remove it from memory...
Set xImage=Nothing

If blnViewPic=True Then%>
<img src="picproxy.asp">
<%Else%>
<img src="nopic.jpg">
<%End If%>


<%
'Notices - Properties - Methods
'Properties:

'Public Property Let Password(ByVal vData As String)
'Public Property Get Password() As String
'Public Property Let UID(ByVal vData As String)
'Public Property Get UID() As String
'Public Property Let ConnectionString(ByVal vData As String)
'Public Property Get ConnectionString() As String

'Methods:

'Public Sub ErrClear()
'Public Sub SaveImage(Table As String, Column As String, Where As String, Filename As String)
'Public Sub ClearImage(Table As String, Column As String, Where As String)
'Public Function GetASPImage(Table As String, Column As String, Where As String) As Variant
'Public Sub MakeASPProxy(Filename As String, SessionVariable As String)
'Public Sub MakePixel(Filename As String, Webcolor As String, Transparent As Boolean)
'Public Sub DeleteFile(Filename As String)
'Public Function GetImage(Table As String, Column As String, Where As String, Optional Filename As String) As Variant
'^--Notice this function returns a variant data type - if using in VB you can set a picture/imagebox to equal it's
'   return value to populate the picture/imagebox as such:  Set picture1.picture=GetImage(blah, blah, blah)
'   If using in VB it is not required to pass a filename unless you want to actually store the file on disk as well.
'   or if you only want to store it on disk, pass the filename and just call the function...
'^--If using this function with ASP, you need to pass a filename of where you want to store the file, the variant returned
'   will get you nothing in this case.  If you do not want to store the file on disk, use the GetASPImage method.
'---Hope that all made sense!

'Notices:

'Known Issues/bugs/errors:
'
'    Only designed to work with image type data fields, not OLE stored data.
'    Would like to make it work with both when I get time.
'
'    Does not work with Access Database files, was only coded to look at SQL databases.
'    Has been tried with both MS SQL 6.5 and MS SQL 7.0 with the latest patches.
'
'The end!  Hope it helps you on your database picture endeavors.......
'As usuall, use this at your own risk, I am not responsible for any problems you may have. 
'If you find this useful or update it please let me know - sharmonvpc@zdnetonebox.com
'Oh by the way, hopefully you know how to register a dll on your system.....
'If you are just using it in VB, you don't have to use the dll, compile the class
'file directly into your project....

'Oh yeah, there are references to ado 2.1 so make sure you have it.........
%>    





