Notices - Properties - Methods
------------------------------


Properties:

Public Property Let Password(ByVal vData As String)
Public Property Get Password() As String
Public Property Let UID(ByVal vData As String)
Public Property Get UID() As String
Public Property Let ConnectionString(ByVal vData As String)
Public Property Get ConnectionString() As String

Methods:

Public Sub ErrClear()
Public Sub SaveImage(Table As String, Column As String, Where As String, Filename As String)
Public Sub ClearImage(Table As String, Column As String, Where As String)
Public Function GetASPImage(Table As String, Column As String, Where As String) As Variant
Public Sub MakeASPProxy(Filename As String, SessionVariable As String)
Public Sub MakePixel(Filename As String, Webcolor As String, Transparent As Boolean)
Public Sub DeleteFile(Filename As String)
Public Function GetImage(Table As String, Column As String, Where As String, Optional Filename As String) As Variant
^--Notice this function returns a variant data type - if using in VB you can set a picture/imagebox to equal it's
   return value to populate the picture/imagebox as such:  Set picture1.picture=GetImage(blah, blah, blah)
   If using in VB it is not required to pass a filename unless you want to actually store the file on disk as well.
   or if you only want to store it on disk, pass the filename and just call the function...
^--If using this function with ASP, you need to pass a filename of where you want to store the file, the variant returned
   will get you nothing in this case.  If you do not want to store the file on disk, use the GetASPImage method.
---Hope that all made sense!

Notices:

Known Issues/bugs/errors:

    Only designed to work with image type data fields, not OLE stored data.
    Would like to make it work with both when I get time.

    Does not work with Access Database files, was only coded to look at SQL databases.
    Has been tried with both MS SQL 6.5 and MS SQL 7.0 with the latest patches.

The end!  Hope it helps you on your database picture endeavors.......
As usuall, use this at your own risk, I am not responsible for any problems you may have. 
If you find this useful or update it please let me know - sharmonvpc@zdnetonebox.com
If you are just using it in VB, you don't have to use the dll, compile the class
file directly into your project....

Oh yeah, there are references to ado 2.1 so make sure you have it.........
Made with VB6 SP3, not tried on any other version....




