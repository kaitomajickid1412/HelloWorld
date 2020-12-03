Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.Attributes
Imports System.Windows.Media.Imaging
<Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)>
Public Class clsAppPanel
    Implements IExternalApplication
    Public Function OnStartup(application As UIControlledApplication) As Result Implements IExternalApplication.OnStartup
        application.CreateRibbonTab("RevitAPITest")
        application.CreateRibbonTab("Kaitomajickid")
        Dim Panel As RibbonPanel = application.CreateRibbonPanel("RevitAPITest", "Hello")
        Dim Panel1 As RibbonPanel = application.CreateRibbonPanel("Kaitomajickid", "Assassin's creed")
        AddButton(Panel1, "btnAss", "Assassin's creed", "AssIcon.jpg", "Ass.jpg", "clsDemDam", "Assassin's creed Valhalla", " The Brotherhood")
        AddButton(Panel1, "btnCreateEStr", "Create Columns", "CoLumns.jpg", "Ass.jpg", "Lib_Kaitomajickidvb", "Vẽ cột", " Ví dụ vẽ cột")
        AddButton(Panel, "btnHelloRiboon", "Hello World", "Hello.jpg", "Hello.jpg", "clsHelloWorld", "Ví dụ đầu tiên Hello World", "Đây là ví dụ đầu tiên về HelloWorld")
        AddPulldownButton(Panel, "pulldownComanData", "PullDown", "HelloWorld.dll", "clsCanGiua", "Hello.jpg")
        Return Result.Succeeded
    End Function
    Public Function OnShutdown(application As UIControlledApplication) As Result Implements IExternalApplication.OnShutdown
        Return Result.Succeeded
    End Function
    Public Function AddButton(ByVal panel As RibbonPanel, ByVal btnName As String, ByVal btnText As String, ByVal ImageName As String, ImageNameTooltip As String, ByVal clsName As String, ByVal TextTooltip As String, ByVal LongDescription As String) As PushButton
        Dim Source_File As String = Return_App_Path() & ImageName
        Dim Source_File_Tooltip As String = Return_App_Path() & ImageNameTooltip
        Dim pushButtonDataHello As New PushButtonData(btnName, btnText, Return_App_Path() & "HelloWorld.dll", "HelloWorld." & clsName)
        Dim pushButtonHello As PushButton = panel.AddItem(pushButtonDataHello)
        Dim help As ContextualHelp = New ContextualHelp(ContextualHelpType.Url, "https://www.autodesk.com/")
        pushButtonHello.LargeImage = New BitmapImage(New Uri(Source_File))
        pushButtonHello.ToolTip = TextTooltip
        pushButtonHello.LongDescription = LongDescription
        pushButtonHello.ToolTipImage = New BitmapImage(New Uri(Source_File_Tooltip))
        pushButtonDataHello.SetContextualHelp(help)
        Return pushButtonHello
    End Function
    Public Sub AddPushButton(ByVal Panel As RibbonPanel, ByVal PathDLL As String, ByVal NameDLL As String, ByVal ImageName As String, ByVal btnName As String, ByVal btnText As String, ByVal ToolTip As String, ByVal LongDescription As String, ByVal ImageTooltip As String)
        Dim Pathicon As String = Return_App_Path() & ImageName
        Dim PathTooltipimage As String = Return_App_Path() & ImageTooltip
        Dim pushButtonData As New PushButtonData(btnName, btnText, PathDLL, NameDLL)
        Dim pushButton As PushButton = Panel.AddItem(pushButtonData)
        pushButton.LargeImage = New BitmapImage(New Uri(Pathicon))
        pushButton.ToolTip = ToolTip
        pushButton.LongDescription = LongDescription
        pushButton.ToolTipImage = New BitmapImage(New Uri(PathTooltipimage))
    End Sub
    Function Return_App_Path() As String
        Dim AppPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location() ' chưa hiểu 
        AppPath = System.IO.Path.GetDirectoryName(AppPath)
        If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
        Return AppPath
    End Function
    Sub AddPulldownButton(ByVal panel As RibbonPanel, ByVal btnName As String, ByVal btnText As String, ByVal PathDLL As String, ByVal clsName As String, ByVal ImageName As String)
        Dim Source_File As String = Return_App_Path() & ImageName
        Dim PathDLLsp As String = Return_App_Path() & PathDLL
        'Khai báo data nút 1
        Dim pushButtonData1 As New PushButtonData(btnName & "1", btnText & "1", Return_App_Path() & "HelloWorld.dll", "HelloWorld." & clsName)
        pushButtonData1.LargeImage = New BitmapImage(New Uri(Source_File))
        'Khai báo data nút 2
        Dim pushButtonData2 As New PushButtonData(btnName & "2", btnText & "2", Return_App_Path() & "HelloWorld.dll", "HelloWorld." & clsName)
        pushButtonData2.LargeImage = New BitmapImage(New Uri(Source_File))
        'Khai báo data nút 3
        Dim pushButtonData3 As New PushButtonData(btnName & "3", btnText & "3", Return_App_Path() & "HelloWorld.dll", "HelloWorld." & clsName)
        pushButtonData3.LargeImage = New BitmapImage(New Uri(Source_File))
        'Khởi tạo PulldownButton và gán các pushButton
        Dim pulldownbtnData As New PulldownButtonData("PulldownButton", "Pulldown")
        Dim pulldownbtn As PulldownButton = panel.AddItem(pulldownbtnData)
        pulldownbtn.AddPushButton(pushButtonData1)
        pulldownbtn.AddPushButton(pushButtonData2)
        pulldownbtn.AddPushButton(pushButtonData3)

    End Sub
End Class
