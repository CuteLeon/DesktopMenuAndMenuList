Imports Microsoft.Win32

Public Class DesktopMenuForm

    ''' <summary>
    ''' 在桌面菜单列表里新建一个菜单项
    ''' </summary>
    ''' <param name="MenuName">菜单项的标识名称</param>
    ''' <param name="MenuText">菜单项显示的名称</param>
    ''' <param name="IconPath">菜单项的图标路径</param>
    ''' <param name="FilePath">菜单项指定的程序路径</param>
    Private Sub AddDesktopMenu(ByVal MenuName As String, ByVal MenuText As String, ByVal IconPath As String, ByVal FilePath As String)
        Dim BGShellKey As RegistryKey = Registry.ClassesRoot.OpenSubKey("Directory\Background\shell\", True)
        Dim MenuKey As RegistryKey = BGShellKey.CreateSubKey(MenuName)
        Dim CommandKey As RegistryKey = MenuKey.CreateSubKey("command")
        MenuKey.SetValue("", MenuText)
        MenuKey.SetValue("Icon", IconPath)
        MenuKey.SetValue("Position", "top") '此处不填时菜单项在菜单列表中间；top时 为顶部；bottom时 为底部
        CommandKey.SetValue("", FilePath)
    End Sub

    ''' <summary>
    ''' 删除指定标识的菜单项
    ''' </summary>
    ''' <param name="MenuName">指定的菜单项标识</param>
    Private Sub DeleteDesktopMenu(ByVal MenuName As String)
        Dim BGShellKey As RegistryKey = Registry.ClassesRoot.OpenSubKey("Directory\Background\shell\", True)
        BGShellKey.DeleteSubKeyTree(MenuName)
    End Sub

    ''' <summary>
    ''' 在桌面菜单新建一个菜单列表
    ''' </summary>
    ''' <param name="MenuName">列表的标识</param>
    ''' <param name="MenuText">列表的名称</param>
    ''' <param name="IconPath">列表的图标</param>
    ''' <param name="SubMenuName">子菜单的标识（数组）</param>
    ''' <param name="SubMenuText">子菜单的名称（数组）</param>
    ''' <param name="SubMenuIconPath">子菜单的图标（数组）</param>
    ''' <param name="SubMenuFilePath">子菜单指向的目标（数组）</param>
    Private Sub AddDesktopMenuList(ByVal MenuName As String, ByVal MenuText As String, ByVal IconPath As String,
        ByVal SubMenuName As String(), ByVal SubMenuText As String(), ByVal SubMenuIconPath As String(), ByVal SubMenuFilePath As String())
        Dim BGShellKey As RegistryKey = Registry.ClassesRoot.OpenSubKey("Directory\Background\shell\", True)
        Dim MenuKey As RegistryKey = BGShellKey.CreateSubKey(MenuName)
        Dim CommandKey As RegistryKey
        MenuKey.SetValue("MUIVerb", MenuText)
        MenuKey.SetValue("Icon", IconPath)
        MenuKey.SetValue("Position", "top") '此处不填时菜单项在菜单列表中间；top时 为顶部；bottom时 为底部
        MenuKey.SetValue("SubCommands", Strings.Join(SubMenuName, ";"))

        '32位程序在64位心痛会出现注册表重定向的问题，使用RegistryView解决重定向问题
        Dim CSShellKey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, IIf(Environment.Is64BitOperatingSystem, RegistryView.Registry64, RegistryView.Registry32))
        CSShellKey = CSShellKey.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\", True)

        For Index As Integer = LBound(SubMenuName) To UBound(SubMenuName)
            MenuKey = CSShellKey.CreateSubKey(SubMenuName(Index))
            MenuKey.SetValue("", SubMenuText(Index))
            MenuKey.SetValue("icon", SubMenuIconPath(Index))
            CommandKey = MenuKey.CreateSubKey("command")
            CommandKey.SetValue("", SubMenuFilePath(Index))
        Next
    End Sub

    ''' <summary>
    ''' 在桌面菜单删除一个菜单列表
    ''' </summary>
    ''' <param name="MenuName">要删除的额菜单列表标识</param>
    Private Sub DeleteDesktopMenuList(ByVal MenuName As String)
        Dim BGShellKey As RegistryKey = Registry.ClassesRoot.OpenSubKey("Directory\Background\shell\", True)
        Dim SubMenuName() As String = Strings.Split(BGShellKey.OpenSubKey(MenuName).GetValue("SubCommands"), ";")
        BGShellKey.DeleteSubKeyTree(MenuName)

        Dim CSShellKey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, IIf(Environment.Is64BitOperatingSystem, RegistryView.Registry64, RegistryView.Registry32))
        CSShellKey = CSShellKey.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\", True)
        For Index As Integer = LBound(SubMenuName) To UBound(SubMenuName)
            CSShellKey.DeleteSubKeyTree(SubMenuName(Index))
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        AddDesktopMenu("LeonTest", "Leon测试", "explorer.exe", "explorer.exe")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DeleteDesktopMenu("LeonTest")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        AddDesktopMenuList("LeonTestList", "Leon测试菜单组", Application.ExecutablePath,
            New String() {"MenuNotepad", "MenuCalc", "MenuCMD", "MenuExplorer"},
            New String() {"记事本", "计算器", "命令行", "我的电脑"},
            New String() {"notepad.exe", "calc.exe", "cmd.exe", "explorer.exe"},
            New String() {"notepad.exe", "calc.exe", "cmd.exe", "explorer.exe"})
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        DeleteDesktopMenuList("LeonTestList")
    End Sub
End Class
