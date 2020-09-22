Attribute VB_Name = "modWindowsVersion"
Option Explicit

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion  As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Private Const VER_PLATFORM_WIN32_NT As Long = 2
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Private Const VER_PLATFORM_WIN32s As Long = 0

Private Const PRODUCT_UNLICENSED As Long = &HABCDABCD
Private Const PRODUCT_BUSINESS As Long = &H6
Private Const PRODUCT_BUSINESS_N As Long = &H10
Private Const PRODUCT_CLUSTER_SERVER As Long = &H12
Private Const PRODUCT_DATACENTER_SERVER As Long = &H8
Private Const PRODUCT_DATACENTER_SERVER_CORE As Long = &HC
Private Const PRODUCT_ENTERPRISE As Long = &H4
Private Const PRODUCT_ENTERPRISE_N As Long = &H1B
Private Const PRODUCT_ENTERPRISE_SERVER As Long = &HA
Private Const PRODUCT_ENTERPRISE_SERVER_CORE As Long = &HE
Private Const PRODUCT_ENTERPRISE_SERVER_IA64 As Long = &HF
Private Const PRODUCT_HOME_BASIC As Long = &H2
Private Const PRODUCT_HOME_BASIC_N As Long = &H5
Private Const PRODUCT_HOME_PREMIUM As Long = &H3
Private Const PRODUCT_HOME_PREMIUM_N As Long = &H1A
Private Const PRODUCT_HOME_SERVER As Long = &H13
Private Const PRODUCT_SERVER_FOR_SMALLBUSINESS As Long = &H18
Private Const PRODUCT_SMALLBUSINESS_SERVER As Long = &H9
Private Const PRODUCT_SMALLBUSINESS_SERVER_PREMIUM As Long = &H19
Private Const PRODUCT_STANDARD_SERVER As Long = &H7
Private Const PRODUCT_STANDARD_SERVER_CORE As Long = &HD
Private Const PRODUCT_STARTER As Long = &H8
Private Const PRODUCT_STORAGE_ENTERPRISE_SERVER As Long = &H17
Private Const PRODUCT_STORAGE_EXPRESS_SERVER As Long = &H14
Private Const PRODUCT_STORAGE_STANDARD_SERVER As Long = &H15
Private Const PRODUCT_STORAGE_WORKGROUP_SERVER As Long = &H16
Private Const PRODUCT_UNDEFINED As Long = &H0
Private Const PRODUCT_ULTIMATE As Long = &H1
Private Const PRODUCT_ULTIMATE_N As Long = &H1C
Private Const PRODUCT_WEB_SERVER As Long = &H11
Public Function GetWindowsVersion() As String

' VariablesDimension
Dim retOSVersionInf As OSVERSIONINFOEX
Dim retLng As Long

'Set the structure size
retOSVersionInf.dwOSVersionInfoSize = Len(retOSVersionInf)

'Get the Windows version
retLng = GetVersionEx(retOSVersionInf)

If retLng = 0 Then
    GetWindowsVersion = "Error"
    Exit Function
End If

With retOSVersionInf

If .dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
        
        Select Case .dwMajorVersion
            Case 4
                Select Case .dwMinorVersion
                    Case 0 ' Win 95
                        Select Case .szCSDVersion
                            Case "C" ' OSR2
                                GetWindowsVersion = "Windows 95 OSR2"
                            Case "B" ' OSR2
                                GetWindowsVersion = "Windows 95 OSR2"
                            Case Else
                                GetWindowsVersion = "Windows 95"
                        End Select
                    Case 10 ' Win 98
                        Select Case .szCSDVersion
                            Case "A" ' SE
                                GetWindowsVersion = "Windows 98 SE"
                            Case Else
                                GetWindowsVersion = "Windows 98"
                        End Select
                    Case 90 ' Win ME
                        GetWindowsVersion = "Windows ME"
                End Select
        End Select

ElseIf .dwPlatformId = VER_PLATFORM_WIN32_NT And .dwMajorVersion = 5 Then

    Select Case .dwMinorVersion
        Case 0 ' Win 2000
            Select Case .wProductType
                Case 1
                    Select Case .wSuiteMask
                        Case &H80 ' Data center
                            GetWindowsVersion = "Windows 2000 Data center"
                        Case &H2 ' Advanced
                            GetWindowsVersion = "Windows 2000 Advanced"
                        Case Else
                            GetWindowsVersion = "Windows 2000"
                    End Select
            End Select
        Case 1 ' Win XP
            Select Case .wProductType
                Case 1
                    Select Case .wSuiteMask
                        Case &H0 ' Pro
                            GetWindowsVersion = "Windows XP Professional"
                        Case &H200 ' Home
                            GetWindowsVersion = "Windows XP Home"
                        Case Else ' XP
                            GetWindowsVersion = "Windows XP"
                    End Select
            End Select
        Case 2 ' Win Server 2003
            Select Case .wProductType
                Case 3
                    Select Case .wSuiteMask
                        Case &H2
                            GetWindowsVersion = "Windows Server 2003 Enterprise"
                        Case &H80
                            GetWindowsVersion = "Windows Server 2003 Data center"
                        Case &H400
                            GetWindowsVersion = "Windows Server 2003 Web Edition"
                        Case &H0
                            GetWindowsVersion = "Windows Server 2003 Standard"
                        Case Else
                            GetWindowsVersion = "Windows Server 2003"
                    End Select
            End Select
    End Select

    If .wServicePackMajor > 0 Then
        GetWindowsVersion = GetWindowsVersion & " Service Pack " & .wServicePackMajor & IIf(.wServicePackMinor > 0, "." & .wServicePackMinor, vbNullString)
    End If

ElseIf .dwPlatformId = VER_PLATFORM_WIN32_NT And .dwMajorVersion = 6 Then

    Select Case .dwMinorVersion
        Case 0
                Select Case .wProductType
                    Case PRODUCT_BUSINESS
                        GetWindowsVersion = "Business Edition"
                    Case PRODUCT_BUSINESS_N
                        GetWindowsVersion = "Business Edition (N)"
                    Case PRODUCT_CLUSTER_SERVER
                        GetWindowsVersion = "Cluster Server Edition"
                    Case PRODUCT_DATACENTER_SERVER
                        GetWindowsVersion = "Server Datacenter Edition (full installation)"
                    Case PRODUCT_DATACENTER_SERVER_CORE
                        GetWindowsVersion = "Server Datacenter Edition (core installation)"
                    Case PRODUCT_ENTERPRISE
                        GetWindowsVersion = "Enterprise Edition"
                    Case PRODUCT_ENTERPRISE_N
                        GetWindowsVersion = "Enterprise Edition (N)"
                    Case PRODUCT_ENTERPRISE_SERVER
                        GetWindowsVersion = "Server Enterprise Edition (full installation)"
                    Case PRODUCT_ENTERPRISE_SERVER_CORE
                        GetWindowsVersion = "Server Enterprise Edition (core installation)"
                    Case PRODUCT_ENTERPRISE_SERVER_IA64
                        GetWindowsVersion = "Server Enterprise Edition for Itanium-based Systems"
                    Case PRODUCT_HOME_BASIC
                        GetWindowsVersion = "Home Basic Edition"
                    Case PRODUCT_HOME_BASIC_N
                        GetWindowsVersion = "Home Basic Edition (N)"
                    Case PRODUCT_HOME_PREMIUM
                        GetWindowsVersion = "Home Premium Edition"
                    Case PRODUCT_HOME_PREMIUM_N
                        GetWindowsVersion = "Home Premium Edition (N)"
                    Case PRODUCT_HOME_SERVER
                        GetWindowsVersion = "Home Server Edition"
                    Case PRODUCT_SERVER_FOR_SMALLBUSINESS
                        GetWindowsVersion = "Server for Small Business Edition"
                    Case PRODUCT_SMALLBUSINESS_SERVER
                        GetWindowsVersion = "Small Business Server"
                    Case PRODUCT_SMALLBUSINESS_SERVER_PREMIUM
                        GetWindowsVersion = "Small Business Server Premium Edition"
                    Case PRODUCT_STANDARD_SERVER
                        GetWindowsVersion = "Server Standard Edition (full installation)"
                    Case PRODUCT_STANDARD_SERVER_CORE
                        GetWindowsVersion = "Server Standard Edition (core installation)"
                    Case PRODUCT_STARTER
                        GetWindowsVersion = "Starter Edition"
                    Case PRODUCT_STORAGE_ENTERPRISE_SERVER
                        GetWindowsVersion = "Storage Server Enterprise Edition"
                    Case PRODUCT_STORAGE_EXPRESS_SERVER
                        GetWindowsVersion = "Storage Server Express Edition"
                    Case PRODUCT_STORAGE_STANDARD_SERVER
                        GetWindowsVersion = "Storage Server Standard Edition"
                    Case PRODUCT_STORAGE_WORKGROUP_SERVER
                        GetWindowsVersion = "Storage Server Workgroup Edition"
                    Case PRODUCT_ULTIMATE
                        GetWindowsVersion = "Ultimate Edition"
                    Case PRODUCT_ULTIMATE_N
                        GetWindowsVersion = "Ultimate Edition (N)"
                    Case PRODUCT_UNDEFINED
                        GetWindowsVersion = "An unknown product"
                    Case PRODUCT_UNLICENSED
                        GetWindowsVersion = "Not activated product"
                    Case PRODUCT_WEB_SERVER
                        GetWindowsVersion = "Web Server Edition"
                End Select
                
                Select Case .wProductType
                    Case 1 ' Win Vista
                        GetWindowsVersion = "Windows Vista " & GetWindowsVersion
                    Case 3 ' Win Server 2008
                        GetWindowsVersion = "Windows Server 2008 " & GetWindowsVersion
                    Case Else
                        GetWindowsVersion = "Windows Vista " & GetWindowsVersion
                End Select
                
                    If .wServicePackMajor > 0 Then
                        GetWindowsVersion = GetWindowsVersion & " Service Pack " & .wServicePackMajor & IIf(.wServicePackMinor > 0, "." & .wServicePackMinor, vbNullString)
                    End If
                    
        End Select

End If

GetWindowsVersion = GetWindowsVersion & " [Version: " & .dwMajorVersion & "." & .dwMinorVersion & "." & .dwBuildNumber & "]"

End With


End Function
