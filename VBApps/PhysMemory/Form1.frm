VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Port Address Read"
   ClientHeight    =   3120
   ClientLeft      =   5400
   ClientTop       =   3735
   ClientWidth     =   3570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Map"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   2640
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Antonio Giuliana, (C) 2003

Private Const OBJ_INHERIT = &H2&
Private Const OBJ_PERMANENT = &H10&
Private Const OBJ_EXCLUSIVE = &H20&
Private Const OBJ_CASE_INSENSITIVE = &H40&
Private Const OBJ_OPENIF = &H80&
Private Const OBJ_OPENLINK = &H100&
Private Const OBJ_KERNEL_HANDLE = &H200&
Private Const OBJ_VALID_ATTRIBUTES = &H3F2&

Private Const SECTION_QUERY = &H1
Private Const SECTION_MAP_WRITE = &H2
Private Const SECTION_MAP_READ = &H4
Private Const SECTION_MAP_EXECUTE = &H8

Private Const PAGE_READONLY = 2

Private Const VIEW_SHARE = 1

Private Type UNICODE_STRING
  usLength As Integer
  usMaximumLength As Integer
  usBuffer As String
End Type

Private Type OBJECT_ATTRIBUTES
    Length As Long
    RootDirectory As Long
    ObjectName As Long
    Attributes As Long
    SecurityDescriptor As Long
    SecurityQualityOfService As Long
End Type

Private Type PHYSICAL_ADDRESS
    lowpart As Long
    highpart As Long
End Type

Private Declare Function NtOpenSection _
    Lib "NTDLL.DLL" _
    (hdlSection As Long, _
     ByVal desAccess As Long, _
     objAtt As OBJECT_ATTRIBUTES) As Long
    
Private Declare Function NtMapViewOfSection _
    Lib "NTDLL.DLL" _
    (ByVal hdlSection As Long, _
     ByVal hdlProcess As Long, _
     BaseAddress As Long, _
     ByVal ZeroBits As Long, _
     ByVal CommitSize As Long, _
     SectionOffset As PHYSICAL_ADDRESS, _
     ViewSize As Long, _
     ByVal InheritDisposition As Long, _
     ByVal AllocationType As Long, _
     ByVal Protect As Long) As Long

Private Declare Function NtUnmapViewOfSection _
    Lib "NTDLL.DLL" _
    (ByVal hdlProcess As Long, _
     ByVal BaseAddress As Long) As Long

Private Declare Function CloseHandle _
    Lib "kernel32" _
    (ByVal hObject As Long) As Long

Private Declare Sub CopyMemory _
    Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, _
    Source As Any, _
    ByVal Length As Long)

Private Sub Command1_Click()

    '1'''''''''''
    Dim status As Long
    Dim ia As OBJECT_ATTRIBUTES
    Dim hdlPhysMem As Long
    Dim usDevName As UNICODE_STRING
    
    With usDevName
        .usBuffer = "\device\physicalmemory" & Chr(0)
        .usMaximumLength = Len(.usBuffer) * 2
        .usLength = .usMaximumLength - 2
    End With
    
    With ia
        .Length = Len(ia)
        .ObjectName = VarPtr(usDevName)
        .Attributes = OBJ_CASE_INSENSITIVE
        .SecurityDescriptor = 0
        .RootDirectory = 0
        .SecurityQualityOfService = 0
    End With
    
    status = NtOpenSection(hdlPhysMem, SECTION_MAP_READ, ia)
    
    'Debug.Print "NtOpenSection: "; Hex(status)
    ''''''''''''''
    
    '2''''''''''''
    Dim viewBase As PHYSICAL_ADDRESS
    Dim memVirtualAddress As Long
    Dim memLen As Long
    
    memVirtualAddress = 0
    viewBase.highpart = 0
    viewBase.lowpart = &H400
    memLen = &H10
    status = NtMapViewOfSection(hdlPhysMem, -1&, memVirtualAddress, _
                                 0&, memLen, viewBase, memLen, _
                                 VIEW_SHARE, 0&, PAGE_READONLY)
                                 
    'Debug.Print "NtMapViewOfSection: "; Hex(status)
    ''''''''''''''
    
    '3''''''''''''
    Dim phMem(0 To &HF) As Byte
    Dim ix As Integer
    
    CopyMemory phMem(0), ByVal memVirtualAddress - viewBase.lowpart + &H400, &H10
    
    List1.AddItem "Serial Port List: "
    For ix = 0 To &H7 Step 2
        List1.AddItem "COM" & ix / 2 + 1 & ": " & Hex(phMem(ix + 1) * 256& + phMem(ix))
    Next ix
    List1.AddItem ""
    
    List1.AddItem "Parallel Port List: "
    For ix = 8 To &HD Step 2
        List1.AddItem "LPT" & ix / 2 - 3 & ": " & Hex(phMem(ix + 1) * 256& + phMem(ix))
    Next ix
    ''''''''''''''
    
    '4''''''''''''
    status = NtUnmapViewOfSection(-1&, memVirtualAddress)
    
    'Debug.Print "NtUnmapViewOfSection: "; Hex(status)
    ''''''''''''''
    
    '5''''''''''''
    status = CloseHandle(hdlPhysMem)
    
    'Debug.Print "CloseHandle: "; Hex(status)
    ''''''''''''''
    
End Sub

