VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RubyCheckRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mNo As Long
Private mPageNumber As Long
Private mTargetText As String
Private mRubyText As String
Private mRubyPresence As String
Private mRubyFontName As String
Private mRubyFontSize As String
Private mObjectType As String
Private mNotes As String

Public Property Get No() As Long
    No = mNo
End Property

Public Property Let No(ByVal value As Long)
    mNo = value
End Property

Public Property Get PageNumber() As Long
    PageNumber = mPageNumber
End Property

Public Property Let PageNumber(ByVal value As Long)
    mPageNumber = value
End Property

Public Property Get TargetText() As String
    TargetText = mTargetText
End Property

Public Property Let TargetText(ByVal value As String)
    mTargetText = value
End Property

Public Property Get RubyText() As String
    RubyText = mRubyText
End Property

Public Property Let RubyText(ByVal value As String)
    mRubyText = value
End Property

Public Property Get RubyPresence() As String
    RubyPresence = mRubyPresence
End Property

Public Property Let RubyPresence(ByVal value As String)
    mRubyPresence = value
End Property

Public Property Get RubyFontName() As String
    RubyFontName = mRubyFontName
End Property

Public Property Let RubyFontName(ByVal value As String)
    mRubyFontName = value
End Property

Public Property Get RubyFontSize() As String
    RubyFontSize = mRubyFontSize
End Property

Public Property Let RubyFontSize(ByVal value As String)
    mRubyFontSize = value
End Property

Public Property Get ObjectType() As String
    ObjectType = mObjectType
End Property

Public Property Let ObjectType(ByVal value As String)
    mObjectType = value
End Property

Public Property Get Notes() As String
    Notes = mNotes
End Property

Public Property Let Notes(ByVal value As String)
    mNotes = value
End Property
