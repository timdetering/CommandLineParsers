VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 1998-2000 Microsoft Corporation.
'
'
' File Name:
'   Parser
'
' Features
'   * accepts a set of options and a list of parameters
'   * case insensitive, takes shorthanded abbriviations, handles ambigiouties
'   * handles optional and mandatory options and parameters
'   * Generates a default usage message or use a custom one, when needed
'   * An option can be a collection
'
' Description:
'
'    parse command line according to the given Schema:
'    a list of options and parameters, seperated with blanks
'    prefix characters:
'           "*" a collection (more then one can be specified)
'           "/" an option (default is a parameter), ("-" can also be used)
'           "+" or "-" - an option, gets True or False if no value is needed
'    suffixes
'          ":" option must be followed by a value
'          "#" option must be followed by a numeric value
'
'    an optional option should be wrapped in "[" and "]" (by default it is mandatory)
'    options are unordered, paramaters are.
'
'    usage is always generated if "/?" is present in the parameters
'
'    example 1:
'              /user: [/password:] [/*target:] [/persist] filename
'      Will match:
'              /UsEr:me -Target:http://microsoft.com +Persist /Tar:ftp://localhost/ums c:\
'              -pa:*** -u:me d:\windir
'              -u:me http://microsoft.com
'      will reject:
'              /u:A /p:***  c:\  (p is ambiguous)
'              /u me /pass:X c:\ (User must have a value)
'              /u:A /user:B c:\  (user must be unique)
'              /u:A /pers:No c:\ (Persist must not be folowed by a value)
'              /Pass:X c:\       (User is mandatory)
'              /data  70 c:\     (unknowen option)
'              /u:a              (filename paramater is mandatory)
'
'
' Modification history:
'   Written            Erez Amir (erezam) - 1999-12-12
'   Translated to VB                        2000-07-17
'
' Sample usage:
'     See bottom of this file
'---------------------------------------------------------------------->
Option Explicit

Const COMMAND_LINE_PARSER_VERSION = "1.0" ' you can check this const to see if you include a parser

Private m_strUsage As String
Private m_strSchema As String
Private m_oArguments As New Collection
Private m_oVariables As New Collection
Dim m_ErrMsg As String             ' string: accumelated error message (including CR-LF)

Private Sub PopulateArgumentsFromString(Args As String)
    Dim i As Long
    Dim InQ As Boolean
    Const Q = """"
    Do While m_oArguments.Count > 0
        m_oArguments.Remove 0
    Loop
    
    InQ = (Left(Args, 1) = Q)
    Do While Args <> ""
        i = 0
        Do While i + 1 <= Len(Args) And (Mid(Args, i + 1, 1) <> " " Or InQ)
            i = i + 1
            If Mid(Args, i, 1) = Q Then InQ = Not InQ
        Loop
        m_oArguments.Add Trim(Left(Args, i))
        Args = Trim(Mid(Args, i + 1))
    Loop
End Sub

Private Sub Class_Initialize()
    '
    ' Constractor event
    '
    PopulateArgumentsFromString Command$ ' default
End Sub

'Private Sub Class_Terminate()
'   descructor placeholder
'End Sub

Private Sub Echo(txt As String)
    MsgBox txt, vbOKOnly + vbInformation, "Command line Parser"
End Sub

Private Sub Trace(txt As String)
    ' uncomment for debug tracing
    ' echo txt
End Sub

Function CleanParam(ByVal V As String) As String
    '
    ' Removes all prefixes/suffixes from a schema element
    '
    Dim C As String
    C = V
    C = Replace(C, "*", " ")
    C = Replace(C, "[", " ")
    C = Replace(C, "]", " ")
    C = Replace(C, ":", " ")
    C = Replace(C, "#", " ")
    C = Replace(C, "/", " ")
    C = Replace(C, "+", " ")
    C = Replace(C, "-", " ")
    C = UCase(Trim(C))
    If InStr(C, " ") Then
        Err.Raise 2, , "Invalid schema """ & V & """"
    End If
    CleanParam = C
End Function

Public Function Parse(Optional AbortOnError As Boolean = True) As Boolean
    '
    ' do the command line arguments parsing acording to the given schema
    '
    Dim i As Long
    Dim j As Long
    Dim Clean() As String              ' String array : Schema name stripped from prefix/suffix chars, uppercase
    Dim IsCollection() As Boolean      ' Bool array: is this parameter a collection
    Dim IsMandatory() As Boolean       ' Bool array: is this parameter mandatory (true) or optional
    Dim MustHaveValue() As Boolean     ' Bool array: must the param be folowed by a value?
    Dim MustBeNumeric() As Boolean     ' Bool array: must the value represent a numeric value
    Dim isParam() As Boolean           ' Bool array: is it an option or a parameter
    Dim Used() As Boolean              ' Bool array: was it used at all (needed for mandatory ones)
    Dim Arg                            ' variant string : argument from the command line
    Dim ArgKey As String               ' string: left  side of command line argument, before the ":"
    Dim ArgValue As String             ' string: right side of command line argument, after  the ":"
    Dim ArgPrefix As String            ' char: "/" | "+" | "-"
    Dim Found As Long                  ' integer: how many matches are there between the argument and the schema
    Dim s() As String                  ' string array: splitted scema
    Dim HaveOptionalParams As Boolean  ' boolean: any optional parameters found in the schema so far?
    Dim ParamCount As Long             ' integer: index of current parameter in the schema

    '
    ' Parse the schema
    '
    s = Split(m_strSchema)
    i = UBound(s)
    ReDim Clean(i)
    Clean = s         ' ucase, stripped Schema
    ReDim IsCollection(i)
    ReDim IsMandatory(i)
    ReDim MustHaveValue(i)
    ReDim MustBeNumeric(i)
    ReDim isParam(i)
    ReDim Used(i)
    
    HaveOptionalParams = False
    For j = LBound(Clean) To UBound(Clean)
            Clean(j) = CleanParam(Clean(j))
            IsCollection(j) = InStr(s(j), "*") > 0
            IsMandatory(j) = Not Left(s(j), 1) = "[" Or Not Right(s(j), 1) = "]"
            MustHaveValue(j) = InStr(s(j), ":") > 0 Or InStr(s(j), "#") > 0
            MustBeNumeric(j) = InStr(s(j), "#") > 0
            isParam(j) = InStr(s(j), "/") + InStr(s(j), "-") + InStr(s(j), "+") = 0
            Used(j) = False

            Trace j & ": " & Clean(j) & " Col:" & IsCollection(j) & "  Man:" & IsMandatory(j) & " MustValue:" & MustHaveValue(j) & " num:" & MustBeNumeric(j) & " Param:" & isParam(j)
            If isParam(j) Then
                If IsMandatory(j) And HaveOptionalParams Then
                    ' avaid "[param1] [param2] param3" which is ambiguous
                    Err.Raise 2, , "Mandatory params cannot follow optional ones"
                End If
                If IsCollection(j) Or MustHaveValue(j) Then
                    Err.Raise 2, , "A paramater cannot be a collection or take a value: " & Clean(j)
                End If
                HaveOptionalParams = Not IsMandatory(j) Or HaveOptionalParams
            End If
    Next

    '
    ' parse the arguments
    '
    Do While m_oVariables.Count > 0
        m_oVariables.Remove 0
    Loop
    ParamCount = 0
    m_ErrMsg = ""
    For Each Arg In m_oArguments
        j = InStr(Arg + ":", ":")
        ArgKey = Left(Arg, j - 1)
        ArgValue = Mid(Arg, j + 1)
        If Left(ArgValue, 1) = """" Then ArgValue = Mid(ArgValue, 2)
        If Right(ArgValue, 1) = """" Then ArgValue = Left(ArgValue, Len(ArgValue) - 1)

        If InStr("/+-", Left(ArgKey, 1)) > 0 Then
            '
            ' it is an option
            '
            ArgPrefix = Left(ArgKey, 1)
            ArgKey = Mid(ArgKey, 2)

            If ArgKey = "?" Or ArgKey = "Help" Then
                ShowUsageAndQuit ""
                End
            End If

            Found = 0
            For j = LBound(Clean) To UBound(Clean)
                If Left(Clean(j), Len(ArgKey)) = UCase(ArgKey) And Not isParam(j) Then
                    Found = Found + 1
                    If MustHaveValue(j) And ArgValue = "" Then
                        m_ErrMsg = m_ErrMsg & vbCrLf & "Expected value after " & Clean(j)
                    End If
                    If MustBeNumeric(j) And Not IsNumeric(ArgValue) Then
                        m_ErrMsg = m_ErrMsg & vbCrLf & "Expected numeric value after " & Clean(j)
                    End If
                    If Not MustHaveValue(j) And ArgValue <> "" Then
                        m_ErrMsg = m_ErrMsg & vbCrLf & "No value expected after " & Clean(j)
                    End If

                    If IsCollection(j) Then
                        ' Create a new collection
                        Dim Coll As Collection
                        Set Coll = New Collection
                        On Error Resume Next
                        m_oVariables.Add Coll, Clean(j)
                        On Error GoTo 0
                        m_oVariables(Clean(j)).Add ArgValue
                        Used(j) = True
                    Else
                        If Used(j) Then
                            m_ErrMsg = m_ErrMsg & vbCrLf & "Option " & Clean(j) & " Should be used only once"
                        End If
                        
                        If ArgValue = "" Then ArgValue = ArgPrefix = "/" Or ArgPrefix = "+"

                        ' example: ArgValue="123" (string) and Clear(j)="X" (string)
                        If MustBeNumeric(j) And IsNumeric(ArgValue) Then
                            On Error Resume Next
                            m_oVariables.Add CDbl(ArgValue), Clean(j)
                            On Error Resume Next
                        Else
                            On Error Resume Next
                            m_oVariables.Add ArgValue, Clean(j)
                            On Error Resume Next
                        End If
                        Used(j) = True
                    End If
                End If
            Next
            If Found = 0 Then
                m_ErrMsg = m_ErrMsg & vbCrLf & "Unknowen option " & ArgKey
            End If
            If Found > 1 Then
                m_ErrMsg = m_ErrMsg & vbCrLf & "Ambiguous option " & ArgKey
            End If
        
        Else
            ' it is a parameter, not an option (no "/" prefix)
            ParamCount = ParamCount + 1 ' this paramater number
            Dim Index As Long
            Index = 0
            For j = LBound(Clean) To UBound(Clean)
                If isParam(j) Then
                    Index = Index + 1
                    If Index = ParamCount Then
                        On Error Resume Next
                        m_oVariables.Add Arg, Clean(j)
                        On Error Resume Next
                        Used(j) = True
                        Exit For
                    End If
                End If
            Next
            If Index < ParamCount Then
                m_ErrMsg = m_ErrMsg & vbCrLf & "Extra parameter: " & Arg
            End If
        End If

    Next

    For i = LBound(Clean) To UBound(Clean)
        If IsMandatory(i) And Not Used(i) Then
            m_ErrMsg = m_ErrMsg & vbCrLf & "Missing: " & Clean(i)
        End If
    Next

    Parse = (m_ErrMsg = "") ' did it succeed?
    
    If m_ErrMsg <> "" And AbortOnError Then ' error found
        ShowUsageAndQuit Trim(m_ErrMsg)
    End If
    
End Function


Public Sub ShowUsageAndQuit(Optional ByVal Msg As String = "")
    If Msg <> "" Then Msg = Msg & vbCrLf & vbCrLf
    If m_strUsage <> "" Then
        Echo Msg & Replace(m_strUsage, "%this%", App.EXEName)
    Else
        ' generate a default usage
        Echo Msg & "Usage: " & App.EXEName & " " & m_strSchema
    End If
    End
End Sub


Public Function Dump() As String
    ' for debugging only
    Dim Clean, j, V, D, str
    str = "----- Start of parser state dump:"
    Clean = Split(m_strSchema)
    For j = LBound(Clean) To UBound(Clean)
        V = CleanParam(Clean(j))

        If IsObject(Argument(V)) Then
            Dim s
            Set D = Argument(V)
            str = str & vbCrLf & vbTab & TypeName(D) & " " & V & " = "
            For Each s In D
                 str = str & vbCrLf & vbTab & vbTab & s
            Next
        ElseIf IsEmpty(Argument(V)) Then
            str = str & vbCrLf & vbTab & "empty """ & V & """"
        Else
            str = str & vbCrLf & vbTab & TypeName(Argument(V)) & " " & V & " = " & Argument(V)
        End If
    Next
    str = str & vbCrLf & "----- End of parser state dump:"
    Dump = str
End Function

Public Property Get Schema() As String
    Schema = m_strSchema
End Property

Public Property Let Schema(strSchema As String)
    m_strSchema = strSchema
End Property

Public Property Get usage() As String
    usage = m_strUsage
End Property

Public Property Let usage(strUsage As String)
    m_strUsage = strUsage
End Property

Public Property Let Arguments(Args As String)
    '
    ' This allows to override the default command-line arguments with
    ' a list of arguments, it is rarely expected to be used.
    ' The main purpuse is to increase testability
    '
    PopulateArgumentsFromString Args
End Property

Public Property Get Argument(ByVal Key As String)
    On Error Resume Next
    If IsObject(m_oVariables(Key)) Then
        Set Argument = m_oVariables(Key)
    Else
        Argument = m_oVariables(Key)
    End If
End Property

Public Property Get LastError() As String
    LastError = m_ErrMsg
    m_ErrMsg = ""
End Property


' sample usage
'    Dim p As New Parser
'    p.Schema = "/Remote: /Command: [/Arguments] [/WindowState:#] p1 [p2]"
'    p.usage = "bla bla"
'    p.Parse
'    Foo p.Argument("Remote"), p.Argument("Command"), p.Argument("Arguments"), p.Argument("WindowState")


