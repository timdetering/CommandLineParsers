'	
' Copyright (c) 1998-1999 Microsoft Corporation.
'
'
' Script Name:  
'	Parser.VBS
' 
' Features
'	* accepts a set of options and a list of parameters 
'	* case insensitive, takes shorthanded abbriviations, handles ambigiouties 
'   * handles optional and mandatory options and parameters
'	* Generates a default usage message or use a custom one, when needed
'   * An option can be a collection
'
' Description:
' 
'	 parse command line according to the given Schema: 
'	 a list of options and parameters, seperated with blanks
'	 prefix characters:
'			"*" a collection (more then one can be specified)
'			"/" an option (default is a parameter), ("-" can also be used)
'			"+" or "-" - an option, gets True or False if no value is needed
'	 suffixes
'		   ":" option must be followed by a value
'		   "#" option must be followed by a numeric value
'
'	 an optional option should be wrapped in "[" and "]" (by default it is mandatory)
'    options are unordered, paramaters are. 
'	
'    usage is always generated if "/?" is present in the parameters
'
'	 example 1:
'		       /user: [/password:] [/*target:] [/persist] filename
'	   Will match:
'		       /UsEr:me -Target:http://microsoft.com +Persist /Tar:ftp://localhost/ums c:\
'		       -pa:*** -u:me d:\windir
'		       -u:me http://microsoft.com 
'	   will reject:
'		       /u:A /p:***  c:\  (p is ambiguous)
'		       /u me /pass:X c:\ (User must have a value)
'		       /u:A /user:B c:\  (user must be unique)
'		       /u:A /pers:No c:\ (Persist must not be folowed by a value)
'		       /Pass:X c:\       (User is mandatory)
'		       /data  70 c:\     (unknowen option)
'              /u:a				 (filename paramater is mandatory) 	
'	
'
' Modification history:
'	Written - Erez Amir (erezam) - 1999-12-12
'
' Sample usage: 
'     See bottom of this file
'---------------------------------------------------------------------->
option explicit

CONST COMMAND_LINE_PARSER_VERSION = "2.0" ' you can check this const to see if you include a parser
CONST AbortOnError = True

Class Parser
	Private m_strUsage  
	Private m_strSchema 
	Private m_oArguments
    Private m_strErrMsg				' string: accumelated error message (including CR-LF)

	
	Private Sub Class_Initialize
		'
		' Constractor event
		'
		set m_oArguments = wscript.Arguments ' default 
	end sub

	'Private Sub Class_Terminate()
	'	descructor placeholder 
	'End Sub

	private sub print(txt)
		wscript.echo txt
	end sub

	private sub Trace(txt)
		' uncomment for debug tracing
		' wscript.echo "«TRACE» " & txt
	end sub

	Function CleanParam(V)
		'
		' Removes all prefixes/suffixes from a schema element	
		'
		dim C
		C = V
		C = Replace(C, "*", " ")
		C = Replace(C, "[", " ")
		C = Replace(C, "]", " ")
		C = Replace(C, "#", " ")
		C = left(C, Instr(C & ":", ":") - 1) ' trim all before 1st ":"
		C = Replace(C, "/", " ")
		C = Replace(C, "+", " ")
		C = Replace(C, "-", " ")
		C = UCase(Trim(C))
		if instr(C, " ") then
			err.raise 2,,"Invalid schema """ & V & """"
		end if
		CleanParam = C
	end function

	Public Sub Parse
		' for backward compatability with V1.0
		ParseEx True
	end sub

	Public Function ParseEx(AbortOnError)
		' 
		' do the command line arguments parsing acording to the given schema
		'
		dim i 
		dim j 
		dim Clean				' String array : Schema name stripped from prefix/suffix chars, uppercase
		dim IsCollection		' Bool array: is this parameter a collection 
		dim IsMandatory			' Bool array: is this parameter mandatory (true) or optional
		dim MustHaveValue		' Bool array: must the param be folowed by a value?
		dim PossibleValues      ' Bool array: a "|" delimited string of possible values
		dim MustBeNumeric		' Bool array: must the value represent a numeric value	
		dim isParam				' Bool array: is it an option or a parameter
		dim Used				' Bool array: was it used at all (needed for mandatory ones)
		dim Arg					' string: argument from the command line
		dim ArgKey				' string: left  side of command line argument, before the ":"	
		dim ArgValue			' string: right side of command line argument, after  the ":"	
		dim ArgPrefix			' char: "/" | "+" | "-"
		dim Found				' integer: how many matches are there between the argument and the schema
		dim S					' string array: splitted scema 
		dim HaveOptionalParams	' boolean: any optional parameters found in the schema so far?
		dim ParamCount			' integer: index of current parameter in the schema
    
		'
		' Parse the schema
		'
		S = Split(m_strSchema)
		' setting array dimentions to be the same as S
		Clean            = S ' ucase, stripped Schema
		IsCollection     = S
		IsMandatory      = S
		MustHaveValue    = S
		MustBeNumeric    = S
		IsParam          = S
		Used		     = S
		PossibleValues   = S
		HaveOptionalParams = False
		For j = LBound(Clean) To UBound(Clean)
				Clean(j) = CleanParam(Clean(j))
				IsParam(j) = instr(S(j), "/") + instr(S(j), "-") + instr(S(j), "+") = 0
				IsCollection(j) = InStr(S(j), "*") > 0
				IsMandatory(j) =  not left(S(j),1) =  "["  or not right(S(j),1) =  "]"
				MustHaveValue(j) = (instr(S(j), ":") > 0  Or instr(S(j), "#") > 0) and not IsParam(j)
				MustBeNumeric(j) = instr(S(j), "#") > 0
				PossibleValues(j) = mid(S(j), instr(S(j) & ":", ":") + 1) ' all after 1st ":"
				PossibleValues(j) = CleanParam(PossibleValues(j))
				Used(j) = False

				trace j& ": " & Clean(j) & " Collection:" & IsCollection(j) & "  Mandatory:" & IsMandatory(j) & " MustHaveValue:" & MustHaveValue(j) & " Numeric:" & MustBeNumeric(j) & " Parameter:" & IsParam(j) & " PossibleValues:""" & PossibleValues(j) & """"

				if IsParam(j) then
					if IsMandatory(j) and HaveOptionalParams then
						' avaid "[param1] [param2] param3" which is ambiguous
						err.raise 2,,"Mandatory params cannot follow optional ones"
					end if
					if IsCollection(j) or MustHaveValue(j) then
						err.raise 2,,"A paramater cannot be a collection or take a value: " & Clean(j)
					end if
					HaveOptionalParams = not IsMandatory(j) or HaveOptionalParams
				end if
		Next
    
		'
		' parse the arguments
		'
		paramCount = 0
		m_strErrMsg = ""
		For Each Arg In m_oArguments
			j = InStr(Arg + ":", ":")
			ArgKey = Left(Arg, j - 1)
			ArgValue = Mid(Arg, j + 1)
			Trace j & " " & ArgValue
			If Left(ArgValue, 1) = """" Then ArgValue = Mid(ArgValue, 2)
			If Right(ArgValue, 1) = """" Then ArgValue = Left(ArgValue, Len(ArgValue) - 1)

			If Instr("/+-", Left(ArgKey, 1)) > 0 Then 
				'
				' it is an option
				'
				ArgPrefix = Left(ArgKey, 1)
				ArgKey = Mid(ArgKey, 2)

				If ArgKey = "?" Or ArgKey = "Help" Then
					ShowUsageAndQuit ""
				End If

				Found = 0
				For j = LBound(Clean) To UBound(Clean)
					If Left(Clean(j), Len(ArgKey)) = UCase(ArgKey) and not IsParam(j) Then
						Found = Found + 1
						If MustHaveValue(j) And ArgValue = "" Then
							m_strErrMsg = m_strErrMsg & VBCrLf & "Expected value after " & Clean(j) 
						End If
						If MustBeNumeric(j) And Not IsNumeric(ArgValue) Then
							m_strErrMsg = m_strErrMsg & VBCrLf & "Expected numeric value after " & Clean(j)
						End If
						If Not MustHaveValue(j) And ArgValue <> "" Then
							m_strErrMsg = m_strErrMsg & VBCrLf & "No value expected after " & Clean(j)
						End If
						if PossibleValues(j) <> "" and _
						   Instr("|" & PossibleValues(j) & "|", "|" & ArgValue & "|") = 0 then
							m_strErrMsg = m_strErrMsg & VBCrLf & "Possible values for " & Clean(j) & _
							                                     " are " & PossibleValues(j)
						end if

						If IsCollection(j) Then
							' Create a dictionary
							if Eval("Typename(" & Clean(j) & ")") <> "Dictionary" then
								execute "Set " & Clean(j) & " = createObject(""Scripting.Dictionary"")"
							end if

							Execute Clean(j) & ".Add ArgValue , ArgValue"
							Used(j) = true
						Else
							if Used(j) then
								m_strErrMsg = m_strErrMsg & VBCrLf & "Option " & Clean(j) & " Should be used only once"
							End If
							
							if ArgValue = "" then ArgValue = ArgPrefix = "/" or ARgPrefix = "+"

							' example: ArgValue="123" (string) and Clear(j)="X" (string)
							if MustBeNumeric(j) and IsNumeric(ArgValue) then
								Execute Clean(j) & "= " & ArgValue ' Evaluates to "X=123" which sets X as an int
							else
								Execute Clean(j) & "= ArgValue"    ' evaluates to "X=ArgValue" which sets X as a string
							end if
							Used(j) = true
						End If
					End If
				Next
				If Found = 0 Then
					m_strErrMsg = m_strErrMsg & VBCrLf & "Unknown option " & ArgKey
				End If
				If Found > 1 Then
					m_strErrMsg = m_strErrMsg & VBCrLf & "Ambiguous option " & ArgKey
				End If
            
			Else
			    ' it is a parameter, not an option (no "/" prefix)
				ParamCount = ParamCount + 1 ' this paramater number
				dim Index
				Index = 0 
				For j = LBound(Clean) To UBound(Clean)
					If IsParam(j) Then
						Index = Index + 1
						if Index = ParamCount then
							Execute Clean(j) & " = Arg"
							Used(j) = true
							exit for
						end if
					end if
				next
				if Index < ParamCount then
					m_strErrMsg = m_strErrMsg & VBCrLf & "Extra parameter: " & Arg
				end if

				if j <= UBound(PossibleValues) then
					if PossibleValues(j) <> "" and _
					   Instr("|" & PossibleValues(j) & "|", "|" & Arg & "|") = 0 then
						m_strErrMsg = m_strErrMsg & VBCrLf & "Possible values for " & Clean(j) & _
															 " are " & PossibleValues(j)
					end if
				end if
			End If

		Next
    
		' check for mandatory missing parameters/options
		For i = LBound(Clean) To UBound(Clean)
			If IsMandatory(i) And not Used(i) Then
				m_strErrMsg = m_strErrMsg & VBCrLf & "Missing: " & Clean(i)
			End If
		Next

		if m_strErrMsg <> "" and AbortOnError then ' error found
			ShowUsageAndQuit Trim(m_strErrMsg)
		end if

		ParseEx = (m_strErrMsg = "")
	End function


	Public Sub ShowUsageAndQuit(Msg)
		If Msg <> "" Then Msg = Msg & vbCrLf & vbCrLf
		if m_strUsage <> "" then
			print Msg & m_strUsage
		else
			' generate a default usage
			print Msg & "Usage: " & Wscript.ScriptName & " " & m_strSchema
		end if
		wscript.quit
	End Sub


	public Sub Dump()
		' for debugging only
		print "----- Start of parser state dump:"
		dim Clean, j, V, D
		Clean = Split(m_strSchema)
		For j = LBound(Clean) To UBound(Clean)
			V = CleanParam(Clean(j))

			If Eval( "IsObject(" & V & ")" )  Then
				dim s
				Execute "Set D = " & V
				print vbTab & Typename(D) & " " & V & " = "  
				For Each s In D
					print vbTab & vbTab & s
				Next
			Elseif Eval( "IsEmpty(" & V & ")" )  Then 
				print vbTab & "empty """ & V & """"
			else
				print vbTab & Typename(Eval(V)) & " " & V & " = " & Eval(V) 
			End If
		Next
		print "----- End of parser state dump:"
	End Sub

	public property get Schema
		Schema = m_strSchema
	end property

	public property Let Schema(strSchema)
		m_strSchema = strSchema
	end property

	public property get Usage
		Usage = m_strUsage
	end property

	public property Let Usage(strUsage)
		m_strUsage = Replace(strUsage, "%this%", Wscript.ScriptName)
	end property

	public property get GetLastError
		GetLastError = m_strErrMsg
		m_strErrMsg = ""
	end property

	public property Let Arguments(oArgs)
		'
		' This allows to override the default command-line arguments with
		' a collection of arguments, it is rarely expected to be used.
		' The main purpuse is to increase testability
		'
		set m_oArguments = oArgs
	end property

end class


'------------------------------------------------------------------------------
'                                     Sample usage
'------------------------------------------------------------------------------
'
' sample command line to try this:
' cscript Parser.vbs /user:A /pass:*** C:\ -Pers /t:Target1 /t:"2nd Target"  /count:17
const DoParserDemo = False 'True '  change to false/true to hide/show demo
if DoParserDemo then
	dim X
	' create a parser object
	set X = new parser

	' declare the wanted params
	dim Action, User, Password, Count, Target, Persist, Filename, URL
	Count = 7 ' default

	' supply the parsing scema (also used to create a default usage)
	x.Schema = "Action:ADD|REMOVE|SHOW /User: /Password: [/Count:#] [*/target:] [/persist] filename [url]"

	' supply a usage, note that "%this%" will be replaces by the script name
	' if you do not give a usage, the schema is used to produce one
	x.Usage = "Usage: %this% ADD|REMOVE|SHOW <options> filename "  & _
			VBCrLf & VbTab & "/User:<username>" & _
			VBCrLf & VbTab & "/Password:<Password>" & _
			VBCrLf & VbTab & "[/Target:<urt>] - can be repeated" & _
			VBCrLf & VbTab & "[/Persist] - a true/false switch" & _
			VBCrLf & VbTab & "[/count:number] - a numeric value expected"  & _
			VBCrLf & VbTab & "filename - a filename to use"  & _
			VBCrLf & VbTab & "[url] - an optional URL, default is http://localhost" 

	x.parse

	wscript.echo "Action=" & Action
	wscript.echo "User=" & User
	wscript.echo "Password=" & Password
	if not isEmpty(Target) then
		for Each T in Target
			wscript.echo "Target*=" & T
		next 
	end if

	wscript.echo "Persist=" & persist
	wscript.echo "Filename=" & Filename
	wscript.echo "URL=" & URL
	wscript.echo "Count=" & Count
	X.Dump ' for debug only
	wscript.quit
end if



