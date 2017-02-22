Hi, Welcome to the wonderful world of command line parsing!
===========================================================

Files in this folder:
--------------------
`ArgsParser.cs`           -  C# version  
`CMain.cs`                -  C# version test/demo  
`parser.cls`              -  VB6 version (Doc at the top, sample usage at the bottom)  
`parser.vbs`              -  Script version (Doc at the top, sample usage at the bottom) *(Should be embedded in a WSF file, can be called from JS)*

Brief doc (Detailed doc is in the source/test code):
----------------------------------------------------
The VB6/Script versions are similar, and use a schema string:

**VB6**

    Dim p As New Parser
    p.Schema = "/Remote: /Command: [/Arguments] [/WindowState:#] p1 [p2]"
    p.usage = "bla bla"
    p.Parse
    Foo p.Argument("Remote"), p.Argument("Command"), p.Argument("Arguments"), p.Argument("WindowState")

**Script**

    set p = new parser
    .... same as above

**C#**
The C# version is more powerful, and you first declare each parameter/option on a separate call.
 
    using Parsing;
    ....
 
    CommandlineParser P = new CommandlineParser();

    P.SetRequiredSwitch_String("Remote", "the remote machine name");
    P.SetRequiredSwitch_String("Command","the command to pass");
    P.SetOptionalSwitch_Boolean("Arguments","Specify arguments?");
    P.SetOptionalSwitch_Numeric("WindowState","the display state of the window");
    P.SetRequiredParam_String("p1","First name");
    P.SetOptionalParam_String("p1","Last Name");

    if(P.Parse)
    {
        Foo(P.GetParameterAsString("Remote"));
        ... 
    }
    else
    {
        Console.WriteLine(P.UsageString());
    }
    ...