using System;
using Parsing;

class CMain
{
	static bool PASS = true;
	static void check(bool cond, string txt)
	{
		if( cond ) 
			ConsoleColor.WriteLine("{Green}OK: " + txt);
		else
		{
			ConsoleColor.WriteLine("{BrightRed}ERROR: " + txt + (char)7);
			PASS = false;
		}
	}
	[STAThread]
	static void Main()
	{
		CommandlineParser P = new CommandlineParser();
		P.caseSensitive = true;

		P.SetRequiredParam_String("name", "the person's name", "Liron Schur", "Yossi Samocha");
		P.SetRequiredSwitch_Numeric("ID","the person's ID number. this description is very!! long deliberatly, it is intended to test the usage line breaks and therefore contains too many words for one, or even two lines",0,999999999);
		P.SetRequiredSwitch_Numeric("age","the age of the person",0, 120);

		P.SetOptionalParam_String("address","home address");
		P.SetOptionalSwitch_Boolean("female","true = Female, false = Male", false);
		P.SetOptionalSwitch_Boolean("fez","true = Female, false = Male", false);
		P.SetOptionalSwitch_Boolean("Native","born here",false);
		P.SetOptionalSwitch_Numeric("Arrived","year of arrival",0, 1900, 2050);

		P.DefineSwitchGroup(1, 1, "Native", "Arrived");

		if (P.ParseString("\"Liron Schur\" /age:39 /ID:777 \"here and there\" -fem +fez +Nat"))
		{
			check(P.GetParameterAsString("name") == "Liron Schur","Name param");
			check(777 == (int)P.GetSwitchAsNumeric("ID"), "ID");
			check(39 ==  (int)P.GetSwitchAsNumeric("age"), "Age");
			check("here and there" == P.GetParameterAsString("address"), "Address");
			check( !P.GetSwitchAsBoolean("female"), "Female");
			check( P.GetSwitchAsBoolean("Native"), "Native");
			check(0 == (int)P.GetSwitchAsNumeric("Arrived"), "Arrived");

			check((string)P.GetParameterList()[0] == P.GetParameterAsString("name"), "ParamList[0]");
			check((string)P.GetParameterList()[1] == P.GetParameterAsString("address"), "ParamList[1]");
			check(P.GetParameterList().Length == 2, "ParamList.Length");

			ConsoleColor.WriteLine("{White}" + P.UsageString());
			ConsoleColor.WriteLine(PASS ? "{Green}PASS :-)" : "{BrightRed}FAIL{Default} :-(");
		}
		else
		{
			Console.WriteLine(P.UsageString());
			ConsoleColor.WriteLine("{BrightRed}FAIL{Default} :-(");
		}

		ConsoleColor.Write("{Default}");

	}
}
