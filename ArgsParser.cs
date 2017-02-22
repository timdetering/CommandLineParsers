//
// ArgsParser - user-friendly command-line parser
//
// Created by Liron Schur (t-lirons), Microsoft R&D Haifa
// Last updated: 2003-FEB-10
//

using System;
using System.Collections;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Runtime.Serialization;

namespace Microsoft.ISA.XRay.Common.Utilities.Parsing
{
    /// <summary>these exceptions are only thrown if the programmer - not the user - makes mistakes</summary>
    [Serializable]
    public class ParseException	: Exception
    {
        public ParseException(string message): base("Program error: " + message) {}
        public ParseException(): base() {}
        public ParseException(string message, Exception e): base("Program error: " + message, e) {}
        protected ParseException(SerializationInfo info,StreamingContext context)
            : base( info, context) {}
    }

    /// <summary>The programmer defined two arguments with the same ID (examine m_lastError for details)</summary>
    [Serializable]
    public class ArgumentAlreadyDeclaredException : ParseException
    {
        public ArgumentAlreadyDeclaredException(string ID): base("Argument '" + ID + "' was already defined") {}
        public ArgumentAlreadyDeclaredException(): base() {}
        public ArgumentAlreadyDeclaredException(string message, Exception e): base("Program error: " + message, e) {}
        protected ArgumentAlreadyDeclaredException(SerializationInfo info,StreamingContext context)
            : base( info, context) {}
    }

    /// <summary>The programmer defined an argument with an empty ID name</summary>
    [Serializable]
    public class EmptyArgumentDeclaredException : ParseException
    {
        public EmptyArgumentDeclaredException(): base("You cannot define an argument with ID=\"\"") {}
        public EmptyArgumentDeclaredException(string message): base("You cannot define an argument with ID: " + message) {}
        public EmptyArgumentDeclaredException(string message, Exception e): base("Program error: " + message, e) {}
        protected EmptyArgumentDeclaredException(SerializationInfo info,StreamingContext context)
            : base( info, context) {}
    }

    /// <summary>The programmer tried to get a value of an argument that was not defined (examine m_lastError for details)</summary>
    [Serializable]
    public class NoSuchArgumentException : ParseException
    {
        public NoSuchArgumentException(string type, string ID): base("The " + type + " '" + ID + "' was not defined") {}
        public NoSuchArgumentException(string ID): base("The "  + ID + "' was not defined") {}
        public NoSuchArgumentException(): base() {}
        public NoSuchArgumentException(string message, Exception e): base("Program error: " + message, e) {}
        protected NoSuchArgumentException(SerializationInfo info,StreamingContext context)
            : base( info, context) {}
    }

    /// <summary>The programmer defined two arguments, one of them being a substring of the other</summary>
    [Serializable]
    public class AmbiguousArgumentException : ParseException
    {
        public AmbiguousArgumentException(string ID1, string ID2): base("Defined arguments '" + ID1 + "' and '" + ID2 + "' are ambiguous") {}
        public AmbiguousArgumentException(string ID1): base("Defined argument '" + ID1 +  "' are ambiguous") {}
        public AmbiguousArgumentException(): base() {}
        public AmbiguousArgumentException(string message, Exception e): base("Program error: " + message, e) {}
        protected AmbiguousArgumentException(SerializationInfo info,StreamingContext context)
            : base( info, context) {}
    }

    /// <summary>The programmer defined an optional parameter followed by a required one</summary>
    [Serializable]
    public class RequiredParamAfterOptionalParamException : ParseException
    {
        public RequiredParamAfterOptionalParamException(): base("An optional param cannot be followed by a required one") {}
        public RequiredParamAfterOptionalParamException(string message): base("Program error: " + message) {}
        public RequiredParamAfterOptionalParamException(string message, Exception e): base("Program error: " + message, e) {}
        protected RequiredParamAfterOptionalParamException(SerializationInfo info,StreamingContext context)
            : base( info, context) {}
    }

    /// <summary>The programmer defined a bad group of switches (examine m_lastError for details)</summary>
    [Serializable]
    public class BadGroupException : ParseException
    {
        public BadGroupException(string message): base(message) {}
        public BadGroupException(): base() {}
        public BadGroupException(string message, Exception e): base("Program error: " + message, e) {}
        protected BadGroupException(SerializationInfo info,StreamingContext context)
            : base( info, context) {}
    }

    /// <summary>The programmer ignored the fact that a call to Parse() failed</summary>
    [Serializable]
    public class ParseFailedException : ParseException
    {
        public ParseFailedException(string message): base(message) {}
        public ParseFailedException(): base() {}
        public ParseFailedException(string message, Exception e): base("Program error: " + message, e) {}
        protected ParseFailedException(SerializationInfo info,StreamingContext context)
            : base( info, context) {}
    }


    /// <summary>abstract container class for arguments (both switches and parameters)</summary>
    abstract internal class CArgument
    {
        public CArgument(string id, string desc, bool fIsOptional)
        {
            m_id = id;
            m_description = desc;
            m_fIsOptional = fIsOptional;
        }

        public string Id
        {
            get {return m_id;}
        }

        public object GetValue() {return m_val;}
        
        // always accepts string from commandline
        abstract public bool SetValue(string val);
        
        abstract public string possibleValues();
        
        public string description
        {
            get
            {
                if (m_description.Length == 0) {return m_id;}
                else {return m_description;}
            }
            set
            {
                m_description = value;
            }
        }

        public bool			isOptional
        {
            get {return m_fIsOptional;}
        }

        public bool			isAssigned
        {
            get {return m_fIsAssigned;}
        }

        protected object	m_val = "";
        protected bool		m_fIsAssigned = false;

        private string		m_id = "";
        private string		m_description = "";
        private bool		m_fIsOptional = true;
    }

    /// <summary>container class for numeric arguments (both switches and parameters)</summary>
    internal class CNumericArgument : CArgument
    {
        public CNumericArgument(
            string id,
            string desc,
            bool fIsOptional,
            double defVal,
            double minRange,
            double maxRange
            ): base(id, desc, fIsOptional)
        {
            m_val = defVal;
            m_minRange = minRange;
            m_maxRange = maxRange;
        }

        public override bool SetValue(string val)
        {
            if (isAssigned)
            {
                // ISSUE: reassign?
            }
            m_fIsAssigned = true;

            try
            {
                if (val.ToLower().StartsWith("0x"))
                {
                    m_val = (double)int.Parse(val.Substring(2), NumberStyles.HexNumber);
                }
                else
                {
                    m_val = double.Parse (val, NumberStyles.Any);
                }
            }
            catch (ArgumentNullException)
            {
                return false;
            }
            catch (FormatException)
            {
                return false;
            }

            catch (OverflowException)
            {
                return false;
            }

            return ((double)m_val >= m_minRange && (double)m_val <= m_maxRange);
        }

        public override string possibleValues()
        {
            return "between " + m_minRange + " and " + m_maxRange;
        }

        private double m_minRange = double.MinValue;
        private double m_maxRange = double.MaxValue;
    }

    /// <summary>container class for string arguments (both switches and parameters)</summary>
    internal class CStringArgument : CArgument
    {
        public CStringArgument(
            string id,
            string desc,
            bool fIsOptional,
            string defVal,
            bool fIsPossibleValsCaseSensitive,
            params string[] possibleVals
            ): base(id, desc, fIsOptional)
        {
            m_possibleVals = possibleVals;
            m_val = defVal;
            m_fIsPossibleValsCaseSensitive = fIsPossibleValsCaseSensitive;
        }

        public override bool SetValue(string val)
        {
            if (isAssigned)
            {
                // ISSUE: reassign?
            }
            m_fIsAssigned = true;

            m_val = val;

            if (m_possibleVals.Length == 0)
            {
                // empty list: free text allowed
                return true;
            }

            foreach(string V in m_possibleVals)
            {
                if ((string)m_val ==  V ||
                    (!m_fIsPossibleValsCaseSensitive && ((string)m_val).ToLower() ==  V.ToLower()))
                {
                    return true;
                }
            }
            
            return false;
        }

        public override string possibleValues()
        {
            if (m_possibleVals.Length == 0)
            {
                return "free text";
            }
            else
            {
                string str = "{" + (m_fIsPossibleValsCaseSensitive ? m_possibleVals[0] : m_possibleVals[0].ToLower());
                for (int i = 1; i < m_possibleVals.Length; i++)
                {
                    str += "|" + (m_fIsPossibleValsCaseSensitive ? m_possibleVals[i] : m_possibleVals[i].ToLower());
                }
                str += "}";
                return str;
            }
        }

        private string[] m_possibleVals = {""};
        private bool m_fIsPossibleValsCaseSensitive = true;
    }
    
    /// <summary>container class for boolean switches</summary>
    internal class CBooleanArgument : CArgument
    {
        public CBooleanArgument(
            string id,
            string desc,
            bool fIsOptional,
            bool defVal
            ): base(id, desc, fIsOptional)
        {
            m_val = defVal;
        }

        public override bool SetValue(string token)
        {
            if (isAssigned)
            {
                // ISSUE: reassign?
            }
            m_fIsAssigned = true;

            m_val = ( token != "-" );
            return true;
        }

        public override string possibleValues()
        {
            return "precede by [+] or [-]";
        }
    }
    

    /// <summary>contains information about dependecies between switches (cannot/must appear together etc.)</summary>
    [CLSCompliantAttribute(false)]
    internal class CArgGroups
    {
        public CArgGroups(uint min, uint max, params string[] args)
        {
            m_minAppear = min;
            m_maxAppear = max;
            m_args = args;
        }
                
        
        public uint m_minAppear;
        public uint	m_maxAppear;
        public bool InRange(uint num)
        {
            return (num >= m_minAppear && num <= m_maxAppear);
        }

        
        public string[] Args
        {
            get {return m_args;}
        }
        public string ArgList()
        {
            string list = "{";
            foreach(string A in Args)
            {
                list += "," + A;
            }
            list = list.Replace("{,","{") + "}";


            return list;
        }

        public string RangeDescription()
        {
            if (m_minAppear == 1 && m_maxAppear == 1)
                return "one of the switches " + ArgList() + " must be used exclusively";

            else if (m_minAppear == 1 && m_maxAppear == Args.Length)
                return "one or more of the switches " + ArgList() + " must be used";

            else if (m_minAppear == 1 && m_maxAppear > 1)
                return "one (but not more than " + + m_maxAppear + ") of the switches " + ArgList() + " must be used";

            else if (m_minAppear == 0 && m_maxAppear == 1)
                return "only one of the switches " + ArgList() + " can be used";

            else if (m_minAppear == 0 && m_maxAppear > 1)
                return "only " + m_maxAppear + " of the switches " + ArgList() + " can be used";

            else return "between " + m_minAppear + " and " + m_maxAppear + " of the switches " + ArgList() + " must be used";
        }

        private string[]	m_args;
    }
    
    
    /// <summary>contains methods to define arguments, parse a program's command-line and retrieve the results</summary>
    /// <remarks>this is the only class that is exposed to the programmer using this module</remarks>
    public class CommandlineParser
    {
        /// <summary>use default syntax of "/numericArg:[n] /stringArg:[x|"x"] [/|+|-]booleanArg"</summary>
        public CommandlineParser()
        {
            BuildRegularExpression();
        }

        /// <summary>use syntax with your own switch and delimiter</summary>
        /// <param name="yourOwnSwitch">override the default '/' switch character (NOTE: '-' is still 'False' for boolean switches)</param>
        /// <param name="yourOwnDelim">override the default ':' delimiter</param>
        public CommandlineParser(char yourOwnSwitch, char yourOwnDelim)
        {
            m_switchChar = yourOwnSwitch;
            m_delimChar = yourOwnDelim;
            BuildRegularExpression();
        }


        /// <summary>Declare an optional numeric command-line switch - order is unimportant</summary>
        /// <param name="id">the switch (e.g. "name" in "/name=[x]")</param>
        /// <param name="description">a full description of the switch</param>
        /// <param name="defaultVal">the switch's default value</param>
        /// <param name="minRange">minimum value of the switch</param>
        /// <param name="maxRange">maximum value of the switch</param>
        public void SetOptionalSwitch_Numeric(string id, string description, double defaultVal, double minRange, double maxRange)
        {
            DeclareNumericSwitch(id, description, true, defaultVal, minRange, maxRange);
        }

        /// <summary>Declare an optional numeric command-line switch - order is unimportant</summary>
        /// <param name="id">the switch (e.g. "name" in "/name=[x]")</param>
        /// <param name="description">a full description of the switch</param>
        /// <param name="defaultVal">the switch's default value</param>
        public void SetOptionalSwitch_Numeric(string id, string description, double defaultVal)
        {
            DeclareNumericSwitch(id, description, true, defaultVal, int.MinValue, int.MaxValue);
        }

        /// <summary>Declare a Required command-line switch - order is unimportant</summary>
        /// <param name="id">the switch (e.g. "name" in "/name=[x]")</param>
        /// <param name="description">a full description of the switch</param>
        /// <param name="minRange">minimum value of the switch</param>
        /// <param name="maxRange">maximum value of the switch</param>
        public void SetRequiredSwitch_Numeric(string id, string description, double minRange, double maxRange)
        {
            DeclareNumericSwitch(id, description, false, 0, minRange, maxRange);
        }

        /// <summary>Declare a Required command-line switch - order is unimportant</summary>
        /// <param name="id">the switch (e.g. "name" in "/name=[x]")</param>
        /// <param name="description">a full description of the switch</param>
        public void SetRequiredSwitch_Numeric(string id, string description)
        {
            DeclareNumericSwitch(id, description, false, 0, int.MinValue, int.MaxValue);
        }

        /// <summary>Declare an optional string command-line switch - order is unimportant</summary>
        /// <param name="id">the switch (e.g. "name" in "/name=[x]")</param>
        /// <param name="description">a full description of the switch</param>
        /// <param name="defaultVal">the switch's default value</param>
        /// <param name="possibleVals">list of possible values (or empty for free text)</param>
        public void SetOptionalSwitch_String(string id, string description, string defaultVal, params string[] possibleVals)
        {
            DeclareStringSwitch(id, description, true, defaultVal, true, possibleVals);
        }

        /// <summary>Declare an optional string command-line switch - order is unimportant</summary>
        /// <param name="id">the switch (e.g. "name" in "/name=[x]")</param>
        /// <param name="description">a full description of the switch</param>
        /// <param name="defaultVal">the switch's default value</param>
        /// <param name="fIsPossibleValsCaseSensitive">should possible string values be case-sensitive</param>
        /// <param name="possibleVals">list of possible values (or empty for free text)</param>
        public void SetOptionalSwitch_String(string id, string description, string defaultVal, bool fIsPossibleValsCaseSensitive, params string[] possibleVals)
        {
            DeclareStringSwitch(id, description, true, defaultVal, fIsPossibleValsCaseSensitive, possibleVals);
        }

        /// <summary>Declare an optional string command-line switch - order is unimportant</summary>
        /// <param name="id">the switch (e.g. "name" in "/name=[x]")</param>
        /// <param name="description">a full description of the switch</param>
        public void SetOptionalSwitch_String(string id, string description)
        {
            DeclareStringSwitch(id, description, true, "", true);
        }

        /// <summary>Declare a Required string command-line switch - order is unimportant</summary>
        /// <param name="id">the switch (e.g. "name" in "/name=[x]")</param>
        /// <param name="description">a full description of the switch</param>
        /// <param name="possibleVals">list of possible values (or empty for free text)</param>
        public void SetRequiredSwitch_String(string id, string description, params string[] possibleVals)
        {
            DeclareStringSwitch(id, description, false, "", true, possibleVals);
        }
        
        /// <summary>Declare a Required string command-line switch - order is unimportant</summary>
        /// <param name="id">the switch (e.g. "name" in "/name=[x]")</param>
        /// <param name="description">a full description of the switch</param>
        /// <param name="fIsPossibleValsCaseSensitive">should possible string values be case-sensitive</param>
        /// <param name="possibleVals">list of possible values (or empty for free text)</param>
        public void SetRequiredSwitch_String(string id, string description, bool fIsPossibleValsCaseSensitive, params string[] possibleVals)
        {
            DeclareStringSwitch(id, description, false, "", fIsPossibleValsCaseSensitive, possibleVals);
        }
        
        /// <summary>Declare a Required string command-line switch - order is unimportant</summary>
        /// <param name="id">the switch (e.g. "name" in "/name=[x]")</param>
        /// <param name="description">a full description of the switch</param>
        public void SetRequiredSwitch_String(string id, string description)
        {
            DeclareStringSwitch(id, description, false, "", true);
        }

        /// <summary>Declare an optional boolean command-line switch - order is unimportant</summary>
        /// <remarks>"/k" and "+k" are True; "-k" is False</remarks>
        /// <param name="id">the switch (e.g. "name" in "/name=[x]")</param>
        /// <param name="description">a full description of the switch</param>
        /// <param name="defaultVal">the switch's default value</param>
        public void SetOptionalSwitch_Boolean(string id, string description, bool defaultVal)
        {
            DeclareBooleanSwitch(id, description, true, defaultVal);
        }


        /// <summary>Declare an optional numeric command-line parameter - order is important!</summary>
        /// <remarks>Optional parameter cannot be followed by a Required one</remarks>
        /// <param name="id">the parameter name (for reference)</param>
        /// <param name="description">a full description of the parameter</param>
        /// <param name="defaultVal">the parameter's default value</param>
        /// <param name="minRange">minimum value of the parameter</param>
        /// <param name="maxRange">maximum value of the parameter</param>
        public void SetOptionalParam_Numeric(string id, string description, double defaultVal, double minRange, double maxRange)
        {
            DeclareParam_Numeric(id, description, true, defaultVal, minRange, maxRange);
        }

        /// <summary>Declare an optional numeric command-line parameter - order is important!</summary>
        /// <remarks>Optional parameter cannot be followed by a Required one</remarks>
        /// <param name="id">the parameter name (for reference)</param>
        /// <param name="description">a full description of the parameter</param>
        /// <param name="defaultVal">the parameter's default value</param>
        public void SetOptionalParam_Numeric(string id, string description, double defaultVal)
        {
            DeclareParam_Numeric(id, description, true, defaultVal, int.MinValue, int.MaxValue);
        }

        /// <summary>Declare a required numeric command-line parameter - order is important!</summary>
        /// <param name="id">the parameter name (for reference)</param>
        /// <param name="description">a full description of the parameter</param>
        /// <param name="minRange">minimum value of the parameter</param>
        /// <param name="maxRange">maximum value of the parameter</param>
        public void SetRequiredParam_Numeric(string id, string description, double minRange, double maxRange)
        {
            DeclareParam_Numeric(id, description, false, 0, minRange, maxRange);
        }

        /// <summary>Declare a required numeric command-line parameter - order is important!</summary>
        /// <param name="id">the parameter name (for reference)</param>
        /// <param name="description">a full description of the parameter</param>
        public void SetRequiredParam_Numeric(string id, string description)
        {
            DeclareParam_Numeric(id, description, false, 0, int.MinValue, int.MaxValue);
        }

        /// <summary>Declare an optional string command-line parameter - order is important!</summary>
        /// <remarks>Optional parameter cannot be followed by a Required one</remarks>
        /// <param name="id">the parameter name (for reference)</param>
        /// <param name="description">a full description of the parameter</param>
        /// <param name="defaultVal">the parameter's default value</param>
        /// <param name="possibleVals">possible values for this parameter (or empty for free text)</param>
        public void SetOptionalParam_String(string id, string description, string defaultVal, params string[] possibleVals)
        {
            DeclareStringParam(id, description, true, defaultVal, true, possibleVals);
        }

        /// <summary>Declare an optional string command-line parameter - order is important!</summary>
        /// <remarks>Optional parameter cannot be followed by a Required one</remarks>
        /// <param name="id">the parameter name (for reference)</param>
        /// <param name="description">a full description of the parameter</param>
        /// <param name="defaultVal">the parameter's default value</param>
        /// <param name="fIsPossibleValsCaseSensitive">should possible string values be case-sensitive</param>
        /// <param name="possibleVals">possible values for this parameter (or empty for free text)</param>
        public void SetOptionalParam_String(string id, string description, string defaultVal, bool fIsPossibleValsCaseSensitive, params string[] possibleVals)
        {
            DeclareStringParam(id, description, true, defaultVal, fIsPossibleValsCaseSensitive, possibleVals);
        }

        /// <summary>Declare an optional string command-line parameter - order is important!</summary>
        /// <remarks>Optional parameter cannot be followed by a Required one</remarks>
        /// <param name="id">the parameter name (for reference)</param>
        /// <param name="description">a full description of the parameter</param>
        public void SetOptionalParam_String(string id, string description)		
        {
            DeclareStringParam(id, description, true, "", true);
        }

        /// <summary>Declare a required string command-line parameter - order is important!</summary>
        /// <param name="id">the parameter name (for reference)</param>
        /// <param name="description">a full description of the parameter</param>
        /// <param name="possibleVals">possible values for this parameter (or empty for free text)</param>
        public void SetRequiredParam_String(string id, string description, params string[] possibleVals)
        {
            DeclareStringParam(id, description, false, "", true, possibleVals);
        }

        /// <summary>Declare a required string command-line parameter - order is important!</summary>
        /// <param name="id">the parameter name (for reference)</param>
        /// <param name="description">a full description of the parameter</param>
        /// <param name="fIsPossibleValsCaseSensitive">should possible string values be case-sensitive</param>
        /// <param name="possibleVals">possible values for this parameter (or empty for free text)</param>
        public void SetRequiredParam_String(string id, string description, bool fIsPossibleValsCaseSensitive, params string[] possibleVals)
        {
            DeclareStringParam(id, description, false, "", fIsPossibleValsCaseSensitive, possibleVals);
        }

        /// <summary>Declare a required string command-line parameter - order is important!</summary>
        /// <param name="id">the parameter name (for reference)</param>
        /// <param name="description">a full description of the parameter</param>
        public void SetRequiredParam_String(string id, string description)
        {
            DeclareStringParam(id, description, false, "", true);
        }

        
        /// <summary>parse the program's Commandline</summary>
        /// <returns>was parsing successful?</returns>
        public bool ParseCommandLine()
        {
            SetFirstArgumentAsAppName();
            return ParseString(Environment.CommandLine);
        }

        /// <summary>parse a string with a list of space-seperated arguments (argument can be "wrapped in inverted commas")</summary>
        /// <param name="argumentsLine">the line of arguments</param>
        /// <param name="fIsFirstArgTheAppName">is the first argument the application's name?</param>
        /// <returns>did the parsing succeed?</returns>
        /// <example>ParseString("PROG.EXE file1 /s /q /type:"binary" -P", true)</example>
        public bool ParseString(string argumentsLine, bool fIsFirstArgTheAppName)
        {
            if (fIsFirstArgTheAppName)
            {
                SetFirstArgumentAsAppName();
            }

            return ParseString(argumentsLine);
        }

        /// <summary>parse a string with a list of space-seperated arguments (argument can be "wrapped in inverted commas")</summary>
        /// <param name="argumentsLine">the line of arguments WITHOUT THE APPLICATION NAME</param>
        /// <returns>did the parsing succeed?</returns>
        public bool ParseString(string argumentsLine)
        {
            if (m_parseSuccess)
            {
                throw new ParseFailedException("You cannot parse twice!");
            }

            // add "/?" as an optional switch
            SetOptionalSwitch_Boolean("?", "Displays this usage string", false);

            int paramIndex = 0;

            // add a white space at the end (for the regular expression)
            argumentsLine = argumentsLine.TrimStart() + " ";
            
            Regex Reg = new Regex(m_Syntax);
            
            Match argMatch = Reg.Match(argumentsLine);

            while(argMatch.Success)
            {
                string swtch	= argMatch.Result("${" + SWITCH_TOKEN + "}");
                string ID		= argMatch.Result("${" + ID_TOKEN + "}");
                string delim	= argMatch.Result("${" + DELIM_TOKEN + "}");
                string val		= argMatch.Result("${" + VALUE_TOKEN + "}");
                val = val.TrimEnd();
                if (val.StartsWith("\"") && val.EndsWith("\""))
                {
                    val = val.Substring(1, val.Length - 2);
                }

                if (ID.Length == 0)
                {
                    // only value? it's a parameter!
                    if (!InputParam(val, paramIndex++))
                    {
                        return false;
                    }
                }
                else
                {
                    // if "/?" (Usage) switch was used, return "Parse failed"
                    if (ID == "?")
                    {
                        m_lastError = "Usage Info requested";
                        m_parseSuccess = false;
                        return false;
                    }

                    // has an ID? then it's a switch!
                    if (!InputSwitch(swtch, ID, delim, val))
                    {
                        return false;
                    }
                }

                argMatch = argMatch.NextMatch();
            }

            foreach(CArgument A in m_declaredSwitches)
            {
                if (!A.isOptional && !A.isAssigned)
                {
                    m_lastError = "Required switch '" + A.Id + "' was not assigned a value";
                    return false;
                }
            }
            foreach(CArgument A in m_declaredParams)
            {
                if (!A.isOptional && !A.isAssigned)
                {
                    m_lastError = "Required parameter '" + A.Id + "' was not assigned a value";
                    return false;
                }
            }

            m_parseSuccess = IsGroupRulesKept();
            return m_parseSuccess;
        }


        /// <summary>Get a switch value after parsing</summary>
        /// <param name="id">the switch id</param>
        /// <returns>switch value (use GetSwitchAsXXX for easy type casts)</returns>
        public object GetSwitch(string id)
        {
            if (!m_parseSuccess)
            {
                throw new ParseFailedException(lastError);
            }

            if (id == APPLICATION_NAME)
            {
                throw new ParseException(APPLICATION_NAME + " is a reserved internal id and must not be used");
            }

            CArgument arg = FindExactArg(id, m_declaredSwitches);
            
            if (arg == null)
            {
                throw new NoSuchArgumentException("switch", id);
            }

            return arg.GetValue();
        }

        /// <summary>Get a numeric switch value after parsing</summary>
        /// <param name="id">the switch id</param>
        /// <returns>switch value type-cast to double</returns>
        public double GetSwitchAsNumeric(string id)
        {
            return (double)GetSwitch(id);
        }
        /// <summary>Get a string switch value after parsing</summary>
        /// <param name="id">the switch id</param>
        /// <returns>switch value type-cast to string</returns>
        public string GetSwitchAsString(string id)
        {
            return (string)GetSwitch(id);
        }
        /// <summary>Get a boolean switch value after parsing</summary>
        /// <param name="id">the switch id</param>
        /// <returns>switch value type-cast to bool</returns>
        public bool GetSwitchAsBoolean(string id)
        {
            return (bool)GetSwitch(id);
        }
                
        /// <summary>Was this switch given a value in the parsed line?</summary>
        /// <param name="id">the switch id</param>
        /// <returns>true if the switch appeared in the parsed line</returns>
        public bool IsAssignedSwitch(string id)
        {
            if (!m_parseSuccess)
            {
                throw new ParseFailedException(lastError);
            }

            if (id == APPLICATION_NAME)
            {
                throw new ParseException(APPLICATION_NAME + " is a reserved internal id and must not be used");
            }

            CArgument arg = FindExactArg(id, m_declaredSwitches);
            
            if (arg == null)
            {
                throw new NoSuchArgumentException("switch", id);
            }

            return arg.isAssigned;
        }

        
        /// <summary>Get a parameter value after parsing</summary>
        /// <param name="id">the parameter id</param>
        /// <returns>parameter value (use GetparameterAsXXX for easy type casts)</returns>
        public object GetParameter(string id)
        {
            if (!m_parseSuccess)
            {
                throw new ParseFailedException(lastError);
            }

            if (id == APPLICATION_NAME)
            {
                throw new ParseException(APPLICATION_NAME + " is a reserved internal id and must not be used");
            }

            CArgument arg = FindExactArg(id, m_declaredParams);
            
            if (arg == null)
            {
                throw new NoSuchArgumentException("parameter", id);
            }

            return arg.GetValue();
        }

        /// <summary>Get a numeric parameter value after parsing</summary>
        /// <param name="id">the parameter id</param>
        /// <returns>parameter value type-cast to double</returns>
        public double GetParameterAsNumeric(string id)
        {
            return (double)GetParameter(id);
        }
        /// <summary>Get a string parameter value after parsing</summary>
        /// <param name="id">the parameter id</param>
        /// <returns>parameter value type-cast to string</returns>
        public string GetParameterAsString(string id)
        {
            return (string)GetParameter(id);
        }
                
        /// <summary>Was this parameter given a value in the parsed line?</summary>
        /// <param name="id">the parameter id</param>
        /// <returns>true if the switch appeared in the parsed line</returns>
        public bool IsAssignedParameter(string id)
        {
            if (!m_parseSuccess)
            {
                throw new ParseFailedException(lastError);
            }

            if (id == APPLICATION_NAME)
            {
                throw new ParseException(APPLICATION_NAME + " is a reserved internal id and must not be used");
            }

            CArgument arg = FindExactArg(id, m_declaredParams);
            
            if (arg == null)
            {
                throw new NoSuchArgumentException("parameter", id);
            }

            return arg.isAssigned;
        }

        
        /// <summary>Returns the parameter list after parsing</summary>
        /// <returns>an array of the parsed parameters, in order</returns>
        public object[] GetParameterList()
        {
            int firstParamIndex = (IsFirstArgumentAppName() ? 1 : 0);
            
            if (m_declaredParams.Count == firstParamIndex)
            {
                return null;
            }

            object [] list = new object[m_declaredParams.Count - firstParamIndex];
            for (int i = firstParamIndex; i < m_declaredParams.Count; i++)
            {
                list[i - firstParamIndex] = ((CArgument)m_declaredParams[i]).GetValue();
            }

            return list;
        }

        /// <summary>
        /// GetSwitchesList - Returns the switch list after parsing</summary>
        /// <returns>two-dim array of the parsed parameters: ParamID | ParamValue</returns>
        public Array GetSwitchesList()
        {
            Array list = Array.CreateInstance(typeof(object),m_declaredSwitches.Count,2);

            for (int i = 0; i < m_declaredSwitches.Count; i++)
            {
                list.SetValue(((CArgument)m_declaredSwitches[i]).Id,i,1);
                list.SetValue(((CArgument)m_declaredSwitches[i]).GetValue(),i,0);
            }

            return list;
        }

        
        /// <summary>Declare argument flag aliases (e.g. "/s" treated as "/size")</summary>
        /// <param name="alias">the alias (shortcut) name</param>
        /// <param name="treatedAs">replace the alias with this ID</param>
        /// <remarks>aliases may be ambiguous with real names, e.g. "/s" treated as "/size" and not "/source"</remarks>
        public void SetAlias(string alias, string treatedAs)
        {
            if (alias != treatedAs)
            {
                m_aliases[alias] = treatedAs;
            }
        }
    

        /// <summary>define relations between switches (must/can appear together etc.)</summary>
        /// <param name="minAppear">minimum number of switches appearing from this group</param>
        /// <param name="maxAppear">maximum number of switches appearing from this group</param>
        /// <param name="IDs">list of two or more switch IDs</param>
        [CLSCompliantAttribute(false)]
        public void DefineSwitchGroup(uint minAppear, uint maxAppear, params string[] IDs)
        {
            if (IDs.Length < 2 || maxAppear < minAppear || maxAppear == 0)
            {
                throw new BadGroupException("A group must have at least two members");
            }

            if (minAppear == 0 && maxAppear == IDs.Length)
            {
                // waste of time - but we'll be forgiving...?
                return;
            }

            if (minAppear > IDs.Length)
            {
                throw new BadGroupException("You cannot have " + minAppear + " appearance(s) in a group of " + IDs.Length + " switch(es)!");
            }

            foreach(string id in IDs)
            {
                if (FindExactArg(id, m_declaredSwitches) == null)
                {
                    throw new NoSuchArgumentException("switch", id);
                }
            }

            CArgGroups G = new CArgGroups(minAppear, maxAppear, IDs);
            m_argGroups.Add(G);

            if (m_usageGroups.Length == 0)
            {
                m_usageGroups = "NOTES:" + Environment.NewLine;
            }

            m_usageGroups += " - " + G.RangeDescription() + Environment.NewLine;
        }
        

        /// <summary>returns an automatically generated usage string (plus lastError if not OK)</summary>
        /// <returns>a multi-line text description</returns>
        public string UsageString()
        {
            string str = "";
            
            if (m_lastError.Length != 0)
            {
                str = ">> " + m_lastError + Environment.NewLine + Environment.NewLine;
            }

            str += "Usage: " + new System.IO.FileInfo(Environment.GetCommandLineArgs()[0]).Name +
                m_usageCmdLine + Environment.NewLine + Environment.NewLine +
                m_usageArgs + Environment.NewLine +
                m_usageGroups + Environment.NewLine;

            return str;
        }


        /// <summary>gets or sets whether argument IDs are case sensitive</summary>
        public bool caseSensitive
        {
            get {return m_caseSensitive;}
            set {m_caseSensitive = value; CheckNotAmbiguous();}
        }

        
        /// <summary>gets the last reported error</summary>
        public string lastError
        {
            get {return (m_lastError.Length != 0 ? m_lastError : "There was no error");}
        }

        

        /// <summary>defines an internal required parameter, ID=APPLICATION_NAME, as a first parameter</summary>
        private void SetFirstArgumentAsAppName()
        {
            if (m_declaredParams.Count > 0 && ((CArgument)m_declaredParams[0]).Id == CommandlineParser.APPLICATION_NAME)
            {
                return;
            }

            CheckNotAmbiguous(CommandlineParser.APPLICATION_NAME);
            CArgument arg = new CStringArgument(CommandlineParser.APPLICATION_NAME, "the application's name", false, "", true);
            m_declaredParams.Insert(0,arg);
            m_iRequiredParams++;
        }


        /// <summary>creates the regular expression for parsing</summary>
        private void BuildRegularExpression()
        {
            // REGULAR EXPRESSION OPTIONS USED IN THIS CODE:
            //
            // Atomic zero-width assertions:	"\G" means "must be after end of previous match";
            // Quantifiers:						"{1}" - exactly once;  "?" - optional (0 or 1 times);
            //									"+"   - at least once; "*" - any number of times
            // Grouping constructs:				"?<id>" - the following section is stored as field "id"
            // Character Classes:				"\w" - a alphanumric character (equal to [a-zA-Z_0-9]
            //									"\s" - any white space character
            // Alteration constructs:			"cat|dog" - either matches are good (checked left-to-right)
            //
            // e.g.: the expression "^[/-]{1}(?<ID>[\w]+)([=:]{1}(?<val>[\w]*))?" will match these:
            //			- "/name=jerry"			ID = "name", val = "jerry"
            //			- "-folder:current		ID = "folder", val = "current"
            //			- "send=true"			PARSE ERROR - no '/' or '-' before ID ("[/-]{1}")
            //			- "s/end=true"			PARSE ERROR - "/end=true" was not at the beginning! ("^")
            //
            m_Syntax
                = @"\G((?<" + SWITCH_TOKEN + @">[\+\-" + m_switchChar + @"]{1})(?<" + ID_TOKEN + @">[\w|?]+)"
                + "(?<"	+ DELIM_TOKEN + ">[" + m_delimChar + @"]?))?"	// optional switch
                + @"(?<" + VALUE_TOKEN + ">(\"[^\"]*\"|\\S*)\\s+){1}";
        }


        // used by the public equivalents
        private void DeclareNumericSwitch(string id, string description, bool fIsOptional, double defaultVal, double minRange, double maxRange)
        {
            if (id.Length == 0)
            {
                throw new EmptyArgumentDeclaredException();
            }
            CheckNotAmbiguous(id);
            CArgument arg = new CNumericArgument(id, description, fIsOptional, defaultVal, minRange, maxRange);
            m_declaredSwitches.Add(arg);
            AddUsageInfo(arg, true, defaultVal);
        }
        private void DeclareStringSwitch(string id, string description, bool fIsOptional, string defaultVal, bool fIsPossibleValsCaseSensitive, params string[] possibleVals)
        {
            if (id.Length == 0)
            {
                throw new EmptyArgumentDeclaredException();
            }
            CheckNotAmbiguous(id);
            CArgument arg = new CStringArgument(id, description, fIsOptional, defaultVal, fIsPossibleValsCaseSensitive, possibleVals);
            m_declaredSwitches.Add(arg);
            AddUsageInfo(arg, true, defaultVal);
        }
        private void DeclareBooleanSwitch(string id, string description, bool fIsOptional, bool defaultVal)
        {
            if (id.Length == 0)
            {
                throw new EmptyArgumentDeclaredException();
            }
            CheckNotAmbiguous(id);
            CArgument arg = new CBooleanArgument(id, description, fIsOptional, defaultVal);
            m_declaredSwitches.Add(arg);
            AddUsageInfo(arg, true, defaultVal);
        }
        private void DeclareParam_Numeric(string id, string description, bool fIsOptional, double defaultVal, double minRange, double maxRange)		
        {
            if (id.Length == 0)
            {
                throw new EmptyArgumentDeclaredException();
            }
            
            if (!fIsOptional && (m_declaredParams.Count > m_iRequiredParams))
            {
                throw new RequiredParamAfterOptionalParamException();
            }
            
            CheckNotAmbiguous(id);
            CArgument arg = new CNumericArgument(id, description, fIsOptional, defaultVal, minRange, maxRange);

            if (!fIsOptional)
            {
                m_iRequiredParams++;
            }

            m_declaredParams.Add(arg);
            AddUsageInfo(arg, false, defaultVal);
        }

        private void DeclareStringParam(string id, string description, bool fIsOptional, string defaultVal, bool fIsPossibleValsCaseSensitive, params string[] possibleVals)		
        {
            if (id.Length == 0)
            {
                throw new EmptyArgumentDeclaredException();
            }
            
            if (!fIsOptional && (m_declaredParams.Count > m_iRequiredParams))
            {
                throw new RequiredParamAfterOptionalParamException();
            }
            
            CheckNotAmbiguous(id);
            CArgument arg = new CStringArgument(id, description, fIsOptional, defaultVal, fIsPossibleValsCaseSensitive, possibleVals);

            if (!fIsOptional)
            {
                m_iRequiredParams++;
            }

            m_declaredParams.Add(arg);
            AddUsageInfo(arg, false, defaultVal);
        }
        
        
        /// <summary>add information to the usage text</summary>
        /// <param name="arg">argument object</param>
        /// <param name="isSwitch">is this a switch? (or a parameter)</param>
        /// <param name="defVal">the default value for this argument</param>
        private void AddUsageInfo(CArgument arg, bool isSwitch, object defVal)
        {
            m_usageCmdLine += (arg.isOptional ? " [" : " ");

            if (isSwitch)
            {
                if (arg.GetType() != typeof(CBooleanArgument))
                {
                    m_usageCmdLine += m_switchChar.ToString() + arg.Id + m_delimChar.ToString() + "x";
                }
                else if (arg.Id == "?")
                {
                    m_usageCmdLine += m_switchChar.ToString() + "?";
                }
                else
                {
                    m_usageCmdLine += "[+|-]" + arg.Id;
                }
            }
            else
            {
                m_usageCmdLine += arg.Id;
            }
            
            m_usageCmdLine += (arg.isOptional ? "]" : "");

            string str =
                (arg.Id == "?" || (isSwitch  && arg.GetType() != typeof(CBooleanArgument)) ? m_switchChar.ToString() : "") + arg.Id;
            if( arg.isOptional ) str = "[" + str + "]";
            str = "  " + str.PadRight(USAGE_COL1 - 3,(char)(0xb7)) + " "; // 0xb7= small middle dot
            
            str += arg.description;
            
            if (arg.Id != "?")
            {
                str += ". Values: " + arg.possibleValues();
                if (arg.isOptional)
                {
                    str += "; default= " + defVal.ToString();
                }
            }

            while( str.Length > 0 )
            {
                if( str.Length <= USAGE_WIDTH )
                {
                    m_usageArgs += str + Environment.NewLine;
                    break;
                }
                int p; // = str.IndexOf(" ",USAGE_WIDTH - 10, 10);
                for( p = USAGE_WIDTH; p > USAGE_WIDTH - 10; p--)
                    if( str[p] == ' ' ) break;
                if( p <= USAGE_WIDTH - 10 ) p = USAGE_WIDTH;
                m_usageArgs += str.Substring(0,p) + Environment.NewLine;
                str = str.Substring(p).TrimStart();
                if( str.Length > 0 ) str = "".PadLeft(USAGE_COL1,' ') + str;
            }
        }
        
        
        /// <summary>read the input of a single parsed switch</summary>
        /// <param name="token">the token before the switch (e.g. "/" in "/name:Avi") - can also be "+" or "-" for booleans</param>
        /// <param name="ID">the switch ID (e.g. "name" in "/name:Avi")</param>
        /// <param name="delim">the delimiter (e.g. ":" in "/name:Avi")</param>
        /// <param name="val">the switch ID (e.g. "Avi" in "/name:Avi")</param>
        /// <returns>true if the value was valid</returns>
        private bool InputSwitch(string token, string ID, string delim, string val)
        {
            if (m_aliases.ContainsKey(ID))
            {
                ID = (string)m_aliases[ID];
            }

            CArgument arg = FindSimilarArg(ID, m_declaredSwitches);
            if (arg == null)
            {
                return false;
            }

            if (arg.GetType() == typeof(CBooleanArgument))
            {
                arg.SetValue(token);
                if (delim.Length != 0 || val.Length != 0)
                {
                    m_lastError
                        = "A boolean switch cannot be followed by a delimiter. Use \"-booleanFlag\", not \"-booleanFlag" + m_delimChar + "\"";
                    return false;
                }

                return true;
            }
            else
            {
                if (delim.Length == 0)
                {
                    m_lastError = "you must use the delimiter '" + m_delimChar + "', e.g. \"" + m_switchChar + "arg" + m_delimChar + "x\"";
                    return false;
                }

                if (arg.SetValue(val))
                {
                    return true;
                }
                else
                {
                    m_lastError = "Switch '" + ID + "' cannot accept '" + val + "' as a value";
                    return false;
                }
            }
        }
        
        /// <summary>read the input of a single parsed parameter</summary>
        /// <param name="val">the string value</param>
        /// <param name="paramIndex">index of parameter in the parsed line</param>
        /// <returns>true if the value was valid</returns>
        private bool InputParam(string val, int paramIndex)
        {
            if (m_declaredParams.Count < paramIndex + 1)
            {
                m_lastError = "Command-line has too many parameters";
                return false;
            }

            CArgument arg = (CArgument)m_declaredParams[paramIndex];

            if (arg.SetValue(val))
            {
                return true;
            }
            else
            {
                m_lastError = "Parameter '" + arg.Id + "' did not have a legal value";
                return false;
            }
        }

        
        /// <summary>get reference to an argument with the exact argument ID</summary>
        /// <param name="argID">the argument's ID</param>
        /// <param name="list">ArrayList to look in (either m_declaredParams or m_declaredSwitches)</param>
        /// <returns>pointer to the argument object, or null if not found</returns>
        private CArgument FindExactArg(string argID, ArrayList list)
        {
            foreach (CArgument A in list)
            {
                if (caseSensitive)
                {
                    if (A.Id == argID)
                    {
                        return A;
                    }
                }
                else
                {
                    if (A.Id.ToUpper() == argID.ToUpper())
                    {
                        return A;
                    }
                }
            }

            return null;
        }

        // get reference to an argument using a substring or its full ID
        private CArgument FindSimilarArg(string argSubstringID, ArrayList list)
        {
            argSubstringID = (caseSensitive ? argSubstringID : argSubstringID.ToUpper());
                
            CArgument arg = null;
            foreach (CArgument A in list)
            {
                string a = (caseSensitive ? A.Id : A.Id.ToUpper());

                if (a.StartsWith(argSubstringID))
                {
                    if (arg != null)
                    {
                        string b = (caseSensitive ? arg.Id : arg.Id.ToUpper());
                        m_lastError = "Ambiguous ID: '" + argSubstringID + "' matches both '" + b + "' and '" + a + "'";
                        return null;
                    }
                    
                    arg = A;
                }
            }

            if (arg == null)
            {
                m_lastError = "No such argument '" + argSubstringID + "'";
            }

            return arg;
        }

        
        private void CheckNotAmbiguous() {CheckNotAmbiguous("");}

        private void CheckNotAmbiguous(string newId)
        {
            CheckNotAmbiguous(newId, m_declaredSwitches);
            CheckNotAmbiguous(newId, m_declaredParams);
        }

        private void CheckNotAmbiguous(string newID, ArrayList argList)
        {
            newID = (caseSensitive ? newID : newID.ToUpper());

            foreach (CArgument A in argList)
            {
                string a = (caseSensitive ? A.Id : A.Id.ToUpper());

                if (a == newID)
                {
                    throw new ArgumentAlreadyDeclaredException(a);
                }

                if (newID.Length != 0 && (a.StartsWith(newID) || newID.StartsWith(a)))
                {
                    throw new AmbiguousArgumentException(a, newID);
                }

                foreach (CArgument B in argList)
                {
                    if (!A.Equals(B))
                    {
                        string b = (caseSensitive ? B.Id : B.Id.ToUpper());

                        if (a == b)
                        {
                            throw new ArgumentAlreadyDeclaredException(a);
                        }

                        if (a.StartsWith(b) || b.StartsWith(a))
                        {
                            throw new AmbiguousArgumentException(a, b);
                        }
                    }
                }
            }
        }


        private bool IsGroupRulesKept()
        {
            foreach (CArgGroups G in m_argGroups)
            {
                uint count = 0;
                foreach (string id in G.Args)
                {
                    CArgument A = FindExactArg(id, m_declaredSwitches);
                    if (A != null && A.isAssigned)
                    {
                        count++;
                    }
                }
                if (!G.InRange(count))
                {
                    m_lastError = G.RangeDescription();
                    return false;
                }
            }
            
            return true;
        }

        private bool IsFirstArgumentAppName()
        {
            return (m_declaredParams.Count > 0 && ((CArgument)m_declaredParams[0]).Id == APPLICATION_NAME);
        }
        
        
        private char		m_switchChar			= DEFAULT_SWITCH;
        private char		m_delimChar				= DEFAULT_DELIM;

        // the regular expression used to parse the command-line
        private string		m_Syntax				= "";
        
        // arguments, groups and aliases declared by the user
        private ArrayList	m_declaredSwitches		= new ArrayList();
        private ArrayList	m_declaredParams		= new ArrayList();
        private uint		m_iRequiredParams		= 0;
        private ArrayList	m_argGroups				= new ArrayList();
        private SortedList	m_aliases				= new SortedList();
        
        private bool		m_caseSensitive			= false;

        // built on-the-fly as arguments are defined, and used by UsageString()
        private string		m_lastError				= "";
        private string		m_usageCmdLine			= "";
        private string		m_usageArgs				= "";
        private string		m_usageGroups			= "";
        
        private bool		m_parseSuccess			= false;

        private const char		DEFAULT_SWITCH		= '/';
        private const char		DEFAULT_DELIM		= ':';
        private const string	SWITCH_TOKEN		= "switchToken";
        private const string	ID_TOKEN			= "idToken";
        private const string	DELIM_TOKEN			= "delimToken";
        private const string	VALUE_TOKEN			= "valueToken";
        private const int		USAGE_COL1			= 25;
        private const int		USAGE_WIDTH			= 79;

        private const string	APPLICATION_NAME	= "RESERVED_ID_APPLICATION_NAME";
    }
}


