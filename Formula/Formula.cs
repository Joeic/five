// Skeleton written by Joe Zachary for CS 3500, September 2013
// Read the entire skeleton carefully and completely before you
// do anything else!

// Version 1.1 (9/22/13 11:45 a.m.)

// Change log:
//  (Version 1.1) Repaired mistake in GetTokens
//  (Version 1.1) Changed specification of second constructor to
//                clarify description of how validation works

// (Daniel Kopta) 
// Version 1.2 (9/10/17) 

// Change log:
//  (Version 1.2) Changed the definition of equality with regards
//                to numeric tokens


using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel.Design;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using System.Globalization;
using StringExtension;
using System.Runtime.CompilerServices;
using SpreadsheetUtilities;

namespace SpreadsheetUtilities
{
  /// <summary>
  /// Represents formulas written in standard infix notation using standard precedence
  /// rules.  The allowed symbols are non-negative numbers written using double-precision 
  /// floating-point syntax (without unary preceeding '-' or '+'); 
  /// variables that consist of a letter or underscore followed by 
  /// zero or more letters, underscores, or digits; parentheses; and the four operator 
  /// symbols +, -, *, and /.  
  /// 
  /// Spaces are significant only insofar that they delimit tokens.  For example, "xy" is
  /// a single variable, "x y" consists of two variables "x" and y; "x23" is a single variable; 
  /// and "x 23" consists of a variable "x" and a number "23".
  /// 
  /// Associated with every formula are two delegates:  a normalizer and a validator.  The
  /// normalizer is used to convert variables into a canonical form, and the validator is used
  /// to add extra restrictions on the validity of a variable (beyond the standard requirement 
  /// that it consist of a letter or underscore followed by zero or more letters, underscores,
  /// or digits.)  Their use is described in detail in the constructor and method comments.
  /// </summary>
  public class Formula
  {
    
    private List<string> memo;//memory to hold formula
    //storage opreation
    static Stack<string> op = new Stack<string>();
    //storage number
    static Stack<double> num = new Stack<double>();
    //storage priority
    static Dictionary<string, int> pr = new Dictionary<string, int>
	    {{"+", 1}, {"-", 1}, {"*", 2}, {"/", 2}};
    public Formula(String formula) :
      this(formula, s => s, s => true)
    {
    }
     public delegate double Lookup(String variableName);



        //do once operation
        public  void Eval()
        {

            if (num.Count < 2 || op.Count == 0)
                throw new ArgumentException("Error");
            double b = num.Peek(); num.Pop();
            double a = num.Peek(); num.Pop();
            string c = op.Peek(); op.Pop();
           
            double x = 0;
            if (c == "+")
                x = a + b;
            else if (c == "-")
                x = a - b;
            else if (c == "*")
                x = a * b;
            else
            {
                if (b == 0)
                {
                    throw new ArgumentException("error: divide by zero");
                }
                x = a / b;
            }
                
           
            num.Push(x);

        }
        //return true if s is a digit
        public  bool Isdigit(string s)
        {
            int n = s.Length;
            ;


            int count = 0;
            for (int i = 0; i <n; i++)
            {
                if (s[i] == '.')
                {
                    count++;
                    continue;
                }
                if ((s[i] < '0') || (s[i] > '9'))
                    return false;
            }
            
            if (n == 0)
                return false;

            if (s[0] == '.' || s[n - 1] == '.')
                return false;
            if (count > 1)
                return false;
            return true;
        }

        //get the int value of string
        public double Decode(string s)
        {
            int n = s.Length;
            double r = 1;
            double sum = 0;
            int count=0;
            for (int i = n - 1; i > 0; i--)
            {
                if (s[i] == '.')
                {
                    count = (n - 1 - i);
                    break;
                }
            }
            s = s.Replace(".", "");
            n = s.Length;
            for (int i = n - 1; i >= 0; i--)
            {
                sum += (s[i]-'0') * r;
                r *= 10;
            }

            r = 1;
            while (count>0)
            {
                r *= 10;
                count--;
            }

           // Console.WriteLine(sum/r);
            return sum/r;
        }

        public  bool isVar(string s)
        {
            if (s.Length == 0)
                return false;
            if (s[0] >= '0' && s[0] <= '9')
                return false;
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == '_' || (s[i] >= 'a' && s[i] <= 'z') || (s[i] >= 'A' && s[i] <= 'Z' || (s[i]<='9'&&s[i]>='0')))
                    continue;
                else
                {
                    return false;
                }
				
            }

            return true;
        }


        public  double Evaluate(String s, Func<string, double> variableEvaluator)
        {

           
            s = s.Replace(" ", "");
            if (s.Length == 0)
                throw new ArgumentException("Erroe");
            string[] substrings = Regex.Split(s, "(\\()|(\\))|(-)|(\\+)|(\\*)|(/)");
            
            
            foreach (string str in substrings)
            {
               
                if (Isdigit(str))
                {
                   
                    num.Push(Decode(str));
                }
                else if (str == "(")
                    op.Push(str);
                else if (str == ")")
                {
                    if (op.Count == 0)
                        throw new ArgumentException("Error");
                    while (true)
                    {
                        if (op.Count == 0)
                            throw new ArgumentException("Error");
                        if (op.Peek() == "(")
                        {
                            break;
                        }
                       
                        Eval();
                    }
                    op.Pop();
                }
                else if ((str == "+") || (str == "-") || (str == "*") || (str == "/"))
                {
                    
                    while ((op.Count > 0) && (op.Peek() != "(") && (pr[op.Peek()] >= pr[str]))
                        Eval();
                 
                    op.Push(str);
                }
                else if (str == "")
                {
                   
                    continue;
                }
                else if (isVar(str))
                {
                    double value= variableEvaluator(str);
                    num.Push(value);

                }
                else
                {
                   
                    throw new ArgumentException("This expression contains invalid operators or variables.");
                }
            }


            while (op.Count > 0)
                Eval();

            double res =  (num.Peek());
            num.Pop();
            if(num.Count==0)
            return res;
            else
            {
                throw new ArgumentException("wrong ");
            }
        }
    
    public Formula(String formula, Func<string, string> normalize, Func<string, bool> isValid)
    {
      
			//may be a empty string
			if (formula.Length > 0)
			{
				//get token
				formula = formula.Replace(" ", "");
				string[]  trialTokens= GetTokens(formula).ToArray();

				int nt= trialTokens.Length - 1;
				//string strt = trialTokens[0];
				//check begin
				if (!(trialTokens[0].IsBegin()))
				{
					throw new FormulaFormatException("Error,formula begin with wrong character");
				}
				//check end
				if (!((trialTokens[nt]).IsRightEnd()))
				{
					throw new FormulaFormatException("Error,formula end with wrong character");
				}
				
				int leftParen  = 0;//count"("
				int rightParen = 0;//count")"
				for (int i = 0; i < trialTokens.Length; i++)
				{
					
					if (trialTokens[i].IsOperator())
					{
						//check next string 
						if (i <= trialTokens.Length - 2 && !trialTokens[i+1].IsRightAfterOperator())
						{
							throw new FormulaFormatException("error: wrong character after"+trialTokens[i]);
						}
					}
					
					else if (trialTokens[i] == "(")
					{
						leftParen++;
						
						if (i <= trialTokens.Length - 2 && !(trialTokens[i+1].IsRigheAfterLeftParent()))
						{
							throw new FormulaFormatException("error: wrong character after (");
						}
					}
					
					else if (trialTokens[i] == ")")
					{
						rightParen++;
						
						if (i <= trialTokens.Length - 2 && !trialTokens[i+1].IsRightAfterRightParent())
						{
							throw new FormulaFormatException("error: wrong character after )");
						}
					}
					
					else if (trialTokens[i].IsVar())
					{
						
						if (i <= trialTokens.Length - 2 &&
							!(trialTokens[i+1].IsRighrAfterVar()))
						{
							throw new FormulaFormatException("error: wrong character after var "+trialTokens[i]);
						}
						else
						{
							trialTokens[i] = normalize(trialTokens[i]);
							if (!isValid(trialTokens[i]))
							{
								throw new FormulaFormatException("error: invalid normalized variable.");
							}
						}
						


					}
					else if (trialTokens[i].IsDouble())
					{
						
						if (i <= trialTokens.Length - 2 && !(trialTokens[i+1].IsRightAfterDouble()))
						{
							throw new FormulaFormatException("error: wrong character after value " + trialTokens[i]);
						}

					}
					else
					{
						//pass
					}
					
					

				}
				if (rightParen != leftParen)
				{
					throw new FormulaFormatException("error:  parentheses do not match");
				}
				
				memo = new List<string>(trialTokens);
				
			}
			else
			{
				//empty string
				throw new FormulaFormatException("error: empty input");
			}
    }

   
    public object Evaluate(Func<string, double> lookup)
    {
      return  this.Evaluate(this.ToString(),lookup);

    }

    /// <summary>
    /// Enumerates the normalized versions of all of the variables that occur in this 
    /// formula.  No normalization may appear more than once in the enumeration, even 
    /// if it appears more than once in this Formula.
    /// 
    /// For example, if N is a method that converts all the letters in a string to upper case:
    /// 
    /// new Formula("x+y*z", N, s => true).GetVariables() should enumerate "X", "Y", and "Z"
    /// new Formula("x+X*z", N, s => true).GetVariables() should enumerate "X" and "Z".
    /// new Formula("x+X*z").GetVariables() should enumerate "x", "X", and "z".
    /// </summary>
    public IEnumerable<String> GetVariables()
    {
	    HashSet<string> res = new HashSet<string>();
	    foreach (var t in memo)
	    {
		    if (t.IsVar())
			    res.Add(t);
	    }
	    return res.ToList();
    }

    /// <summary>
    /// Returns a string containing no spaces which, if passed to the Formula
    /// constructor, will produce a Formula f such that this.Equals(f).  All of the
    /// variables in the string should be normalized.
    /// 
    /// For example, if N is a method that converts all the letters in a string to upper case:
    /// 
    /// new Formula("x + y", N, s => true).ToString() should return "X+Y"
    /// new Formula("x + Y").ToString() should return "x+Y"
    /// </summary>
    public override string ToString()
    {
	    string res = string.Empty;
	    foreach (var t in memo)
	    {
		    res += t;
	    }
	    return res;
    }

    /// <summary>
    ///  <change> make object nullable </change>
    ///
    /// If obj is null or obj is not a Formula, returns false.  Otherwise, reports
    /// whether or not this Formula and obj are equal.
    /// 
    /// Two Formulae are considered equal if they consist of the same tokens in the
    /// same order.  To determine token equality, all tokens are compared as strings 
    /// except for numeric tokens and variable tokens.
    /// Numeric tokens are considered equal if they are equal after being "normalized" 
    /// by C#'s standard conversion from string to double, then back to string. This 
    /// eliminates any inconsistencies due to limited floating point precision.
    /// Variable tokens are considered equal if their normalized forms are equal, as 
    /// defined by the provided normalizer.
    /// 
    /// For example, if N is a method that converts all the letters in a string to upper case:
    ///  
    /// new Formula("x1+y2", N, s => true).Equals(new Formula("X1  +  Y2")) is true
    /// new Formula("x1+y2").Equals(new Formula("X1+Y2")) is false
    /// new Formula("x1+y2").Equals(new Formula("y2+x1")) is false
    /// new Formula("2.0 + x7").Equals(new Formula("2.000 + x7")) is true
    /// </summary>
    public override bool Equals(object? obj)
    {
	    
	    {
		    string s1 = this.ToString();
		    string? s2 = obj.ToString();
		   return s1 == s2;
	    }
	    
	    //return false;
    }

    /// <summary>
    ///   <change> We are now using Non-Nullable objects.  Thus neither f1 nor f2 can be null!</change>
    /// Reports whether f1 == f2, using the notion of equality from the Equals method.
    /// 
    /// </summary>
    public static bool operator ==(Formula f1, Formula f2)
    {
     
      if (Object.ReferenceEquals(f1, null) && Object.ReferenceEquals(f2, null))
      {
        return true;
        
      }
      else if(Object.ReferenceEquals(f1,null) || Object .ReferenceEquals(f2,null))
      {
        return false;
      }
      else
      {
        return f1.Equals(f2);
      }
        
      
     
    }

    /// <summary>
    ///   <change> We are now using Non-Nullable objects.  Thus neither f1 nor f2 can be null!</change>
    ///   <change> Note: != should almost always be not ==, if you get my meaning </change>
    ///   Reports whether f1 != f2, using the notion of equality from the Equals method.
    /// </summary>
    public static bool operator !=(Formula f1, Formula f2)
    {
      return !(f1 == f2);
    }

    /// <summary>
    /// Returns a hash code for this Formula.  If f1.Equals(f2), then it must be the
    /// case that f1.GetHashCode() == f2.GetHashCode().  Ideally, the probability that two 
    /// randomly-generated unequal Formulae have the same hash code should be extremely small.
    /// </summary>
    public override int GetHashCode()
    {
	    string str = this.ToString();
      int  r = 1;
      int sum = 0;
      for (int i = 0; i < str.Length; i++)
      {
        sum += (((str[i]) )*r)%236887699;
        r *= 15485867;
        r %= 236887699;
        sum%=236887699;
      }

      return sum;
    }

    /// <summary>
    /// Given an expression, enumerates the tokens that compose it.  Tokens are left paren;
    /// right paren; one of the four operator symbols; a string consisting of a letter or underscore
    /// followed by zero or more letters, digits, or underscores; a double literal; and anything that doesn't
    /// match one of those patterns.  There are no empty tokens, and no token contains white space.
    /// </summary>
    private static IEnumerable<string> GetTokens(String formula)
    {
      // Patterns for individual tokens
      String lpPattern = @"\(";
      String rpPattern = @"\)";
      String opPattern = @"[\+\-*/]";
      String varPattern = @"[a-zA-Z_](?: [a-zA-Z_]|\d)*";
      String doublePattern = @"(?: \d+\.\d* | \d*\.\d+ | \d+ ) (?: [eE][\+-]?\d+)?";
      String spacePattern = @"\s+";

      // Overall pattern
      String pattern = String.Format("({0}) | ({1}) | ({2}) | ({3}) | ({4}) | ({5})",
                                      lpPattern, rpPattern, opPattern, varPattern, doublePattern, spacePattern);

      // Enumerate matching tokens that don't consist solely of white space.
      foreach (String s in Regex.Split(formula, pattern, RegexOptions.IgnorePatternWhitespace))
      {
        if (!Regex.IsMatch(s, @"^\s*$", RegexOptions.Singleline))
        {
          yield return s;
        }
      }

    }
  }

  /// <summary>
  /// Used to report syntactic errors in the argument to the Formula constructor.
  /// </summary>
  public class FormulaFormatException : Exception
  {
    /// <summary>
    /// Constructs a FormulaFormatException containing the explanatory message.
    /// </summary>
    public FormulaFormatException(String message)
        : base(message)
    {
    }
  }

  /// <summary>
  /// Used as a possible return value of the Formula.Evaluate method.
  /// </summary>
  public struct FormulaError
  {
    /// <summary>
    /// Constructs a FormulaError containing the explanatory reason.
    /// </summary>
    /// <param name="reason"></param>
    public FormulaError(String reason)
        : this()
    {
      Reason = reason;
    }

    /// <summary>
    ///  The reason why this FormulaError was created.
    /// </summary>
    public string Reason { get; private set; }
  }
}


// <change>
//   If you are using Extension methods to deal with common stack operations (e.g., checking for
//   an empty stack before peeking) you will find that the Non-Nullable checking is "biting" you.
//
//   To fix this, you have to use a little special syntax like the following:
//
//       public static bool OnTop<T>(this Stack<T> stack, T element1, T element2) where T : notnull
//
//   Notice that the "where T : notnull" tells the compiler that the Stack can contain any object
//   as long as it doesn't allow nulls!
// </change>
namespace StringExtension
{
	//use to handle string 
	public static class StringExtension
	{

		public static bool IsDouble(this string s)
		{
			double temp;
			return s.Igit()||(Double.TryParse(s, System.Globalization.NumberStyles.AllowExponent, CultureInfo.InvariantCulture,out temp));

		}

		public static bool IsBegin(this string s)
		{
			return (s.IsVar() || s.IsDouble() || s.Equals("("));
		}
		
		public static bool IsVar(this string s)
		{
			
			return true;
		}

		public static bool IsOperator(this string s)
		{
			return (s == "+" || s == "-" || s == "*" || s == "/");
		}

		public static bool IsRightAfterOperator(this string s)
		{
			
			return s.IsDouble() || s.IsVar() || s == "(";
		}
		
		public static bool IsRigheAfterLeftParent(this string s)
		{
			return s.IsDouble() || s == ")" || s.IsVar();
		}

		public static bool IsRighrAfterVar(this string s)
		{
			
			return s.IsOperator() || s == ")";
		}

		public static bool IsRightAfterRightParent(this string s)
		{
			
			return s.IsOperator() || s == ")";
		}

		public static bool IsRightAfterDouble(this string s)
		{

			return s.IsOperator() || s == ")";
		}

		public static bool IsRightEnd(this string s)
		{
			
			return (s.IsVar() ||s.IsDouble() || s.Equals(")"));
		}
		
		public static bool Igit(this string s)
		{
			int n = s.Length;
			;


			int count = 0;
			for (int i = 0; i <n; i++)
			{
				if (s[i] == '.')
				{
					count++;
					continue;
				}
				if ((s[i] < '0') || (s[i] > '9'))
					return false;
			}
            
			if (n == 0)
				return false;

			if (s[0] == '.' || s[n - 1] == '.')
				return false;
			if (count > 1)
				return false;
			return true;
		}
	}
}

