using System;
using System.Text.RegularExpressions;
using System.Collections;
namespace FormulaEvaluator
{
    public static class Evaluator
    {
        //storage opreation
        static Stack<string> op = new Stack<string>();
        //storage number
        static Stack<double> num = new Stack<double>();
        //storage priority
           static Dictionary<string, int> pr = new Dictionary<string, int>
            {{"+", 1}, {"-", 1}, {"*", 2}, {"/", 2}};

        public delegate double Lookup(String variableName);



        //do once operation
        public static void Eval()
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
        public static bool Isdigit(string s)
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
        public static double Decode(string s)
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

        public static bool isVar(string s)
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


        public static double Evaluate(String s, Func<string, double> variableEvaluator)
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







    }

}