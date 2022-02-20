using FormulaEvaluator;
using System;
using System.Linq;
using System.Text;
namespace Test_The_Evaluator_Console_App
{
    class Program
    {
        static void Main(string[] args)
        {
            //my test expressions
            Console.WriteLine(Evaluator.Evaluate("1+1", s => 0));  //9
            Console.WriteLine(Evaluator.Evaluate("(1+1)*2", s => 0)); //2
            Console.WriteLine(Evaluator.Evaluate("3*(1-2)", s => 0)); //-3
            Console.WriteLine(Evaluator.Evaluate("100/(3+(3+4))", s => 0));//10
            Console.WriteLine(Evaluator.Evaluate("2 + 3.5",lookerupper)); // results in 25
            Console.WriteLine(Evaluator.Evaluate("$$$$$$$",s => 0)); // results in throw a error
            
            
        }


        //simple lookup function  
        static  Func<string, double> lookerupper = s => (s == "x") ? 2 : 0;

    }
}