﻿using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SpreadsheetUtilities;
using System.Text.RegularExpressions;

namespace FormulaTest
{
	
	[TestClass]
	public class FormulaTest
	{
		//test ToString
        [TestMethod]
        public void ConstructorTest1() {
            Formula f1 = new Formula("2.6*3.6");
            Assert.AreEqual("2.6*3.6", f1.ToString());
        }
	    
        //test Tostring
        [TestMethod]
        public void ConstructorTest2() {
            Formula f1 = new Formula("u2*v3");
            Assert.AreEqual("u2*v3", f1.ToString());
        }
	   
        //test Tostring
        [TestMethod]
        public void ConstructorTest3() {
            Formula f1 = new Formula("x2+(y3)");
            Assert.AreEqual("x2+(y3)", f1.ToString());
        }
		
        //test upper
		[TestMethod]
		public void Constructor2Test()
		{
			Formula f1 = new Formula("z", s => s.ToUpper(), s => (s == "Z") ? true : false);
			Assert.AreEqual("Z", f1.ToString());
		}
		
		//test upper
		[TestMethod]
		public void Constructor2Test1()
		{
			Formula f1 = new Formula("y+2", s => s.ToUpper(), s => (s == "Y") ? true : false);
			Assert.AreEqual("Y+2", f1.ToString());
		}
       //empty test
        [TestMethod]
        [ExpectedException(typeof(FormulaFormatException))]
        public void ConstructorFailTest1() {
            Formula f1 = new Formula("");
        }
	   
        //wrong begin test
	   [TestMethod]
	   [ExpectedException(typeof(FormulaFormatException))]
	   public void ConstructorFailTest2()
	   {
		   Formula f1 = new Formula("*3");
	   }
	  
	   //wrong begin test
	   [TestMethod]
	   [ExpectedException(typeof(FormulaFormatException))]
	   public void ConstructorFailTest3()
	   {
		   Formula f1 = new Formula("(5");
	   }
	  
	   //doesm't match "()"
	   [TestMethod]
	   [ExpectedException(typeof(FormulaFormatException))]
	   public void ConstructorFailTest4()
	   {
		   Formula f1 = new Formula("((2))))");
	   }
	   
	   //wrong begin test
	   [TestMethod]
	   [ExpectedException(typeof(FormulaFormatException))]
	   public void ConstructorFailTest7()
	   {
		   Formula f1 = new Formula("/");
	   }
	 
	   // test equal
        [TestMethod]
        public void EqualsTest() {
            Formula f1 = new Formula("4+5");
            Formula f2 = new Formula("4+5");
            Assert.IsTrue(f1.Equals(f2));
        }
        
        //test equal
        [TestMethod]
        public void EqualsTest2() {
            Formula f1 = new Formula("2+6");
            Formula f2 = new Formula("2+8");
            Assert.IsFalse(f1.Equals(f2));
        }
        [TestMethod]
        public void EvaluateTest1() {
		Func<string,double> looker = s=>0;
			Assert.AreEqual(5.59, new Formula("2.59+3").Evaluate(looker));

		}
		
		
		
		[TestMethod]
		public void tostringTest()
		{
			Formula f1 = new Formula("9*2");
			Assert.IsTrue(f1.ToString() == "9*2");

		}
		
		
		[TestMethod]
		public void tostringTest3()
		{
			Formula f1 = new Formula("9 * 2");
			Assert.IsTrue(f1.ToString() == "9*2");

		}
		[TestMethod]
		public void GetHashcodeTest()
		{
			Formula f1 = new Formula("y*2");
			Formula f2 = new Formula("y*2");
			Assert.IsTrue(f1.GetHashCode() == f2.GetHashCode());
		}
		
		
		[TestMethod]
		public void TestOperatorEq1()
		{
			Formula f1 = new Formula("1+1");
			Formula f2 = null;
			Assert.IsFalse(f1 == f2);
		}
		
		[TestMethod]
		public void TestOperatorEq2()
		{
			Formula f1 = null;
			Formula f2 = new Formula("1+1");
			Assert.IsFalse(f1 == f2);
		}
	
		[TestMethod]
		public void TestOperatorEq3()
		{
			Formula f1 = null;
			Formula f2 = null;
			Assert.IsTrue(f1 == f2);
		}
		
		[TestMethod]
		public void TestOperatorNEq()
		{
			Formula f1 = null;
			Formula f2 = null;
			Assert.IsFalse(f1 != f2);
		}
		
		
	
		[TestMethod]
		public void TestOperatorNEq3()
		{
			Formula f1 = new Formula("1+1");
			Formula f2 = new Formula("2+2");
			Assert.IsTrue(f1 != f2);
		}
		
		[TestMethod]
		public void testEval()
		{
			Formula f1 = new Formula("20 + 31.5");
			
			Assert.AreEqual(51.5, f1.Evaluate(lookerupper));
		}
		
		[TestMethod]
		public void testEval2()
		{
			Formula f1 = new Formula("2 + 3.5 + x");
			Assert.AreEqual(22.5, f1.Evaluate(lookerupper));
		}
		
		
		private Func<string, double> lookerupper = s => (s == "x") ? 17 : 0;
		
	}
}
