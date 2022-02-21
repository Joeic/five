using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SS;
using System.Collections.Generic;
using SpreadsheetUtilities;
using System.IO;

/// <summary> 
/// Author:    Joey Cai
/// Partner:  None
/// Date:      2/18/2022
/// Course:    CS 3500, University of Utah, School of Computing 
/// Copyright: CS 3500 and Joey Cai - This work may not be copied for use in Academic Coursework. 
/// 
/// I, Joey Cai, certify that I wrote this code from scratch and did not copy it in part or whole from  
/// another source.  All references used in the completion of the assignment are cited in my README file. 
/// 
/// File Contents 
/// 
///  This proeject will test all methods from SpreadSheet
/// </summary>
namespace SpreadsheetTests
{
	[TestClass]
	public class SpreadsheetTests
	{
		public Spreadsheet sheet1;
		/// <summary>
		/// init sheet 1
		/// </summary>
		[TestInitialize]
		public void setup()
		{
			sheet1 = new Spreadsheet();
		}
		/// <summary>
		/// test constructors. successfully run each
		/// </summary>
		[TestMethod]
		public void TestConstructor()
		{
			
			
			Assert.IsTrue(sheet1.IsValid("I am for test"));
			Assert.IsTrue(sheet1.Normalize("dead") == "dead");
			Assert.IsTrue(sheet1.Version == "default");
			
			//test 3 arg constructor
			sheet1 = new Spreadsheet(s => (s.Length >= 2) ? true : false, 
				s => s.Replace(" ", ""),
				"version1");
			Assert.IsTrue(sheet1.IsValid("B1"));
			Assert.IsFalse(sheet1.IsValid("B"));
			Assert.IsTrue(sheet1.Normalize("d e a d") == "dead");
			Assert.IsTrue(sheet1.Version == "version1");
			sheet1.SetContentsOfCell("B1","loaded!");

			string savePath = "save test.xml";
			sheet1.Save(savePath);
			sheet1 = new Spreadsheet(
				savePath,
				s => (s.Length >= 2) ? true : false, 
				s => s.Replace(" ", ""),
				"version1");
			Assert.AreEqual("loaded!",(string)sheet1.GetCellContents("B1"));
		}
		
		

	}
}
