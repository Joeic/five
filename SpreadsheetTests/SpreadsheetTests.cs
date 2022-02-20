﻿using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SS;
using System.Collections.Generic;
using SpreadsheetUtilities;
using System.IO;


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
			
			
			Assert.IsTrue(sheet1.IsValid("any old string"));
			Assert.IsTrue(sheet1.Normalize("dead") == "dead");
			Assert.IsTrue(sheet1.Version == "default");
			
			//test 3 arg constructor
			sheet1 = new Spreadsheet(s => (s.Length >= 2) ? true : false, 
				s => s.Replace(" ", ""),
				"version1");
			Assert.IsTrue(sheet1.IsValid("A1"));
			Assert.IsFalse(sheet1.IsValid("A"));
			Assert.IsTrue(sheet1.Normalize("d e a d") == "dead");
			Assert.IsTrue(sheet1.Version == "version1");
			sheet1.SetContentsOfCell("A1","loaded!");

			string savePath = "save 1.xml";
			sheet1.Save(savePath);
			sheet1 = new Spreadsheet(
				savePath,
				s => (s.Length >= 2) ? true : false, 
				s => s.Replace(" ", ""),
				"version1");
			Assert.AreEqual("loaded!",(string)sheet1.GetCellContents("A1"));
		}
		
		

	}
}
