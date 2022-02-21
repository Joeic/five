using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SS;
using SpreadsheetUtilities;
using System.Text.RegularExpressions;
using System.Globalization;
using System.IO;
using System.Linq.Expressions;
using System.Xml;
using StringExtension;

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
///  This proeject will recognize all math string input and calculate its result
/// </summary>
 namespace SS
{

	public class Spreadsheet : AbstractSpreadsheet
	{
        private DependencyGraph DG;
        private Dictionary<string, Cell> Sheet;
        /// <summary>
        /// True if this spreadsheet has been modified since it was created or saved                  
        /// (whichever happened most recently); false otherwise.
        /// </summary>
        public override bool Changed
        {
            get;
            protected set;
        }
        
        
        public Spreadsheet()
            : base(s => true, s => s, "default")
        {
            Changed = false;
            Sheet = new Dictionary<string, Cell>();
            DG = new DependencyGraph();
        }
        /// <summary>
        /// Constructs an abstract spreadsheet by recording its variable validity test,
        /// its normalization method, and its version information.  
        /// </summary>
        /// 
        /// <remarks>
        ///   The variable validity test is used throughout to determine whether a string that consists of 
        ///   one or more letters followed by one or more digits is a valid cell name.  The variable
        ///   equality test should be used throughout to determine whether two variables are equal.
        /// </remarks>
        /// 
        /// <param name="isValid">   defines what valid variables look like for the application</param>
        /// <param name="normalize"> defines a normalization procedure to be applied to all valid variable strings</param>
        /// <param name="version">   defines the version of the spreadsheet (should it be saved)</param>
        public Spreadsheet(Func<string, bool> _isValid, Func<string, string> _normalize, string _version)
            : base(_isValid, _normalize, _version)
        {
            Changed = false;
            Sheet = new Dictionary<string, Cell>();
            DG = new DependencyGraph();

        }
        
        public Spreadsheet(string _filePath, Func<string, bool> _isValid, Func<string, string> _normalize, string _version)
            : base(_isValid, _normalize, _version)
        {
            Changed = false;
            Sheet = new Dictionary<string, Cell>();
            DG = new DependencyGraph();
            
            if (_version == GetSavedVersion(_filePath))
                load(_filePath);
            else
                throw new SpreadsheetReadWriteException("Error: wrong version");
            
        }
        
        /// <summary>
        ///   Look up the version information in the given file. If there are any problems opening, reading, 
        ///   or closing the file, the method should throw a SpreadsheetReadWriteException with an explanatory message.
        /// </summary>
        /// 
        public override string GetSavedVersion(string filename)
        {
           
           
            try
            {
                using (XmlReader reader = XmlReader.Create(filename))
                {
                    try
                    {
                        while (reader.Read())
                        {
                            if (reader.IsStartElement())
                            {
                               
                                  if(reader.Name== "spreadsheet")
                                  {
                                      var v=reader["version"];
                                      return v;
                                  }
                                       
                                  else
                                        throw new SpreadsheetReadWriteException("Error: wrong while loading spreadsheet");
                                
                            }
                        }
                    }
                    catch (XmlException ex)
                    {
                        throw new SpreadsheetReadWriteException("Error: wrong when parsing xml" + ex.Message);
                    }
                   
                    throw new SpreadsheetReadWriteException("Error: xml document is empty");
                }
            }
           
            catch (Exception ex)
            {
                if (ex is FileNotFoundException)
                    throw new SpreadsheetReadWriteException("Error: file  not exist");
                else if (ex is DirectoryNotFoundException)
                    throw new SpreadsheetReadWriteException("Error: directory not exist" + ex.Message);
                else
                    throw new SpreadsheetReadWriteException("Something wrong!" + ex.Message);
            }
           
        }
        /// <summary>
        ///  load the file and set the contents-
        /// </summary>
      
        private void load(string filename)
        {
            
            using (XmlReader reader = XmlReader.Create(filename))
            {
                string name = "";
                string contents = "";
                while (reader.Read())
                {
                    if (reader.IsStartElement())
                    {
                        if (reader.Name == "spreadsheet")
                        {
                            Version = reader["version"];
                        }
                        else if (reader.Name == "cell")
                        {
                            reader.Read();
                            name = reader.ReadElementContentAsString();
                            contents = reader.ReadElementContentAsString();
                            SetContentsOfCell(name, contents);
                        }
                        else
                        {
                            throw new SpreadsheetReadWriteException("Error: wrong XML");
                        }
                        
                    }
                }
            }


        }
        /// <summary>
        /// Writes the contents of this spreadsheet to the named file using an XML format.
        /// The XML elements should be structured as follows:
        /// 
        /// <spreadsheet version="version information goes here">
        /// 
        /// <cell>
        /// <name>cell name goes here</name>
        /// <contents>cell contents goes here</contents>    
        /// </cell>
        /// 
        /// </spreadsheet>
        /// 
        /// There should be one cell element for each non-empty cell in the spreadsheet.  
        /// If the cell contains a string, it should be written as the contents.  
        /// If the cell contains a double d, d.ToString() should be written as the contents.  
        /// If the cell contains a Formula f, f.ToString() with "=" prepended should be written as the contents.
        /// 
        /// If there are any problems opening, writing, or closing the file, the method should throw a
        /// SpreadsheetReadWriteException with an explanatory message.
        /// </summary>
        public override void Save(string filename)
        {
            try
            {
                using (XmlWriter writer = XmlWriter.Create(filename))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("spreadsheet");
                    writer.WriteAttributeString("version", Version);
                    foreach (Cell cell in Sheet.Values)
                    {
                        writer.WriteStartElement("cell");
                        writer.WriteElementString("name", cell.Name);
                        writer.WriteElementString("contents", writeContents(cell.Contents));
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                    writer.Dispose();
                }
            }
            catch { throw new SpreadsheetReadWriteException("error writing to spreadsheet file"); }

        }
        /// <summary>
        /// return the cell contents
        /// 
	/// </summary>
        private string writeContents(object cellContents)
        {
            if (cellContents is Formula)
                return "=" + cellContents.ToString();
            else if (cellContents is double)
                return cellContents.ToString();
            else if (cellContents is string)
                return cellContents as string;
            else
                throw new SpreadsheetReadWriteException("error writing Cell Contents");
        }
	
	/// <summary>
        /// find the contents and return 
        /// 
	/// </summary>
        private object readContents(string objInfo)
        {
            double temp;
            if (Double.TryParse(objInfo, out temp))
                return temp;
            else if (objInfo[0] == '=')
                return new Formula(objInfo.Substring(1));
            else
                 return objInfo;
        }
        
	/// <summary>
        /// If name is invalid, throws an InvalidNameException.
        /// </summary>
        ///
        /// <exception cref="InvalidNameException"> 
        ///   If the name is invalid, throw an InvalidNameException
        /// </exception>
        /// 
        /// <param name="name"> The name of the cell that we want the value of (will be normalized)</param>
        /// 
        /// <returns>
        ///   Returns the value (as opposed to the contents) of the named cell.  The return
        ///   value should be either a string, a double, or a SpreadsheetUtilities.FormulaError.
        /// </returns>
        public override object GetCellValue(string name)
        {

            if (name == null || !name.IsVar() || !IsValid(Normalize(name)))
                throw new InvalidNameException();
            if (Sheet.ContainsKey(name))
                return Sheet[name].Value;
            else
                return string.Empty;
            

        }
        /// <summary>
        ///   look up the name and find return its value
        /// </summary>
      
        public double lookerupper(string name)
        {
            if (!Sheet.ContainsKey(name))
            {
                throw new ArgumentException("Error: cell value was not a double");

               
            }
            var t = Sheet[name].Value;
            if (t is double)
                return (double)t;
            throw new ArgumentException("Error: cell value was not a double");
            
        }
         /// <summary>
        ///   <para>Sets the contents of the named cell to the appropriate value. </para>
        ///   <para>
        ///       First, if the content parses as a double, the contents of the named
        ///       cell becomes that double.
        ///   </para>
        ///
        ///   <para>
        ///       Otherwise, if content begins with the character '=', an attempt is made
        ///       to parse the remainder of content into a Formula.  
        ///       There are then three possible outcomes:
        ///   </para>
        ///
        ///   <list type="number">
        ///       <item>
        ///           If the remainder of content cannot be parsed into a Formula, a 
        ///           SpreadsheetUtilities.FormulaFormatException is thrown.
        ///       </item>
        /// 
        ///       <item>
        ///           If changing the contents of the named cell to be f
        ///           would cause a circular dependency, a CircularException is thrown,
        ///           and no change is made to the spreadsheet.
        ///       </item>
        ///
        ///       <item>
        ///           Otherwise, the contents of the named cell becomes f.
        ///       </item>
        ///   </list>
        ///
        ///   <para>
        ///       Finally, if the content is a string that is not a double and does not
        ///       begin with an "=" (equal sign), save the content as a string.
        ///   </para>
        /// </summary>
        ///
        /// <exception cref="InvalidNameException"> 
        ///   If the name parameter is null or invalid, throw an InvalidNameException
        /// </exception>
        /// 
        /// <exception cref="SpreadsheetUtilities.FormulaFormatException"> 
        ///   If the content is "=XYZ" where XYZ is an invalid formula, throw a FormulaFormatException.
        /// </exception>
        /// 
       
        /// 
        /// <param name="name"> The cell name that is being changed</param>
        /// <param name="content"> The new content of the cell</param>
        /// 
        /// <returns>
        ///       <para>
        ///           This method returns a list consisting of the passed in cell name,
        ///           followed by the names of all other cells whose value depends, directly
        ///           or indirectly, on the named cell. The order of the list MUST BE any
        ///           order such that if cells are re-evaluated in that order, their dependencies 
        ///           are satisfied by the time they are evaluated.
        ///       </para>
        ///
        ///       <para>
        ///           For example, if name is A1, B1 contains A1*2, and C1 contains B1+A1, the
        ///           list {A1, B1, C1} is returned.  If the cells are then evaluate din the order:
        ///           A1, then B1, then C1, the integrity of the Spreadsheet is maintained.
        ///       </para>
        /// </returns>
        
        public override IList<string> SetContentsOfCell(string name, string content)
        {
            //null content
            if (content == null)
                throw new ArgumentNullException();
            //check name
            if (name == null || ! Normalize(name).IsVar())
                throw new InvalidNameException();
            //is the name vaild 
            if (IsValid(name))
            {
                if (content == "")
                {
                    Sheet.Remove(name);
                    DG.ReplaceDependents(name, new HashSet<string>());
                    return new List<string>(GetCellsToRecalculate(name));
                }
                if (content[0] == '=')
                    return SetCellContents(name, new Formula(content.Substring(1), Normalize, IsValid));
               
                try
                {
                    double temp;
                    NumberStyles styles = NumberStyles.AllowExponent | NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint;
                    temp= Double.Parse(content, styles);

                    return SetCellContents(name, temp);
                }
              
                catch
                {
                    return SetCellContents(name, content);
                }


            }
            throw new InvalidNameException();
        }
        
        /// <summary>
        ///   Returns the names of all non-empty cells.
        /// </summary>
        ///  
        /// <returns>
        ///     Returns an Enumerable that can be used to enumerate
        ///     the names of all the non-empty cells in the spreadsheet.  If 
        ///     all cells are empty then an IEnumerable with zero values will be returned.
        /// </returns>
        public override IEnumerable<string> GetNamesOfAllNonemptyCells()
        {
            foreach (var entry in Sheet)
            {
               
                    yield return entry.Key;
             
            }
        }
        /// <summary>
        ///   Returns the contents (as opposed to the value) of the named cell.
        /// </summary>
        /// 
        /// <exception cref="InvalidNameException"> 
        ///   Thrown if the name is invalid: blank/empty/""
        /// </exception>
        /// 
        /// <param name="name">The name of the spreadsheet cell to query</param>
        /// 
        /// <returns>
        ///   The return value should be either a string, a double, or a Formula.
        ///   See the class header summary 
        /// </returns>
        public override object GetCellContents(string name)
        {
            if ( String.Compare(name, null) == 0 || !(Normalize(name).IsVar()))
            {
                throw new InvalidNameException();
            }
            if (!Sheet.ContainsKey(name))
                return string.Empty;
            return Sheet[name].Contents;
          
            
            
        }
        
         /// <summary>
        ///  Set the contents of the named cell to the given number.  
        /// </summary>
        /// 
        /// <requires> 
        ///   The name parameter must be valid: non-empty/not ""
        /// </requires>
        /// 
        /// <exception cref="InvalidNameException"> 
        ///   If the name is invalid, throw an InvalidNameException
        /// </exception>
        /// 
        /// <param name="name"> The name of the cell </param>
        /// <param name="number"> The new contents/value </param>
        /// 
        /// <returns>
        ///   <para>
        ///       This method returns a LIST consisting of the passed in name followed by the names of all 
        ///       other cells whose value depends, directly or indirectly, on the named cell.
        ///   </para>
        ///
        ///   <para>
        ///       The order must correspond to a valid dependency ordering for recomputing
        ///       all of the cells, i.e., if you re-evaluate each cell in the order of the list,
        ///       the overall spreadsheet will be consistently updated.
        ///   </para>
        ///
        ///   
        /// </returns>
        protected override IList<string> SetCellContents(string name, double number)
        {
            Changed = true;
            if (name == null || !name.IsVar())
            {
                throw new InvalidNameException();
            }
            
            if (!Sheet.ContainsKey(name))
                Sheet.Add(name, new Cell(name, number, lookerupper));
               
            else
                Sheet[name].Contents = number;
            Changed = true;

            DG.ReplaceDependents(name, new HashSet<string>());
            foreach (var t in GetCellsToRecalculate(name))
            {
                Sheet[t].Contents = Sheet[t].Contents;
            }
            List<string> dents = new List<string>(GetCellsToRecalculate(name));
            
            return dents;
        }
        
         /// <summary>
        /// The contents of the named cell becomes the text.  
        /// </summary>
        /// 
        /// <requires> 
        ///   The name parameter must be valid/non-empty ""
        /// </requires>
        /// 
        /// <exception cref="InvalidNameException"> 
        ///   If the name is invalid, throw an InvalidNameException
        /// </exception>       
        /// 
        /// <param name="name"> The name of the cell </param>
        /// <param name="text"> The new content/value of the cell</param>
        /// 
        /// <returns>
        ///   <para>
        ///       This method returns a LIST consisting of the passed in name followed by the names of all 
        ///       other cells whose value depends, directly or indirectly, on the named cell.
        ///   </para>
        ///
        ///   <para>
        ///       The order must correspond to a valid dependency ordering for recomputing
        ///       all of the cells, i.e., if you re-evaluate each cell in the order of the list,
        ///       the overall spreadsheet will be consistently updated.
        ///   </para>
        ///
        ///   
        /// </returns>
        protected override IList<string> SetCellContents(string name, string text)
        {
            Changed = true;
            if (text == null)
            {
                throw new ArgumentNullException();
            }
            if (name == null || !name.IsVar())
            {
                throw new InvalidNameException();
            }

            if (!Sheet.ContainsKey(name))
                Sheet.Add(name, new Cell(name, text, lookerupper));
               
            else
                Sheet[name].Contents = text;
           
            DG.ReplaceDependents(name, new HashSet<string>());
            foreach (var t in GetCellsToRecalculate(name))
            {
                
                Sheet[t].Contents = Sheet[t].Contents;
            }
           
            List<string> dents = new List<string>(GetCellsToRecalculate(name));
            dents.Add(name);
            return dents;
        }
        
         /// <summary>
        /// Set the contents of the named cell to the formula.  
        /// </summary>
        /// 
        /// <requires> 
        ///   The name parameter must be valid/non empty
        /// </requires>
        /// 
        /// <exception cref="InvalidNameException"> 
        ///   If the name is invalid, throw an InvalidNameException
        /// </exception>
        /// 
        /// <exception cref="CircularException"> 
        ///   If changing the contents of the named cell to be the formula would 
        ///   cause a circular dependency, throw a CircularException.  
        ///   (NOTE: No change is made to the spreadsheet.)
        /// </exception>
        /// 
        /// <param name="name"> The cell name</param>
        /// <param name="formula"> The content of the cell</param>
        /// 
        /// <returns>
        ///   <para>
        ///       This method returns a LIST consisting of the passed in name followed by the names of all 
        ///       other cells whose value depends, directly or indirectly, on the named cell.
        ///   </para>
        ///
        ///   <para>
        ///       The order must correspond to a valid dependency ordering for recomputing
        ///       all of the cells, i.e., if you re-evaluate each cell in the order of the list,
        ///       the overall spreadsheet will be consistently updated.
        ///   </para>
        ///
        ///  
        ///   </para>
        /// </returns>
        protected override IList<string> SetCellContents(string name, Formula formula)
        {
            Changed = true;
            //If the formula parameter is null, throws an ArgumentNullExceptio
            if (formula == null)
            {
                throw new ArgumentNullException();
            }
            //If name is null or invalid, throws an InvalidNameException
            if (name == null || !name.IsVar())
            {
                throw new InvalidNameException();
            }
            IEnumerable<string> storedDents = DG.GetDependents(name);
            DG.ReplaceDependents(name, new HashSet<string>());
            foreach (string var in formula.GetVariables())
            {
                try
                {
                    DG.AddDependency(name, var);
                }
                catch (InvalidOperationException)
                {
                    DG.ReplaceDependents(name, storedDents);
                    throw new CircularException();
                }
            }
          
            if (!Sheet.ContainsKey(name))
                Sheet.Add(name, new Cell(name, formula, lookerupper));
                
            else
                Sheet[name].Contents = formula;
            foreach (string nombre in GetCellsToRecalculate(name))
            {
               
                Sheet[nombre].Contents = Sheet[nombre].Contents;
            }
           
            List<string> res = new List<string>(GetCellsToRecalculate(name));
            res.Add(name);
            return res;
        }
        
        /// <summary>
        /// Returns an enumeration, without duplicates, of the names of all cells whose
        /// values depend directly on the value of the named cell. 
        /// </summary>
        /// 
        /// <exception cref="InvalidNameException"> 
        ///   If the name is invalid, throw an InvalidNameException
        /// </exception>
        /// 
        /// <param name="name"></param>
        /// <returns>
        ///   Returns an enumeration, without duplicates, of the names of all cells that contain
        ///   formulas containing name.
        /// 
        ///   
        ///   <list type="bullet">
        ///      <item>A1 contains 3</item>
        ///      <item>B1 contains the formula A1 * A1</item>
        ///      <item>C1 contains the formula B1 + A1</item>
        ///      <item>D1 contains the formula B1 - C1</item>
        ///   </list>
        /// 
        ///   <para>The direct dependents of A1 are B1 and C1</para>
        /// 
        /// </returns>
        protected override IEnumerable<string> GetDirectDependents(string name)
        {
            if (name == null)
            {
                throw new ArgumentNullException();
            }
            else if (!name.IsVar())
            {
                throw new InvalidNameException();
            }
            return DG.GetDependees(name);
        }
        
        
        private class Cell
        {
            
            public String Name { get; private set; }
          
            private object _contents;
            public object Contents
            {
                get { return _contents; }
                set
                {
                    _value = value;
                    if (value is Formula)
                    {
                        _value = (_value as Formula).Evaluate(MyLookup);
                    }
                    _contents = value;
                }
            }
           
            private object _value;
            public object Value
            {
                get { return _value; }
                private set { _value = value; }
            }

            public Func<string, double> MyLookup { get; private set; }
        
            public Cell(string _name, object _contents, Func<string, double> _lookup)
            {
                Name = _name;
                MyLookup = _lookup;
                Contents = _contents;

            }

        }
        
	}
}

///helper class
namespace StringExtension
{
    //use to handle string 
    public static class StringExtension
    {

       

      
		
        public static bool IsVar1(this string s)
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

        public static bool IsOperator(this string s)
        {
            return (s == "+" || s == "-" || s == "*" || s == "/");
        }

       
    }
}

