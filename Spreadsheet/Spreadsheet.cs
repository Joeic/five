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


 namespace SS
{

	public class Spreadsheet : AbstractSpreadsheet
	{
        private DependencyGraph DG;
        private Dictionary<string, Cell> Sheet;
        
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
        
        public override object GetCellValue(string name)
        {

            if (name == null || !name.IsVar() || !IsValid(Normalize(name)))
                throw new InvalidNameException();
            if (Sheet.ContainsKey(name))
                return Sheet[name].Value;
            else
                return string.Empty;
            

        }
        
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
        
        
        public override IEnumerable<string> GetNamesOfAllNonemptyCells()
        {
            foreach (var entry in Sheet)
            {
               
                    yield return entry.Key;
             
            }
        }
        
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

